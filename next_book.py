#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Next Book — kişisel okuma önceliklendirme listesi (Windows masaüstü uygulaması)

* Open Library'den sadece kitap adı, yazar, tür ve sayfa sayısı çekilir.
* Kitap listede tek bir satır olur. Puanlanmamış satırlar soluk/italik,
  puanlanınca koyu görünür (Excel'deki gibi).
* Puanlama ölçütleri ve ağırlıkları kullanıcı tarafından serbestçe ayarlanır.
* Excel'den içe aktarma / Excel'e dışa aktarma (openpyxl kuruluysa).

Gereksinimler: Python 3.9+ (tkinter Windows kurulumuyla birlikte gelir)
İsteğe bağlı : pip install openpyxl   (Excel içe/dışa aktarma için)
Çalıştırma   : python next_book.py
"""

import ctypes
import json
import os
import queue
import re
import threading
import urllib.error
import urllib.parse
import urllib.request
import uuid
import warnings
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox, filedialog, simpledialog
import tkinter.font as tkfont

APP_NAME = "Next Book"
DATA_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "NextBook"
DATA_FILE = DATA_DIR / "next_book.json"

# Varsayılan ölçütler (Excel dosyasındaki Ayarlar sayfasıyla aynı)
DEFAULT_CRITERIA = [
    ("İlgi Düzeyi", 0.30),
    ("Kişisel Katkı", 0.30),
    ("Okuma Kolaylığı", 0.15),
    ("Sosyal Bağlam", 0.10),
    ("Uzun Vadeli / Referans Değeri", 0.15),
]
DEFAULT_SCALE_MAX = 5

# Satır renkleri
UNSCORED_FG = "#9AA3B2"   # puansız: soluk gri, italik
SCORED_FG = "#10233F"     # puanlı: koyu lacivert
SCORED_BG = "#DCE6F5"     # puanlı: hafif mavi zemin
HEADER_BG = "#1F3A5F"


# ============================================================================
# VERİ MODELİ
# ============================================================================

def new_id():
    return uuid.uuid4().hex[:8]


def default_data():
    return {
        "scale_max": DEFAULT_SCALE_MAX,
        "criteria": [{"id": new_id(), "name": n, "weight": w} for n, w in DEFAULT_CRITERIA],
        "books": [],
    }


def new_book(title, author="", genre="", pages=None, key=""):
    return {
        "id": new_id(),
        "key": key,          # Open Library work key (tekrar eklemeyi önlemek için)
        "title": title,
        "author": author,
        "genre": genre,
        "pages": pages,
        "scores": {},        # {ölçüt_id: puan}
    }


def compute_score(book, criteria):
    """Excel'deki formül: boş bırakılan ölçütler hesaba katılmaz.
    skor = Σ(puan × ağırlık) / Σ(dolu ölçütlerin ağırlığı)"""
    num = den = 0.0
    for c in criteria:
        v = book["scores"].get(c["id"])
        if v is None or v == "":
            continue
        num += float(v) * float(c["weight"])
        den += float(c["weight"])
    return round(num / den, 2) if den > 0 else 0.0


def is_scored(book, criteria):
    return any(book["scores"].get(c["id"]) not in (None, "") for c in criteria)


def fmt_num(v):
    if v is None or v == "":
        return ""
    v = float(v)
    return str(int(v)) if v == int(v) else f"{v:.1f}".rstrip("0").rstrip(".")


def parse_float(text):
    return float(text.strip().replace(",", "."))


def load_data():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        base = default_data()
        data.setdefault("scale_max", base["scale_max"])
        data.setdefault("criteria", base["criteria"])
        data.setdefault("books", [])
        for b in data["books"]:
            b.setdefault("scores", {})
            b.setdefault("key", "")
        return data
    except FileNotFoundError:
        return default_data()
    except Exception:
        # Bozuk dosya: yedekle ve temiz başla
        try:
            DATA_FILE.replace(DATA_FILE.with_suffix(".bozuk.json"))
        except Exception:
            pass
        return default_data()


def save_data(data):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    tmp = DATA_FILE.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    tmp.replace(DATA_FILE)


# ============================================================================
# OPEN LIBRARY ARAMA (anahtarsız) — sadece ad, yazar, tür, sayfa sayısı
# ============================================================================

_JUNK_SUBJECT_PARTS = (
    "accessible", "daisy", "in library", "lending", "overdrive", "open library",
    "large type", "reading level", "nyt:", "new york times", "internet archive",
    "protected", "bestseller", "award", "collection", "translations",
)
_PREFERRED_GENRES = (
    "science fiction", "fantasy", "fiction", "biography", "history", "philosophy",
    "psychology", "business", "economics", "self-help", "poetry", "mystery",
    "thriller", "romance", "horror", "science", "politics", "religion",
    "computers", "art", "health", "education", "mathematics", "sociology",
    "classic literature", "short stories", "essays", "drama", "humor",
)


def pick_genre(subjects):
    """Open Library konu listesi gürültülü; makul bir tür seç."""
    clean = []
    for s in subjects or []:
        t = (s or "").strip()
        low = t.lower()
        if not t or len(t) > 30 or any(j in low for j in _JUNK_SUBJECT_PARTS):
            continue
        clean.append(t)
    lows = {t.lower(): t for t in clean}
    for g in _PREFERRED_GENRES:
        if g in lows:
            return lows[g]
    return clean[0] if clean else ""


def search_books(title, author, limit=25):
    params = {
        "limit": limit,
        "fields": "key,title,author_name,subject,number_of_pages_median,first_publish_year",
    }
    if title:
        params["title"] = title
    if author:
        params["author"] = author
    url = "https://openlibrary.org/search.json?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": "NextBook/1.0 (personal reading list)"})
    with urllib.request.urlopen(req, timeout=15) as resp:
        payload = json.load(resp)
    return parse_search_docs(payload.get("docs", []))


def parse_search_docs(docs):
    results = []
    for d in docs:
        pages = d.get("number_of_pages_median")
        results.append({
            "key": d.get("key", ""),
            "title": d.get("title", "").strip(),
            "author": ", ".join((d.get("author_name") or [])[:3]),
            "genre": pick_genre(d.get("subject")),
            "pages": int(pages) if pages else None,
            "year": d.get("first_publish_year") or "",
        })
    return [r for r in results if r["title"]]


# ============================================================================
# EXCEL İÇE / DIŞA AKTARMA (openpyxl gerekir)
# ============================================================================

def _require_openpyxl():
    try:
        import openpyxl  # noqa: F401
        return True
    except ImportError:
        messagebox.showerror(
            APP_NAME,
            "Excel desteği için openpyxl gerekli.\n\nKomut satırında şunu çalıştırın:\n\npip install openpyxl",
        )
        return False


_SCALE_SUFFIX = re.compile(r"\s*\(\s*\d+\s*-\s*(\d+)\s*\)\s*$")


def import_excel(path):
    """Excel dosyasından yeni bir veri sözlüğü üretir.
    Beklenen düzen: Kitap Adı | Yazar | Tür | Sayfa Sayısı | ölçüt sütunları… | Öncelik Skoru
    Ağırlıklar 'Ayarlar' sayfasından (A: ölçüt adı, B: ağırlık) okunur."""
    from openpyxl import load_workbook
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    header = [str(c.value or "").strip() for c in ws[1]]
    low = [h.lower() for h in header]

    def find(*aliases):
        for a in aliases:
            if a in low:
                return low.index(a)
        return None

    ti = find("kitap adı", "kitap", "başlık", "title")
    ai = find("yazar", "author")
    gi = find("tür", "kategori", "genre")
    pi = find("sayfa sayısı", "sayfa", "pages")
    si = next((i for i, h in enumerate(low) if h.startswith("öncelik skoru") or h == "skor"), None)
    if ti is None:
        raise ValueError("'Kitap Adı' sütunu bulunamadı.")

    known = {i for i in (ti, ai, gi, pi, si) if i is not None}
    crit_cols = [i for i, h in enumerate(header) if h and i not in known]
    if not crit_cols:
        raise ValueError("Puanlama ölçütü sütunu bulunamadı.")

    # Ayarlar sayfasından ağırlıklar
    weights = {}
    if "Ayarlar" in wb.sheetnames:
        for row in wb["Ayarlar"].iter_rows(min_row=2, max_col=2, values_only=True):
            if row[0] and isinstance(row[1], (int, float)):
                weights[str(row[0]).strip().lower()] = float(row[1])

    scale_max = DEFAULT_SCALE_MAX
    criteria = []
    for i in crit_cols:
        raw = header[i]
        m = _SCALE_SUFFIX.search(raw)
        if m:
            scale_max = int(m.group(1))
        name = _SCALE_SUFFIX.sub("", raw).strip()
        criteria.append({"id": new_id(), "name": name, "col": i,
                         "weight": weights.get(name.lower(), 1.0)})

    books = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        title = row[ti] if ti < len(row) else None
        if not title or not str(title).strip():
            continue

        def cell(idx):
            return row[idx] if idx is not None and idx < len(row) else None

        pages = cell(pi)
        try:
            pages = int(float(pages)) if pages not in (None, "") else None
        except (TypeError, ValueError):
            pages = None
        b = new_book(str(title).strip(),
                     str(cell(ai) or "").strip(),
                     str(cell(gi) or "").strip(),
                     pages)
        for c in criteria:
            v = cell(c["col"])
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                b["scores"][c["id"]] = float(v) if v != int(v) else int(v)
        books.append(b)

    for c in criteria:
        c.pop("col")
    return {"scale_max": scale_max, "criteria": criteria, "books": books}


def export_excel(data, path):
    """Orijinal Excel düzenine benzer çıktı: formüller, gri/italik puansız satırlar,
    renk skalası ve veri çubuğu ile."""
    from openpyxl import Workbook
    from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as L

    cs, books, smax = data["criteria"], data["books"], data["scale_max"]
    n = len(cs)
    score_col = 5 + n
    wb = Workbook()
    ws = wb.active
    ws.title = "Okuma Önceliklendirme"
    st = wb.create_sheet("Ayarlar")

    head_fill = PatternFill("solid", fgColor=HEADER_BG.lstrip("#"))
    head_font = Font(bold=True, color="FFFFFF")
    headers = ["Kitap Adı", "Yazar", "Tür", "Sayfa Sayısı"] + \
              [f"{c['name']} (1-{smax})" for c in cs] + ["Öncelik Skoru"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.fill, cell.font = head_fill, head_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[1].height = 45

    st.cell(row=1, column=1, value="Ağırlıklar").fill = head_fill
    st.cell(row=1, column=2, value="Ağırlık").fill = head_fill
    st["A1"].font = st["B1"].font = head_font
    for i, c in enumerate(cs):
        st.cell(row=2 + i, column=1, value=c["name"])
        wc = st.cell(row=2 + i, column=2, value=c["weight"])
        wc.fill = PatternFill("solid", fgColor="FFF6DC")
        wc.font = Font(bold=True)
    st.cell(row=2 + n, column=1, value="Toplam").font = Font(bold=True)
    st.cell(row=2 + n, column=2, value=f"=SUM(B2:B{1 + n})").font = Font(bold=True)
    st.cell(row=2 + n, column=3, value="Ağırlıklar birbirine oranlanır; toplamın 1 olması şart değil.")
    st.column_dimensions["A"].width = 36
    st.column_dimensions["C"].width = 60

    for r, b in enumerate(books, start=2):
        ws.cell(row=r, column=1, value=b["title"]).font = Font(bold=True)
        ws.cell(row=r, column=2, value=b["author"])
        ws.cell(row=r, column=3, value=b["genre"])
        ws.cell(row=r, column=4, value=b["pages"])
        for i, c in enumerate(cs):
            v = b["scores"].get(c["id"])
            if v not in (None, ""):
                ws.cell(row=r, column=5 + i, value=v)
        num = "+".join(f'IF({L(5+i)}{r}="",0,{L(5+i)}{r}*Ayarlar!$B${2+i})' for i in range(n))
        den = "+".join(f'IF({L(5+i)}{r}="",0,Ayarlar!$B${2+i})' for i in range(n))
        f = ws.cell(row=r, column=score_col, value=f"=IF(({den})=0,0,ROUND(({num})/({den}),2))")
        f.font = Font(bold=True)

    last = len(books) + 101   # sonradan Excel'de eklenecek satırlar için pay
    sc = L(score_col)
    ws.conditional_formatting.add(
        f"A2:{sc}{last}",
        FormulaRule(formula=[f'AND($A2<>"",${sc}2=0)'], font=Font(italic=True, color="8A94A6")))
    ws.conditional_formatting.add(
        f"E2:{L(4+n)}{last}",
        ColorScaleRule(start_type="num", start_value=1, start_color="FBE5E1",
                       mid_type="num", mid_value=(1 + smax) / 2, mid_color="FFF6DC",
                       end_type="num", end_value=smax, end_color="CDEBD6"))
    ws.conditional_formatting.add(
        f"{sc}2:{sc}{last}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=smax, color="7FA6D9"))

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 17
    ws.column_dimensions["D"].width = 14
    for i in range(n):
        ws.column_dimensions[L(5 + i)].width = 15
    ws.column_dimensions[sc].width = 14
    ws.freeze_panes = "B2"
    wb.save(path)


# ============================================================================
# PUANLAMA AYARLARI PENCERESİ
# ============================================================================

class SettingsDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("Puanlama ayarları")
        self.transient(app.root)
        self.resizable(False, False)
        self.rows = []   # {"id","name","weight","frame"}

        pad = {"padx": 10, "pady": 4}
        ttk.Label(self, text="Ölçütlerini ve ağırlıklarını kendin belirle.",
                  font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", **pad)
        ttk.Label(self, foreground="#555", justify="left",
                  text="Ağırlığı yüksek ölçüt, nihai skoru daha çok etkiler. Ağırlıklar birbirine\n"
                       "oranlanır; toplamın 1 olması gerekmez. Boş bırakılan puan hesaba katılmaz.",
                  ).grid(row=1, column=0, sticky="w", **pad)

        self.body = ttk.Frame(self)
        self.body.grid(row=2, column=0, sticky="ew", padx=10)
        ttk.Label(self.body, text="Ölçüt adı").grid(row=0, column=0, sticky="w")
        ttk.Label(self.body, text="Ağırlık").grid(row=0, column=1, sticky="w", padx=6)

        for c in app.data["criteria"]:
            self._add_row(c["id"], c["name"], c["weight"])

        ctl = ttk.Frame(self)
        ctl.grid(row=3, column=0, sticky="ew", padx=10, pady=(6, 0))
        ttk.Button(ctl, text="+ Ölçüt ekle", command=lambda: self._add_row(new_id(), "", 1.0)).pack(side="left")
        ttk.Label(ctl, text="   Puan üst sınırı:").pack(side="left")
        self.scale_var = tk.StringVar(value=str(app.data["scale_max"]))
        ttk.Spinbox(ctl, from_=2, to=10, width=4, textvariable=self.scale_var).pack(side="left", padx=4)

        self.total_lbl = ttk.Label(self, foreground="#555")
        self.total_lbl.grid(row=4, column=0, sticky="w", padx=10, pady=(6, 0))

        btns = ttk.Frame(self)
        btns.grid(row=5, column=0, sticky="e", padx=10, pady=10)
        ttk.Button(btns, text="İptal", command=self.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Kaydet", command=self.save).pack(side="right")

        self._update_total()
        self.grab_set()
        self.focus_set()

    def _add_row(self, cid, name, weight):
        r = len(self.rows) + 1
        name_var = tk.StringVar(value=name)
        weight_var = tk.StringVar(value=fmt_num(weight) if weight != "" else "")
        weight_var.trace_add("write", lambda *_: self._update_total())
        e1 = ttk.Entry(self.body, textvariable=name_var, width=34)
        e2 = ttk.Entry(self.body, textvariable=weight_var, width=8)
        row = {"id": cid, "name": name_var, "weight": weight_var, "widgets": [e1, e2]}
        btn = ttk.Button(self.body, text="✕", width=3, command=lambda: self._remove_row(row))
        row["widgets"].append(btn)
        e1.grid(row=r, column=0, pady=2)
        e2.grid(row=r, column=1, padx=6, pady=2)
        btn.grid(row=r, column=2)
        self.rows.append(row)
        self._regrid()
        self._update_total()
        if not name:
            e1.focus_set()

    def _remove_row(self, row):
        if len(self.rows) <= 1:
            messagebox.showinfo(APP_NAME, "En az bir ölçüt olmalı.", parent=self)
            return
        for w in row["widgets"]:
            w.destroy()
        self.rows.remove(row)
        self._regrid()
        self._update_total()

    def _regrid(self):
        for i, row in enumerate(self.rows, start=1):
            e1, e2, btn = row["widgets"]
            e1.grid(row=i, column=0, pady=2)
            e2.grid(row=i, column=1, padx=6, pady=2)
            btn.grid(row=i, column=2)

    def _update_total(self):
        total = 0.0
        for row in self.rows:
            try:
                total += parse_float(row["weight"].get())
            except ValueError:
                pass
        self.total_lbl.config(text=f"Ağırlık toplamı: {total:.2f}")

    def save(self):
        criteria = []
        for row in self.rows:
            name = row["name"].get().strip()
            if not name:
                messagebox.showwarning(APP_NAME, "Tüm ölçütlerin bir adı olmalı.", parent=self)
                return
            try:
                w = parse_float(row["weight"].get())
                if w < 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning(APP_NAME, f"“{name}” için ağırlık 0 veya daha büyük bir sayı olmalı.", parent=self)
                return
            criteria.append({"id": row["id"], "name": name, "weight": w})
        if len({c["name"].lower() for c in criteria}) != len(criteria):
            messagebox.showwarning(APP_NAME, "İki ölçüt aynı ada sahip olamaz.", parent=self)
            return
        try:
            smax = int(self.scale_var.get())
            if not 2 <= smax <= 10:
                raise ValueError
        except ValueError:
            messagebox.showwarning(APP_NAME, "Puan üst sınırı 2 ile 10 arasında olmalı.", parent=self)
            return
        self.app.apply_settings(criteria, smax)
        self.destroy()


# ============================================================================
# ANA UYGULAMA
# ============================================================================

class NextBookApp:
    FIXED = ["title", "author", "genre", "pages"]
    FIXED_LABELS = {"title": "Kitap Adı", "author": "Yazar", "genre": "Tür", "pages": "Sayfa"}

    def __init__(self, root):
        self.root = root
        self.data = load_data()
        self.results = []
        self.q = queue.Queue()
        self.edit = None
        self.sort_col = None
        self.sort_rev = False

        root.title(APP_NAME)
        root.geometry("1200x740")
        root.minsize(900, 560)
        self._setup_style()
        self._build_ui()
        self.rebuild_columns()
        self.refresh_table()
        self._poll_queue()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------------------------------------------------------- stil
    def _setup_style(self):
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(name).configure(family="Segoe UI", size=10)
            except tk.TclError:
                pass
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Treeview", rowheight=28, borderwidth=0)
        style.configure("Treeview.Heading", background=HEADER_BG, foreground="white",
                        font=("Segoe UI", 10, "bold"), relief="flat", padding=6)
        style.map("Treeview.Heading", background=[("active", "#2B4F80")])
        style.map("Treeview", background=[("selected", "#F0C060")],
                  foreground=[("selected", "#000000")])
        style.configure("Accent.TButton", background="#C8962E", foreground="#1b1b1b",
                        font=("Segoe UI", 10, "bold"), padding=(14, 4))
        style.map("Accent.TButton", background=[("active", "#DDAA3C"), ("disabled", "#d9d2c0")])

        self.font_normal = tkfont.Font(family="Segoe UI", size=10)
        self.font_italic = tkfont.Font(family="Segoe UI", size=10, slant="italic")

    # ---------------------------------------------------------------- arayüz
    def _build_ui(self):
        r = self.root
        top = ttk.Frame(r, padding=(14, 10, 14, 4))
        top.pack(fill="x")
        ttk.Label(top, text="Next Book", font=("Segoe UI", 18, "bold")).pack(side="left")
        ttk.Label(top, text="   Ara → listeye ekle → puanla → okuma sırası belli olsun",
                  foreground="#666").pack(side="left", pady=(8, 0))

        # Arama
        sf = ttk.Frame(r, padding=(14, 4))
        sf.pack(fill="x")
        self.title_var = tk.StringVar()
        self.author_var = tk.StringVar()
        ttk.Label(sf, text="Kitap adı").grid(row=0, column=0, sticky="w")
        ttk.Label(sf, text="Yazar (isteğe bağlı)").grid(row=0, column=1, sticky="w", padx=(8, 0))
        e1 = ttk.Entry(sf, textvariable=self.title_var, width=46)
        e2 = ttk.Entry(sf, textvariable=self.author_var, width=30)
        e1.grid(row=1, column=0, sticky="ew")
        e2.grid(row=1, column=1, sticky="ew", padx=(8, 0))
        self.search_btn = ttk.Button(sf, text="Ara", style="Accent.TButton", command=self.do_search)
        self.search_btn.grid(row=1, column=2, padx=(8, 0))
        ttk.Button(sf, text="Temizle", command=self.clear_search).grid(row=1, column=3, padx=(6, 0))
        ttk.Button(sf, text="+ Elle ekle", command=self.add_manual).grid(row=1, column=4, padx=(14, 0))
        sf.columnconfigure(0, weight=3)
        sf.columnconfigure(1, weight=2)
        for e in (e1, e2):
            e.bind("<Return>", lambda _e: self.do_search())
        e1.focus_set()

        # Sonuçlar
        rf = ttk.LabelFrame(r, text=" Arama sonuçları — eklemek için çift tıkla ", padding=6)
        rf.pack(fill="x", padx=14, pady=(8, 4))
        cols = ("title", "author", "genre", "pages", "year")
        self.res_tree = ttk.Treeview(rf, columns=cols, show="headings", height=5, selectmode="browse")
        for cid, text, w in (("title", "Kitap Adı", 380), ("author", "Yazar", 240),
                             ("genre", "Tür", 150), ("pages", "Sayfa", 70), ("year", "İlk yayın", 80)):
            self.res_tree.heading(cid, text=text)
            self.res_tree.column(cid, width=w, stretch=(cid == "title"),
                                 anchor="center" if cid in ("pages", "year") else "w")
        rs = ttk.Scrollbar(rf, orient="vertical", command=self.res_tree.yview)
        self.res_tree.configure(yscrollcommand=rs.set)
        self.res_tree.pack(side="left", fill="x", expand=True)
        rs.pack(side="left", fill="y")
        ttk.Button(rf, text="Listeye ekle", command=self.add_selected_result).pack(side="left", padx=(8, 0), anchor="n")
        self.res_tree.tag_configure("added", foreground=UNSCORED_FG)
        self.res_tree.bind("<Double-1>", lambda _e: self.add_selected_result())

        # Ana liste
        mf = ttk.LabelFrame(r, text=" Okuma listem ", padding=6)
        mf.pack(fill="both", expand=True, padx=14, pady=(4, 4))
        self.tree = ttk.Treeview(mf, show="headings", selectmode="extended")
        vs = ttk.Scrollbar(mf, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(mf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        mf.rowconfigure(0, weight=1)
        mf.columnconfigure(0, weight=1)
        self.tree.tag_configure("unscored", foreground=UNSCORED_FG, font=self.font_italic)
        self.tree.tag_configure("scored", foreground=SCORED_FG, background=SCORED_BG, font=self.font_normal)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Delete>", lambda _e: self.delete_selected())
        self.tree.bind("<MouseWheel>", lambda _e: self.finish_edit(True))

        # Alt çubuk
        bf = ttk.Frame(r, padding=(14, 4, 14, 10))
        bf.pack(fill="x")
        ttk.Button(bf, text="⚙ Puanlama ayarları", command=self.open_settings).pack(side="left")
        ttk.Button(bf, text="Puana göre sırala", command=self.sort_by_score).pack(side="left", padx=6)
        ttk.Button(bf, text="Seçileni sil", command=self.delete_selected).pack(side="left")
        ttk.Button(bf, text="Excel'den içe aktar", command=self.import_from_excel).pack(side="left", padx=(18, 0))
        ttk.Button(bf, text="Excel'e aktar", command=self.export_to_excel).pack(side="left", padx=6)
        self.status = ttk.Label(bf, text="", foreground="#555")
        self.status.pack(side="right")
        ttk.Label(r, foreground="#888", padding=(14, 0, 14, 8),
                  text="İpucu: Hücreye çift tıkla → düzenle. Puan hücresinde Enter = aşağı, Tab = sağa, "
                       "Esc = vazgeç, boş bırak = puanı sil. Başlığa tıklayarak sırala."
                  ).pack(fill="x")

    # ---------------------------------------------------------------- sütunlar
    def col_ids(self):
        return self.FIXED + ["c_" + c["id"] for c in self.data["criteria"]] + ["score"]

    def rebuild_columns(self):
        ids = self.col_ids()
        self.tree["columns"] = ids
        self.tree["displaycolumns"] = ids
        self.update_headings()
        widths = {"title": 330, "author": 200, "genre": 130, "pages": 70}
        for cid in ids:
            if cid in widths:
                self.tree.column(cid, width=widths[cid], stretch=(cid == "title"),
                                 anchor="center" if cid == "pages" else "w", minwidth=50)
            elif cid == "score":
                self.tree.column(cid, width=110, stretch=False, anchor="center", minwidth=80)
            else:
                name = next(c["name"] for c in self.data["criteria"] if "c_" + c["id"] == cid)
                self.tree.column(cid, width=min(max(len(name) * 8 + 36, 100), 230),
                                 stretch=False, anchor="center", minwidth=70)

    def update_headings(self):
        for cid in self.col_ids():
            if cid in self.FIXED_LABELS:
                text = self.FIXED_LABELS[cid]
            elif cid == "score":
                text = "Öncelik Skoru"
            else:
                c = next(c for c in self.data["criteria"] if "c_" + c["id"] == cid)
                text = f"{c['name']} ({c_range(self.data)})"
            if cid == self.sort_col:
                text += " ▼" if self.sort_rev else " ▲"
            self.tree.heading(cid, text=text, command=lambda c=cid: self.sort_by(c))

    # ---------------------------------------------------------------- tablo
    def row_values(self, b):
        cs = self.data["criteria"]
        vals = [b["title"], b["author"], b["genre"], "" if b["pages"] in (None, "") else b["pages"]]
        vals += [fmt_num(b["scores"].get(c["id"])) for c in cs]
        vals.append(f"{compute_score(b, cs):.2f}" if is_scored(b, cs) else "—")
        return vals

    def row_tag(self, b):
        return ("scored",) if is_scored(b, self.data["criteria"]) else ("unscored",)

    def refresh_table(self):
        self.finish_edit(False)
        self.tree.delete(*self.tree.get_children())
        for b in self.data["books"]:
            self.tree.insert("", "end", iid=b["id"], values=self.row_values(b), tags=self.row_tag(b))
        self.update_status()

    def update_row(self, b):
        self.tree.item(b["id"], values=self.row_values(b), tags=self.row_tag(b))
        self.update_status()

    def update_status(self):
        cs = self.data["criteria"]
        total = len(self.data["books"])
        scored = sum(1 for b in self.data["books"] if is_scored(b, cs))
        self.set_status(f"{total} kitap · {scored} puanlandı · {total - scored} puan bekliyor")

    def set_status(self, text):
        self.status.config(text=text)

    def book_by_id(self, bid):
        return next((b for b in self.data["books"] if b["id"] == bid), None)

    def save(self):
        try:
            save_data(self.data)
        except OSError as e:
            messagebox.showerror(APP_NAME, f"Kaydedilemedi:\n{e}")

    # ---------------------------------------------------------------- sıralama
    def sort_by(self, cid):
        if self.sort_col == cid:
            self.sort_rev = not self.sort_rev
        else:
            self.sort_col = cid
            self.sort_rev = cid == "score" or cid.startswith("c_")   # puanlar: yüksekten düşüğe
        self.apply_sort()

    def sort_by_score(self):
        self.sort_col, self.sort_rev = "score", True
        self.apply_sort()

    def apply_sort(self):
        self.finish_edit(True)
        cs = self.data["criteria"]
        col = self.sort_col

        def value(b):
            if col == "score":
                return compute_score(b, cs) if is_scored(b, cs) else None
            if col == "pages":
                return b["pages"] if b["pages"] not in (None, "") else None
            if col.startswith("c_"):
                v = b["scores"].get(col[2:])
                return None if v in (None, "") else float(v)
            t = (b.get(col) or "").strip()
            return t.casefold() if t else None

        present = [b for b in self.data["books"] if value(b) is not None]
        missing = [b for b in self.data["books"] if value(b) is None]   # boşlar hep sonda
        present.sort(key=value, reverse=self.sort_rev)
        self.data["books"] = present + missing
        self.save()
        self.update_headings()
        self.refresh_table()

    # ---------------------------------------------------------------- arama
    def do_search(self):
        t, a = self.title_var.get().strip(), self.author_var.get().strip()
        if not t and not a:
            self.set_status("Bir kitap adı veya yazar yazın.")
            return
        self.set_status("Aranıyor…")
        self.search_btn.config(state="disabled")
        threading.Thread(target=self._search_worker, args=(t, a), daemon=True).start()

    def _search_worker(self, t, a):
        try:
            self.q.put(("ok", search_books(t, a)))
        except Exception as e:   # ağ hataları dahil
            self.q.put(("err", e))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                self.search_btn.config(state="normal")
                if kind == "ok":
                    self.show_results(payload)
                else:
                    if isinstance(payload, (urllib.error.URLError, TimeoutError, OSError)):
                        msg = "Open Library'ye ulaşılamadı. İnternet bağlantınızı kontrol edin."
                    else:
                        msg = f"Arama hatası: {payload}"
                    self.set_status(msg)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def show_results(self, results):
        self.results = results
        self.res_tree.delete(*self.res_tree.get_children())
        for i, r in enumerate(results):
            added = self.find_duplicate(r) is not None
            self.res_tree.insert("", "end", iid=str(i), tags=("added",) if added else (),
                                 values=(("✓ " if added else "") + r["title"], r["author"], r["genre"],
                                         r["pages"] or "", r["year"]))
        self.update_status()
        if not results:
            self.set_status("Sonuç bulunamadı. Yazımı değiştirmeyi ya da “+ Elle ekle”yi deneyin.")
        else:
            self.set_status(f"{len(results)} sonuç bulundu. Eklemek için çift tıklayın.")

    def clear_search(self):
        self.title_var.set("")
        self.author_var.set("")
        self.results = []
        self.res_tree.delete(*self.res_tree.get_children())
        self.update_status()

    # ---------------------------------------------------------------- ekleme / silme
    def find_duplicate(self, r):
        tkey = (r["title"].casefold(), (r["author"] or "").casefold())
        for b in self.data["books"]:
            if r.get("key") and b.get("key") == r["key"]:
                return b
            if (b["title"].casefold(), (b["author"] or "").casefold()) == tkey:
                return b
        return None

    def add_selected_result(self):
        sel = self.res_tree.selection()
        if not sel:
            self.set_status("Önce bir sonuç seçin.")
            return
        r = self.results[int(sel[0])]
        dup = self.find_duplicate(r)
        if dup:
            self.tree.selection_set(dup["id"])
            self.tree.see(dup["id"])
            messagebox.showinfo(APP_NAME, f"“{dup['title']}” zaten listenizde.")
            return
        self._append_book(new_book(r["title"], r["author"], r["genre"], r["pages"], r["key"]))
        self.res_tree.item(sel[0], tags=("added",))

    def add_manual(self):
        title = simpledialog.askstring(APP_NAME, "Kitap adı:", parent=self.root)
        if title and title.strip():
            self._append_book(new_book(title.strip()))

    def _append_book(self, b):
        self.data["books"].append(b)
        self.save()
        self.tree.insert("", "end", iid=b["id"], values=self.row_values(b), tags=self.row_tag(b))
        self.tree.selection_set(b["id"])
        self.tree.see(b["id"])
        self.update_status()

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        if not messagebox.askyesno(APP_NAME, f"Seçili {len(sel)} kitap listeden silinsin mi?"):
            return
        self.finish_edit(False)
        ids = set(sel)
        self.data["books"] = [b for b in self.data["books"] if b["id"] not in ids]
        self.save()
        self.tree.delete(*sel)
        self.update_status()

    # ---------------------------------------------------------------- hücre düzenleme
    def on_double_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        idx = int(self.tree.identify_column(event.x)[1:]) - 1
        if iid:
            self.start_edit(iid, idx)

    def start_edit(self, iid, idx):
        self.finish_edit(True)
        ids = self.col_ids()
        if idx < 0 or idx >= len(ids) or ids[idx] == "score":
            return
        cid = ids[idx]
        self.tree.see(iid)
        self.tree.update_idletasks()
        bbox = self.tree.bbox(iid, f"#{idx + 1}")
        if not bbox:
            return
        x, y, w, h = bbox
        b = self.book_by_id(iid)
        if cid.startswith("c_"):
            cur = fmt_num(b["scores"].get(cid[2:]))
        else:
            cur = "" if b[cid] in (None, "") else str(b[cid])
        entry = tk.Entry(self.tree, font=("Segoe UI", 10), relief="solid", bd=1,
                         justify="center" if cid.startswith("c_") or cid == "pages" else "left")
        entry.insert(0, cur)
        entry.select_range(0, "end")
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        self.edit = {"iid": iid, "idx": idx, "cid": cid, "entry": entry}
        entry.bind("<Return>", lambda _e: self._commit_and_move(1, 0))
        entry.bind("<Tab>", lambda _e: self._commit_and_move(0, 1))
        entry.bind("<Escape>", lambda _e: self.finish_edit(False))
        entry.bind("<FocusOut>", lambda _e: self.finish_edit(True))

    def _commit_and_move(self, d_row, d_col):
        e = self.edit
        if not e:
            return "break"
        iid, idx = e["iid"], e["idx"]
        self.finish_edit(True)
        ids = self.col_ids()
        nidx = idx + d_col
        niid = iid
        if d_row:
            niid = self.tree.next(iid)
        elif nidx >= len(ids) - 1:   # skor sütunu düzenlenemez → alt satırın ilk ölçütüne geç
            niid = self.tree.next(iid)
            nidx = len(self.FIXED)
        if niid and nidx >= 0:
            self.tree.selection_set(niid)
            self.root.after(10, lambda: self.start_edit(niid, nidx))
        return "break"

    def finish_edit(self, commit=True):
        e = self.edit
        if not e:
            return
        self.edit = None
        text = e["entry"].get().strip()
        try:
            e["entry"].destroy()
        except tk.TclError:
            pass
        if commit:
            self.apply_edit(e["iid"], e["cid"], text)

    def apply_edit(self, iid, cid, text):
        b = self.book_by_id(iid)
        if not b:
            return
        if cid == "title":
            if not text:
                return
            b["title"] = text
        elif cid in ("author", "genre"):
            b[cid] = text
        elif cid == "pages":
            if text == "":
                b["pages"] = None
            else:
                try:
                    b["pages"] = int(float(text.replace(",", ".")))
                except ValueError:
                    self.root.bell()
                    self.set_status("Sayfa sayısı bir sayı olmalı.")
                    return
        else:   # puan
            key = cid[2:]
            if text == "":
                b["scores"].pop(key, None)
            else:
                smax = self.data["scale_max"]
                try:
                    v = parse_float(text)
                    if not 1 <= v <= smax:
                        raise ValueError
                except ValueError:
                    self.root.bell()
                    self.set_status(f"Geçersiz puan: 1 ile {smax} arasında bir sayı girin.")
                    return
                b["scores"][key] = int(v) if v == int(v) else round(v, 1)
        self.save()
        self.update_row(b)

    # ---------------------------------------------------------------- ayarlar
    def open_settings(self):
        self.finish_edit(True)
        SettingsDialog(self)

    def apply_settings(self, criteria, scale_max):
        alive = {c["id"] for c in criteria}
        for b in self.data["books"]:
            b["scores"] = {k: min(v, scale_max) for k, v in b["scores"].items() if k in alive}
        self.data["criteria"] = criteria
        self.data["scale_max"] = scale_max
        if self.sort_col and self.sort_col.startswith("c_") and self.sort_col[2:] not in alive:
            self.sort_col = None
        self.save()
        self.rebuild_columns()
        self.refresh_table()

    # ---------------------------------------------------------------- Excel
    def import_from_excel(self):
        if not _require_openpyxl():
            return
        path = filedialog.askopenfilename(title="Excel dosyası seç",
                                          filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tümü", "*.*")])
        if not path:
            return
        if self.data["books"] and not messagebox.askyesno(
                APP_NAME, "Mevcut listeniz silinip Excel dosyasındaki kitaplar ve ölçütler yüklenecek.\n"
                          "(Önce “Excel'e aktar” ile yedek almak isteyebilirsiniz.)\n\nDevam edilsin mi?"):
            return
        try:
            new = import_excel(path)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Excel okunamadı:\n{e}")
            return
        self.data = new
        self.sort_col = None
        self.save()
        self.rebuild_columns()
        self.refresh_table()
        messagebox.showinfo(APP_NAME, f"{len(new['books'])} kitap ve {len(new['criteria'])} ölçüt içe aktarıldı.")

    def export_to_excel(self):
        if not _require_openpyxl():
            return
        path = filedialog.asksaveasfilename(title="Excel olarak kaydet", defaultextension=".xlsx",
                                            initialfile="Next_Book.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            export_excel(self.data, path)
        except PermissionError:
            messagebox.showerror(APP_NAME, "Dosya yazılamadı. Excel'de açıksa kapatıp tekrar deneyin.")
            return
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Dışa aktarılamadı:\n{e}")
            return
        self.set_status(f"Kaydedildi: {path}")

    # ---------------------------------------------------------------- kapanış
    def on_close(self):
        self.finish_edit(True)
        self.save()
        self.root.destroy()


def c_range(data):
    return f"1-{data['scale_max']}"


def main():
    try:   # Windows'ta net yazı için
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    NextBookApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
