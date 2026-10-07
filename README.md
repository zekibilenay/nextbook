# Next Book

**TR** · Okuyacağın kitapları kendi ölçütlerine göre puanlayıp kişisel bir okuma sırası oluşturan, ücretsiz ve açık kaynaklı Windows, macOS ve Linux uygulaması.
**EN** · A free, open-source Windows, macOS and Linux app that ranks the books you want to read using scoring criteria you define yourself.

## Özellikler / Features
- 🌍 **6 dil / 6 languages:** Türkçe, English, Русский, Deutsch, Français, 中文 (ilk açılışta seçilir / chosen on first launch)
- 🖥️ Tam ekran boyutunda açılır · Opens maximised
- 🎨 **6 tema (3 aydınlık + 3 karanlık) / 6 themes (3 light + 3 dark)** ve bilinen **10 yazı tipinden** oluşan yazı tipi seçici (sağ üst) · and a **font picker** with 10 familiar fonts (top-right corner)
- 🔎 Kitap adı / yazar ile arama (Open Library): yalnızca ad, yazar, tür ve sayfa sayısı çekilir
- 📊 Kendi ölçütlerini ve ağırlıklarını sınırsız ekle · Add as many scoring criteria and weights as you like
- 🟩 Renk ne kadar yoğunsa skor o kadar yüksek; puanlanmamış satırlar normal yazıyla görünür · The more intense the colour, the higher the score; unscored rows use normal text
- ✍️ Yazar / tür hücrelerinde otomatik tamamlama (Enter/Tab kabul eder) · Autocomplete for author / genre cells
- 🌐 Varsayılan ölçütler seçtiğin dile göre çevrilir · Default criteria follow your language
- 📣 **Duyuru kutusu** (başlıkta, kapatılamaz): üstüne gelince vurgulanır, tıklayınca **duyuru panosu** açılır. Duyurular 6 dilde gösterilir · **Announcement box** (in the header, always visible): hover to highlight, click to open the **announcement board**. Announcements are shown in all 6 languages
- ℹ️ **Hakkında kartı** (sağ üstteki ⓘ) · **About card** (ⓘ, top-right)
- 🗂️ **Sekmeli listeler:** istediğin kadar liste aç (en fazla 20); her listenin kendi ölçütleri ve kitapları olur. Sürükleyerek sırala, çift tıkla yeniden adlandır, sağ tıkla çoğalt / sil · **Tabbed lists:** keep up to 20 lists, each with its own criteria and books. Drag to reorder, double-click to rename, right-click to duplicate / delete
- 📖 **Okuma durumu:** bir kitaba sağ tıkla → Okunacak / Okuyorum / Okudum (çoklu seçimde hepsine uygulanır) veya listeden sil. Okuyorum, yarım bıraktım (⏸) ve okudum satırlarının solunda rozet görünür · **Reading status:** right-click a book → To read / Reading / Read (applies to the whole selection) or remove it from the list. Reading, paused (⏸) and read books get a badge on the left
- ➕ Tablonun en altında hep boş bir satır var: herhangi bir hücresine yazarak kitap ekle; “+ Elle ekle” ad, yazar, tür ve sayfa sayısını sorar · An empty row always sits at the bottom: type in any cell to add a book; “+ Add manually” asks for title, author, genre and pages
- ▦ Göz yormayan hücre ızgarası, sol alttaki **Izgara** düğmesiyle açılıp kapanır · A soft cell grid, toggled by the **Grid** button (bottom left)
- 📁 Excel'den içe aktarma (yeni sekme olarak, mevcut listenin üstüne yazmaz) / Excel'e dışa aktarma (açık liste) · Excel import (added as a new tab, never overwrites) / export (the open list)
- Hesap yok, reklam yok, takip yok · No account, no ads, no tracking

## İndirme / Download

[Releases](../../releases) sayfasından kendi sistemine uygun dosyayı indir / Grab the file for your system from [Releases](../../releases):

| Sistem / System | Dosya / File | Nasıl çalıştırılır / How to run |
|---|---|---|
| Windows | `NextBook-windows.zip` | Zip'i çıkar, `NextBook.exe` / Extract, run `NextBook.exe` |
| macOS (Apple Silicon) | `NextBook-macos-arm64.zip` | Zip'i çıkar, `NextBook.app` / Extract, open `NextBook.app` |
| macOS (Intel) | `NextBook-macos-intel.zip` | Zip'i çıkar, `NextBook.app` / Extract, open `NextBook.app` |
| Linux (x86-64) | `NextBook-linux.tar.gz` | `tar -xzf NextBook-linux.tar.gz && ./NextBook/NextBook` |

### Windows "korumalı" uyarısı verirse / If Windows SmartScreen warns you
Uygulama henüz dijital olarak imzalanmadığı için Windows SmartScreen "Bilinmeyen yayıncı" uyarısı gösterebilir. **Daha fazla bilgi → Yine de çalıştır** diyebilirsin. / The app is not code-signed yet, so Windows may show an "unknown publisher" warning (**More info → Run anyway**).

### macOS "doğrulanamadı" uyarısı verirse / If macOS blocks the app
Uygulama Apple tarafından imzalanmadığı için Gatekeeper ilk açılışta engelleyebilir. `NextBook.app`'e sağ tıklayıp **Aç**'ı seç ya da Terminal'de `xattr -dr com.apple.quarantine NextBook.app` çalıştır. / The app is not notarized, so Gatekeeper may block the first launch. Right-click `NextBook.app` → **Open**, or run `xattr -dr com.apple.quarantine NextBook.app`.

### Linux
Tkinter gerekmez (pakete dahildir); ancak sistemde X11/Wayland (XWayland) ve temel Tk kütüphaneleri bulunmalıdır. / Tkinter is bundled; you only need a desktop session with the usual X11 libraries.

### Güvenmek zorunda değilsin, kontrol edebilirsin / You do not have to trust it blindly
1. **Kaynak kod açık / Open source:** `next_book.py` tek dosyadır, okuyabilirsin. / a single readable file.
2. **Paketleri bu depo derler / Built by GitHub Actions:** `.github/workflows/build.yml`. Doğrulamak için / to verify: `gh attestation verify <dosya/file> --repo zekibilenay/nextbook`
3. **Sağlama toplamı / Checksum:** Her dosyanın yanında bir `.sha256` vardır. Windows: `Get-FileHash <dosya>` · macOS / Linux: `shasum -a 256 <dosya>`; çıktı `.sha256` ile aynı olmalı. / Each file ships with a `.sha256`; the output must match.
4. **Hiç paket istemiyorsan / No package at all:** Python 3.9+ ile `python next_book.py`

## Gizlilik / Privacy
- Verilerin yalnızca bilgisayarında tutulur: Windows `%APPDATA%\NextBook`, macOS `~/Library/Application Support/NextBook`, Linux `~/.local/share/NextBook`. / Your data stays on your computer (paths above).
- Ağ bağlantıları: (1) arama yaptığında `openlibrary.org`'a giden kitap adı / yazar sorgusu, (2) açılışta ve uygulama açık kaldıkça 6 saatte bir, duyuru dosyasını okumak için tek bir HTTPS GET isteği (kimlik, kitap verisi veya çerez gönderilmez; sunucu yalnızca `User-Agent: NextBook/<sürüm>` görür). / Network: (1) your search query to `openlibrary.org`, (2) one HTTPS GET at startup, and again every 6 hours while the app stays open, for the announcement file (no identifiers, no book data; the server only sees `User-Agent: NextBook/<version>`).
- Duyuruları kapatmak için veri klasöründeki `settings.json` içine `"announcements": false` yaz; duyuru kutusu da gizlenir. / To disable announcements set `"announcements": false` in `settings.json`; the announcement box is hidden too.

## Kaynaktan çalıştırma / Run from source
```
pip install -r requirements.txt
python next_book.py
pyinstaller --onedir --windowed --name NextBook --icon next_book.ico --add-data "next_book.ico;." next_book.py   # Windows
pyinstaller --onedir --windowed --name NextBook --add-data "next_book.ico:." --hidden-import openpyxl next_book.py   # macOS / Linux
```

## Duyuru sunucusu / Announcement feed

Tüm uygulamalar aynı JSON dosyasını okuyabilir (GitHub Pages, kendi sunucun, herhangi bir statik barındırma). Örnek: `announcements.example.json`.
Adresi `next_book.py` içindeki `ANNOUNCE_URL` ile (ya da `NEXTBOOK_ANNOUNCE_URL` ortam değişkeniyle) ayarla. Yalnızca HTTPS kabul edilir.

- `apps.<uygulama>.latest`: kullanıcının sürümü bundan eskiyse otomatik "yeni sürüm" duyurusu çıkar.
- `announcements[]`: `id` (benzersiz ad), `apps` (`["*"]` = hepsi), `min_version` / `max_version`, `starts` / `expires` (`YYYY-AA-GG`), `priority` (büyük olan önce), `type` (`update` · `new_app` · `recommended` · `info`), `text`, `url` (isteğe bağlı, yalnızca https).
- **Çok dilli metin / Multilingual text:** `text` dil koduyla yazılır: `tr`, `en`, `ru`, `de`, `fr`, `zh` (`zh-CN` gibi bölgeli yazımlar da kabul edilir). Kullanıcının dilinde metin yoksa `en`, o da yoksa dosyadaki ilk metin gösterilir; bu yüzden her duyuruya en azından `en` ekle. Metin en fazla 600 karakterdir. / Write `text` per language code; if the user's language is missing, `en` is used, then the first available text. Always include `en`. Max 600 characters.
- **Kapatılamaz / Not dismissible:** Kullanıcı duyuruyu kapatamaz. Eski duyuruların kalkması için `expires` kullan veya duyuruyu dosyadan sil. / Users cannot dismiss announcements, so use `expires` or remove the entry to retire one.
- **Duyuru türleri / Types:** `update` (yukarı ok), `new_app` (yıldız), `recommended` (kalp; "önerilen uygulama"), `info` (i). Simgeler yazı tipi karakteri değil, kodla çizilir; bu yüzden her bilgisayarda aynı ve tam ortalı görünür ve temaya uyar. Bilinmeyen bir tür yazarsan `info` sayılır. / The icons are drawn in code (not font glyphs), so they look identical and perfectly centred everywhere and follow the theme. An unknown type counts as `info`.
- **Uygulamaya özel duyuru / Per-app targeting:** `apps` alanı hangi uygulamaların göreceğini belirler: `["*"]` hepsi, `["nextbook"]` yalnızca Next Book, `["nextbook", "bilenay-reader"]` ikisi de. / `apps` decides which apps show an announcement.
- **Pano / Board:** Başlıktaki kutu duyuruları sırayla gösterir (birden fazlaysa `1/3 ›` gibi). Kutuya tıklayınca tüm duyurular panoda kartlar halinde açılır; `url` olan kartlarda bağlantı, kullanıcıya sorulduktan sonra açılır. / The header box rotates through announcements; clicking it opens a board with every announcement as a card. Links open only after the user confirms.

## Sürüm yayınlama / Releasing
```
git tag v1.4.3
git push --tags
```
GitHub Actions Windows, macOS (Apple Silicon + Intel) ve Linux paketlerini derler ve Releases'e yükler. Etiket, `next_book.py` içindeki `APP_VERSION` ile aynı olmalıdır; değilse derleme durur. Release yayınlandıktan sonra duyuru dosyasındaki "yeni sürüm" duyurusunun `max_version` değerini, yayınladığın sürümün bir öncekine ayarla (ör. 1.4.4 çıkınca `1.4.3`). Feed'de aynı anda `apps.nextbook.latest` de yazarsan uygulama kendi sade "yeni sürüm" duyurusunu ayrıca üretir ve eski sürümlerde iki benzer duyuru görünür; elle yazılmış sıcak bir duyuru kullanıyorsan `latest` yazma. / GitHub Actions builds the Windows, macOS (Apple Silicon + Intel) and Linux packages and attaches them to the release. The tag must match `APP_VERSION` in `next_book.py`, otherwise the build stops. After the release is published, set `max_version` of the "new version" announcement in the feed to the version before the one you released (e.g. `1.4.3` when 1.4.4 ships). If you also set `apps.nextbook.latest`, the app generates its own plain "new version" announcement and older versions show two similar ones, so leave `latest` out when you write that announcement by hand.

## Değişiklikler / Changelog

### 1.8.1
- **TR:** Okunacak durumundaki satırlarda durum sütununda çıkan noktalı odak dikdörtgeni düzeltildi.
- **EN:** Fixed a stray dotted focus rectangle in the status column of "to read" rows.

### 1.8.0
- **TR:** Durum rozetleri (okuyorum / yarım bıraktım / okudum) artık hücrede tam ortada; ızgara modunda sağa kayma düzeltildi. Kitap adı, yazar, tür ve sayfa sayısının dördü de boşalan satır otomatik silinir. Kitap adı tamamen silinebilir; yalnızca yazarı olan satırlar oluşturulabilir (tablodan ve "+ Elle ekle" penceresinden). **macOS ve Linux sürümleri eklendi** (macOS için uygulama simgesiyle); veri klasörü her sistemde o sistemin standardına göre seçilir, macOS'ta sağ tık / Ctrl+tık menüsü çalışır.
- **EN:** Status badges (reading / paused / read) are now truly centred in their cell; the right-shift in grid mode is fixed. A row whose title, author, genre and page count are all empty is removed automatically. The title can be cleared completely, so author-only rows are possible (in the table and in "+ Add manually"). **macOS and Linux builds were added** (with an app icon on macOS); the data folder follows each system's convention, and the right-click / Ctrl+click menu works on macOS.

### 1.7.0
- **TR:** Uygulama tam ekran boyutunda açılıyor. 6 tema: 3 aydınlık (cream, moss, ocean), 3 karanlık (night, plum, ember); tema topları gruplanmış durumda (önceki lavanta/kum seçimi varsayılana döner). "+ Elle ekle" artık kitap adının yanında yazar, tür ve sayfa sayısını da soruyor. Tablonun en altında her zaman boş bir satır var; herhangi bir sütununa yazınca kitap eklenir (Enter/Tab ile hızlıca devam edilir). Yeni durum: **Yarım bıraktım** (⏸ rozeti, sağ tık menüsünden); Excel'e yazılır ve içe aktarılır. Hücreler arasına göz yormayan bir ızgara eklendi; alt çubukta içe/dışa aktarmanın yanındaki **Izgara** düğmesiyle açılıp kapanır (seçim saklanır). Hakkında kartındaki GitHub bağlantısı kaldırıldı.
- **EN:** The app opens maximised. 6 themes: 3 light (cream, moss, ocean) and 3 dark (night, plum, ember), grouped in the picker (a previously chosen lavender/sand theme falls back to the default). "+ Add manually" now asks for author, genre and page count as well as the title. The table always ends with an empty row; typing in any column adds a book (Enter/Tab keep you moving). New status: **Paused** (⏸ badge, via the right-click menu), saved to and read from Excel. A soft grid between cells, toggled by the **Grid** button next to Import/Export (the choice is remembered). The GitHub link is gone from the About card.

### 1.6.0
- **TR:** Satırlara sağ tıklayınca menü açılıyor: kitabı **Okunacak / Okuyorum / Okudum** olarak işaretleyebilir veya listeden silebilirsin (çoklu seçimde hepsine uygulanır). Okuyorum için oynat, okudum için onay işaretli rozet satırın solunda görünür; rozetler kodla çizilir, her temada ve her satır renginde okunur. Durum Excel çıktısına "Durum" sütunu olarak yazılır ve içe aktarılır; bu sütun olmayan eski dosyalar sorunsuz açılır. Yeni listelerde puan üst sınırı varsayılan olarak **10** (ölçüt ağırlıkları aynı); mevcut listelerin üst sınırı değişmez, "Puanlama ayarları"ndan değiştirilebilir. Satır renkleri artık üst sınırın beşte biri genişliğinde 5 sabit bantla koyulaşır (10 üzerinden 0–2, 2–4, 4–6, 6–8, 8–10). Alt çubuktaki "Skora göre sırala" ve "Seçileni sil" düğmeleri kaldırıldı (sıralama için sütun başlığına tıkla, silmek için sağ tık ya da Delete tuşu). Yazar kutusundaki "isteğe bağlı" ibaresi kaldırıldı.
- **EN:** Right-clicking a row opens a menu: mark the book as **To read / Reading / Read** or remove it from the list (applies to the whole selection). Reading gets a play badge and read gets a check badge on the left of the row; the badges are drawn in code and stay legible on every theme and row colour. The status is written to the Excel export as a "Status" column and read back on import; older files without that column still open fine. New lists default to a score maximum of **10** (criteria weights unchanged); existing lists keep their maximum and it can be changed in "Scoring settings". Row colours now darken in 5 fixed bands, each a fifth of the maximum (out of 10: 0–2, 2–4, 4–6, 6–8, 8–10). The "Sort by score" and "Delete selected" buttons in the bottom bar are gone (click a column header to sort; use the right-click menu or the Delete key to remove). The "optional" note next to Author is gone.

### 1.5.0
- **TR:** Sekmeli listeler: üstteki sekme çubuğundan yeni liste oluşturabilir, listeleri sürükleyerek sıralayabilir, çift tıklayarak yeniden adlandırabilir, sağ tıklayarak çoğaltabilir veya silebilirsin (orta tık = kapat). Her listenin kendi ölçütleri, puan üst sınırı ve kitapları vardır. Excel'den içe aktarılan dosya artık açık listenin üstüne yazılmaz, dosya adıyla yeni bir sekme olarak eklenir. Excel'e aktarma açık listeyi yazar ve sayfa adı liste adı olur. Yazar / tür otomatik tamamlama tüm listelerden öneri verir. Eski tek listeli veri dosyası ilk açılışta otomatik olarak ilk sekmeye taşınır ve eski dosyanın kopyası `next_book.v1-yedek.json` olarak saklanır.
- **EN:** Tabbed lists: create new lists from the tab bar, drag to reorder, double-click to rename, right-click to duplicate or delete (middle-click closes). Each list has its own criteria, score scale and books. Excel files are no longer imported over the open list; they are added as a new tab named after the file. Excel export writes the open list and uses the list name as the sheet name. Author / genre autocomplete suggests from all lists. The old single-list data file is moved into the first tab on first launch, and a copy of the old file is kept as `next_book.v1-yedek.json`.

### 1.4.3
_(1.4.1 ve 1.4.2 numaraları yayın sırasında atlandı. / Version numbers 1.4.1 and 1.4.2 were skipped while releasing.)_

- **TR:** Yeni duyuru türü `recommended` ("önerilen uygulama", kalp simgesi). Duyuru simgeleri yazı tipi karakteri yerine kodla çizilen, kenarları yumuşatılmış rozetler oldu; "i" simgesinin aşağı kaçması sorunu giderildi. Duyurular uygulamaya göre hedeflenebiliyor (`apps` alanı). Eski sürümleri kullananlara "yeni sürüm" duyurusu gösteriliyor.
- **EN:** New announcement type `recommended` ("recommended app", heart icon). Announcement icons are now anti-aliased badges drawn in code instead of font glyphs, which fixes the misaligned "i" icon. Announcements can be targeted per app (`apps` field). People on older versions get a "new version" announcement.

### 1.4.0
- **TR:** Duyuru kutusu artık kapatılamıyor; üstüne gelince vurgulanıyor, tıklayınca duyuru panosu açılıyor. Duyurular 6 dilde gösteriliyor (bölgeli kodlar ve dil yedekleri dahil). Başlıktaki "Ara → listeye ekle → ..." kılavuz yazısı kaldırıldı. Yazı tipi seçici bilinen 10 yazı tipine indirildi (daha önce listede olmayan bir yazı tipi seçtiysen varsayılana döner). Puanlanmamış kitaplar artık soluk/italik değil, normal yazıyla gösteriliyor (Excel çıktısında da). "Hakkında" artık düz bir kutu yerine tema renklerine uyan bir bilgi kartı. Bir duyuru metni 300 yerine en fazla 600 karakter olabiliyor.
- **EN:** The announcement box can no longer be dismissed; it highlights on hover and opens an announcement board on click. Announcements are shown in all 6 languages (including region codes and language fallbacks). The "Search → add to list → ..." guide text in the header is gone. The font picker is down to 10 familiar fonts (a previously chosen font that is no longer listed falls back to the default). Unscored books are no longer faded/italic, they use normal text (also in the Excel export). "About" is now a themed info card instead of a plain message box. Announcement text can be up to 600 characters instead of 300.

### 1.3.1
- **TR:** Depo ve duyuru adresleri `zekibilenay` kullanıcı adına taşındı. Bozuk veri dosyası artık tarih damgalı bir yedeğe alınıp kullanıcıya bildiriliyor. Duyuru dosyası https'ten http'ye yönlendirilemiyor. Excel çıktısında `=` ile başlayan metinler formül sayılmıyor. Etiket ile `APP_VERSION` uyuşmazsa derleme duruyor.
- **EN:** Repository and feed URLs moved to the `zekibilenay` username. A corrupted data file is now moved to a timestamped backup and the user is told. The announcement feed cannot redirect from https to http. Text starting with `=` is no longer treated as a formula in the Excel export. The build stops if the tag and `APP_VERSION` differ.

Veriler / Data: [Open Library](https://openlibrary.org) · Lisans / License: MIT
