# Next Book

**TR** · Okuyacağın kitapları kendi ölçütlerine göre puanlayıp kişisel bir okuma sırası oluşturan, ücretsiz ve açık kaynaklı Windows uygulaması.
**EN** · A free, open-source Windows app that ranks the books you want to read using scoring criteria you define yourself.

## Özellikler / Features
- 🌍 **6 dil / 6 languages:** Türkçe, English, Русский, Deutsch, Français, 中文 (ilk açılışta seçilir / chosen on first launch)
- 🎨 **6 tema / 6 themes** ve bilinen **10 yazı tipinden** oluşan yazı tipi seçici (sağ üst) · and a **font picker** with 10 familiar fonts (top-right corner)
- 🔎 Kitap adı / yazar ile arama (Open Library): yalnızca ad, yazar, tür ve sayfa sayısı çekilir
- 📊 Kendi ölçütlerini ve ağırlıklarını sınırsız ekle · Add as many scoring criteria and weights as you like
- 🟩 Renk ne kadar yoğunsa skor o kadar yüksek; puanlanmamış satırlar normal yazıyla görünür · The more intense the colour, the higher the score; unscored rows use normal text
- ✍️ Yazar / tür hücrelerinde otomatik tamamlama (Enter/Tab kabul eder) · Autocomplete for author / genre cells
- 🌐 Varsayılan ölçütler seçtiğin dile göre çevrilir · Default criteria follow your language
- 📣 **Duyuru kutusu** (başlıkta, kapatılamaz): üstüne gelince vurgulanır, tıklayınca **duyuru panosu** açılır. Duyurular 6 dilde gösterilir · **Announcement box** (in the header, always visible): hover to highlight, click to open the **announcement board**. Announcements are shown in all 6 languages
- ℹ️ **Hakkında kartı** (sağ üstteki ⓘ) · **About card** (ⓘ, top-right)
- 📁 Excel'den içe aktarma / Excel'e dışa aktarma · Excel import / export
- Hesap yok, reklam yok, takip yok · No account, no ads, no tracking

## İndirme / Download

[Releases](../../releases) sayfasından en son `NextBook-windows.zip` dosyasını indir, zip'i bir klasöre çıkar, `NextBook.exe` dosyasını çalıştır.
Download the latest `NextBook-windows.zip` from [Releases](../../releases), extract it and run `NextBook.exe`.

### Windows "korumalı" uyarısı verirse / If Windows SmartScreen warns you
Uygulama henüz dijital olarak imzalanmadığı için Windows SmartScreen "Bilinmeyen yayıncı" uyarısı gösterebilir. **Daha fazla bilgi → Yine de çalıştır** diyebilirsin. Güvenmek zorunda değilsin, kontrol edebilirsin:
The app is not code-signed yet, so Windows may show an "unknown publisher" warning (**More info → Run anyway**). You do not have to trust it blindly:

1. **Kaynak kod açık / Open source:** `next_book.py` tek dosyadır, okuyabilirsin. / a single readable file.
2. **Exe'yi bu depo derler / Built by GitHub Actions:** `.github/workflows/build.yml`. Doğrulamak için / to verify: `gh attestation verify NextBook-windows.zip --repo zekibilenay/nextbook`
3. **Sağlama toplamı / Checksum:** PowerShell: `Get-FileHash NextBook-windows.zip` çıktısı release'teki `.sha256` ile aynı olmalı / must match the `.sha256` file.
4. **Hiç exe istemiyorsan / No exe at all:** Python 3.9+ ile `python next_book.py`

## Gizlilik / Privacy
- Verilerin yalnızca bilgisayarında, `%APPDATA%\NextBook` klasöründe tutulur. / Your data stays on your computer.
- Ağ bağlantıları: (1) arama yaptığında `openlibrary.org`'a giden kitap adı / yazar sorgusu, (2) açılışta ve uygulama açık kaldıkça 6 saatte bir, duyuru dosyasını okumak için tek bir HTTPS GET isteği (kimlik, kitap verisi veya çerez gönderilmez; sunucu yalnızca `User-Agent: NextBook/<sürüm>` görür). / Network: (1) your search query to `openlibrary.org`, (2) one HTTPS GET at startup, and again every 6 hours while the app stays open, for the announcement file (no identifiers, no book data; the server only sees `User-Agent: NextBook/<version>`).
- Duyuruları kapatmak için `%APPDATA%\NextBook\settings.json` içine `"announcements": false` yaz; duyuru kutusu da gizlenir. / To disable announcements set `"announcements": false` in `settings.json`; the announcement box is hidden too.

## Kaynaktan çalıştırma / Run from source
```
pip install -r requirements.txt
python next_book.py
pyinstaller --onedir --windowed --name NextBook --icon next_book.ico --add-data "next_book.ico;." next_book.py
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
git tag v1.4.1
git push --tags
```
GitHub Actions exe'yi derler ve Releases'e yükler. Etiket, `next_book.py` içindeki `APP_VERSION` ile aynı olmalıdır; değilse derleme durur. Release yayınlandıktan sonra duyuru dosyasındaki "yeni sürüm" duyurusunun `max_version` değerini, yayınladığın sürümün bir öncekine ayarla (ör. 1.4.2 çıkınca `1.4.1`). Feed'de aynı anda `apps.nextbook.latest` de yazarsan uygulama kendi sade "yeni sürüm" duyurusunu ayrıca üretir ve eski sürümlerde iki benzer duyuru görünür; elle yazılmış sıcak bir duyuru kullanıyorsan `latest` yazma. / GitHub Actions builds the exe and attaches it to the release. The tag must match `APP_VERSION` in `next_book.py`, otherwise the build stops. After the release is published, set `max_version` of the "new version" announcement in the feed to the version before the one you released (e.g. `1.4.1` when 1.4.2 ships). If you also set `apps.nextbook.latest`, the app generates its own plain "new version" announcement and older versions show two similar ones, so leave `latest` out when you write that announcement by hand.

## Değişiklikler / Changelog

### 1.4.1
- **TR:** Yeni duyuru türü `recommended` ("önerilen uygulama", kalp simgesi). Duyuru simgeleri yazı tipi karakteri yerine kodla çizilen, kenarları yumuşatılmış rozetler oldu; "i" simgesinin aşağı kaçması sorunu giderildi. Duyurular uygulamaya göre hedeflenebiliyor (`apps` alanı). Eski sürümleri kullananlara "yeni sürüm" duyurusu gösteriliyor.
- **EN:** New announcement type `recommended` ("recommended app", heart icon). Announcement icons are now anti-aliased badges drawn in code instead of font glyphs, which fixes the misaligned "i" icon. Announcements can be targeted per app (`apps` field). People on older versions get a "new version" announcement.

### 1.4.0
- **TR:** Duyuru kutusu artık kapatılamıyor; üstüne gelince vurgulanıyor, tıklayınca duyuru panosu açılıyor. Duyurular 6 dilde gösteriliyor (bölgeli kodlar ve dil yedekleri dahil). Başlıktaki "Ara → listeye ekle → ..." kılavuz yazısı kaldırıldı. Yazı tipi seçici bilinen 10 yazı tipine indirildi (daha önce listede olmayan bir yazı tipi seçtiysen varsayılana döner). Puanlanmamış kitaplar artık soluk/italik değil, normal yazıyla gösteriliyor (Excel çıktısında da). "Hakkında" artık düz bir kutu yerine tema renklerine uyan bir bilgi kartı. Bir duyuru metni 300 yerine en fazla 600 karakter olabiliyor.
- **EN:** The announcement box can no longer be dismissed; it highlights on hover and opens an announcement board on click. Announcements are shown in all 6 languages (including region codes and language fallbacks). The "Search → add to list → ..." guide text in the header is gone. The font picker is down to 10 familiar fonts (a previously chosen font that is no longer listed falls back to the default). Unscored books are no longer faded/italic, they use normal text (also in the Excel export). "About" is now a themed info card instead of a plain message box. Announcement text can be up to 600 characters instead of 300.

### 1.3.1
- **TR:** Depo ve duyuru adresleri `zekibilenay` kullanıcı adına taşındı. Bozuk veri dosyası artık tarih damgalı bir yedeğe alınıp kullanıcıya bildiriliyor. Duyuru dosyası https'ten http'ye yönlendirilemiyor. Excel çıktısında `=` ile başlayan metinler formül sayılmıyor. Etiket ile `APP_VERSION` uyuşmazsa derleme duruyor.
- **EN:** Repository and feed URLs moved to the `zekibilenay` username. A corrupted data file is now moved to a timestamped backup and the user is told. The announcement feed cannot redirect from https to http. Text starting with `=` is no longer treated as a formula in the Excel export. The build stops if the tag and `APP_VERSION` differ.

Veriler / Data: [Open Library](https://openlibrary.org) · Lisans / License: MIT
