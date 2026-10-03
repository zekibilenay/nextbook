# Next Book

**TR** · Okuyacağın kitapları kendi ölçütlerine göre puanlayıp kişisel bir okuma sırası oluşturan, ücretsiz ve açık kaynaklı Windows uygulaması.
**EN** · A free, open-source Windows app that ranks the books you want to read using scoring criteria you define yourself.

## Özellikler / Features
- 🌍 **6 dil / 6 languages:** Türkçe, English, Русский, Deutsch, Français, 中文 (ilk açılışta seçilir / chosen on first launch)
- 🎨 **6 tema / 6 themes** and a **font picker** (top-right corner)
- 🔎 Kitap adı / yazar ile arama (Open Library): yalnızca ad, yazar, tür ve sayfa sayısı çekilir
- 📊 Kendi ölçütlerini ve ağırlıklarını sınırsız ekle · Add as many scoring criteria and weights as you like
- 🟩 Puanlanmamış satırlar soluk, puanlananlar skora göre koyulaşır · Unscored rows are faded; scored rows get more intense with the score
- ✍️ Yazar / tür hücrelerinde otomatik tamamlama (Enter/Tab kabul eder) · Autocomplete for author / genre cells
- 🌐 Varsayılan ölçütler seçtiğin dile göre çevrilir · Default criteria follow your language
- 📣 Başlıktaki duyuru kutusu: yeni sürüm ve yeni uygulama haberleri · Announcement box for new versions and apps
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
- Duyuruları kapatmak için `%APPDATA%\NextBook\settings.json` içine `"announcements": false` yaz. / To disable announcements set `"announcements": false` in `settings.json`.

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
- `announcements[]`: `apps` (`["*"]` = hepsi), `min_version` / `max_version`, `starts` / `expires`, `priority`, `type` (`update` · `new_app` · `info`), `text` (dil koduyla; yoksa `en`), `url` (isteğe bağlı, https).

## Sürüm yayınlama / Releasing
```
git tag v1.3.1
git push --tags
```
GitHub Actions exe'yi derler ve Releases'e yükler. / GitHub Actions builds the exe and attaches it to the release.

Veriler / Data: [Open Library](https://openlibrary.org) · Lisans / License: MIT
