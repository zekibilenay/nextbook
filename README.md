# Next Book

Okuyacağın kitapları kendi ölçütlerine göre puanlayıp kişisel bir okuma sırası oluşturan, ücretsiz ve açık kaynaklı Windows uygulaması.

- Kitap adı / yazar ile arama (Open Library), sadece ad, yazar, tür ve sayfa sayısı çekilir
- Puanlanmamış satırlar soluk, puanlananlar koyu görünür
- Ölçütleri ve ağırlıkları sen belirlersin
- Excel'den içe aktarma / Excel'e dışa aktarma
- Hesap yok, reklam yok, takip yok

## İndirme

[Releases](../../releases) sayfasından en son `NextBook-windows.zip` dosyasını indir, zip'i bir klasöre çıkar, `NextBook.exe` dosyasını çalıştır.

### Windows "korumalı" uyarısı verirse
Uygulama henüz dijital olarak imzalanmadığı için Windows SmartScreen "Bilinmeyen yayıncı" uyarısı gösterebilir. **Daha fazla bilgi → Yine de çalıştır** diyebilirsin. Güvenmek zorunda değilsin, kontrol edebilirsin:

1. **Kaynak kod açık:** `next_book.py` tek dosyadır, okuyabilirsin.
2. **Exe'yi bu depo derler:** Release dosyası, GitHub Actions ile bu koddan otomatik üretilir (`.github/workflows/build.yml`). Doğrulamak için: `gh attestation verify NextBook-windows.zip --repo KULLANICI/DEPO`
3. **Sağlama toplamı:** PowerShell'de `Get-FileHash NextBook-windows.zip` çıktısı, release'teki `.sha256` dosyasıyla aynı olmalı.
4. **Hiç exe istemiyorsan:** Python 3.9+ kurulu ise `python next_book.py` ile doğrudan çalıştır.

## Gizlilik
- Verilerin yalnızca bilgisayarında, `%APPDATA%\NextBook\next_book.json` dosyasında tutulur.
- Uygulamanın tek ağ bağlantısı, arama yaptığında `openlibrary.org`'a gönderilen kitap adı / yazar sorgusudur.

## Kaynaktan çalıştırma / derleme
```
pip install -r requirements.txt
python next_book.py
pyinstaller --onedir --windowed --name NextBook next_book.py
```

## Sürüm yayınlama
```
git tag v1.0.0
git push --tags
```
GitHub Actions exe'yi derler ve Releases'e yükler.

Veriler [Open Library](https://openlibrary.org) üzerinden alınır. Lisans: MIT
