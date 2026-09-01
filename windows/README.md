# Pazaryeri Paneli — Windows

Trendyol ve Hepsiburada satıcı hesaplarını tek masaüstü panelinde yöneten araç.
Buybox takibi, komisyon baremi bazlı fiyat önerisi, gerçek kesintilerle kâr hesabı.

## Ne yapar

- **🎯 Buybox & Fiyat** — stoktaki her üründe buybox bizde mi, kaçıncı sıradayız,
  bir üstteki satıcı kaçtan satıyor. Kâr marjı kategori için tanımlı kabul aralığının
  altındaysa fiyat yükseltme önerisi verir.
- **📈 Ciro Raporu** — sipariş tarihi bazlı ciro, komisyon, %1 stopaj, kargo,
  platform hizmet bedeli ve ürün maliyeti düşülmüş net kâr.
- **🗂 Ürünler** — satıştaki tüm ürünler; maliyet, desi ve komisyon girişi.
- **⚡ Fırsat Programları** — flaş ürün ve yıldızlı ürün tekliflerinin her birinde
  kalem kalem kâr dökümü.

Kâr hesabı gerçek kesintileri içerir: komisyon (fiyat baremine göre) · %1 e-ticaret
stopajı · desi bazlı kargo · gönderi başına platform hizmet bedeli · ürün maliyeti.

## Kurulum

Dağıtılabilir bir kurulum dosyası (`PazaryeriPaneli_Setup.exe`) üretilir; son kullanıcı
sadece onu çalıştırır. Üretim **Windows'ta bir kez** yapılır — PyInstaller da Inno Setup da
Windows'ta çalışır, Mac'te çapraz derleme yoktur.

1. Bu klasörü Windows makineye kopyala.
2. **`SETUP_YAP.bat`** dosyasını çalıştır. Sırayla:
   - Python, PyInstaller ve Inno Setup'ı `winget` ile kurar
   - `PazaryeriPanel.exe` üretir (PyInstaller, tek dosya, konsolsuz)
   - `pazaryeri.iss` ile **`PazaryeriPaneli_Setup.exe`** derler
3. Dağıtılan tek dosya bu setup'tır: İleri-İleri-Kur, Başlat Menüsü + masaüstü kısayolu,
   Denetim Masası'ndan kaldırma.

Kurulum `%LOCALAPPDATA%\Programs` altına gider — yönetici izni sormaz, klasör yazılabilir olur.
Setup imzasız olduğu için Windows "Bilinmeyen yayımcı" uyarısı verir; kaldırmak için kod
imzalama sertifikası gerekir.

### Kurulum yapmadan çalıştırmak (geliştirme)

```
pip install pywebview openpyxl requests
python baslat.py
```
pywebview için **Edge WebView2 Runtime** gerekir (çoğu Windows 11'de zaten var).

### İlk açılış

Panel, **Trendyol API bilgilerini** soran bir form açar (Satıcı No / API Key / API Secret).
Trendyol satıcı panelinde: Hesabım → Entegrasyon Bilgileri → API Anahtarı.
"Bağlantıyı test et" gerçekten istek atar; test geçmeden kaydetmez. Hepsiburada isteğe bağlı.
Ayarları sonradan değiştirmek: `python baslat.py --ayar`

## Veri nerede durur

Tüm veri **`veri/`** klasöründe. Uygulama klasörü yazılamıyorsa (Program Files gibi)
veri `%LOCALAPPDATA%\Pazaryeri Paneli\` altına kayar — güncelleme kullanıcının verisini silmez.

- `Trendyol_Karlilik.xlsx` — ana çalışma kitabı (maliyet, desi, komisyon tarifesi,
  hedef kârlar). İlk açılışta `sablon/` içindeki boş kitaptan oluşturulur.
- `Tarifeler/` — Trendyol'dan indirilen haftalık komisyon/flaş/yıldız export'ları,
  panel açılışında Downloads'tan otomatik arşivlenir.
- `_*.json` — otomatik oluşan defterler, elleme.

`veri/` ve API dosyaları `.gitignore` içindedir; repoya girmez.

## Haftalık akış

1. Trendyol satıcı panelinden komisyon tarifesi / flaş / yıldız export'larını indir
   (Downloads'a inmesi yeterli).
2. Paneli aç — export'lar otomatik arşivlenir.
3. Sağ üstteki **⟳ Güncelle** ile tarifeler işlenir, fiyat/stok senkronlanır.

## Dosyalar

| Dosya | İş |
|---|---|
| `SETUP_YAP.bat` | Windows'ta bir kez: exe + kurulum dosyası üretir |
| `pazaryeri.iss` | Inno Setup betiği |
| `baslat.py` | giriş noktası; kurulum formu, hata penceresi, paneli açar |
| `ayar.py` | bütün yollar; paketin kendi konumundan türetilir |
| `trendyol_panel.py` | panelin tamamı (arayüz + hesap motoru) |
| `trendyol_sync.py` | API'den fiyat/stok senkronu |
| `trendyol_guncelle.py` | haftalık tarife export'larını Excel'e işler |
| `tarife_arsivle.py` | export'ları `veri/Tarifeler/<hafta>/` altına arşivler |
| `fatura_maliyet.py` | e-fatura PDF'lerinden birim maliyet okur (isteğe bağlı, PyMuPDF) |

## Uyarı

Panelde ürün detayındaki **📤 Trendyol'a Gönder** düğmesi, onay diyaloğundan sonra
**canlı satış fiyatını değiştirir**. Salt-okunur kullanım istiyorsan o düğmeyi kaldır.
