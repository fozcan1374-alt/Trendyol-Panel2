# Pazaryeri Paneli — macOS

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

> **İndirdikten sonra iki adım gerekebilir (macOS güvenliği):**
> ```bash
> cd "<klasörün yolu>"
> chmod +x "Pazaryeri Paneli.command"     # GitHub'dan zip olarak inince çalıştırma izni kaybolur
> xattr -dr com.apple.quarantine .        # "geliştirici doğrulanamadı" uyarısını kaldırır
> ```
> Alternatif: `.command` dosyasına **sağ tık → Aç** deyip bir kez "Aç"ı onayla.

1. **Python 3** gerekli (yoksa `brew install python` ya da python.org).
2. `Pazaryeri Paneli.command` dosyasına çift tıkla. İlk açılışta `pywebview`,
   `openpyxl` ve `requests` paketlerini kurar.
3. Açılan pencereden **Trendyol API bilgilerini** gir (Satıcı No / API Key / API Secret).
   Trendyol satıcı panelinde: Hesabım → Entegrasyon Bilgileri → API Anahtarı.
   "Bağlantıyı test et" gerçekten istek atar; test geçmeden kaydetmez.
   Hepsiburada isteğe bağlı.
4. Panel açılır. Ürün / fiyat / stok / komisyon / buybox API'den gelir;
   elle girilecek tek şey **birim maliyet**tir.

Ayarları sonradan değiştirmek için: `python3 baslat.py --ayar`

## Veri nerede durur

Tüm veri, uygulamanın yanındaki **`veri/`** klasöründe:

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
