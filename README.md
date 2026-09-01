# Pazaryeri Paneli

Trendyol ve Hepsiburada satıcı hesapları için masaüstü kâr/fiyat paneli.
Buybox takibi, komisyon baremi bazlı fiyat önerisi ve gerçek kesintilerle net kâr hesabı.

İki kurulum aynı kodu paylaşır, yalnızca başlatma biçimi farklıdır:

- **[`mac/`](mac/)** — `.command` dosyasına çift tık, paketleri kendi kurar
- **[`windows/`](windows/)** — `SETUP_YAP.bat` ile `.exe` + kurulum dosyası üretilir

Kurulum ve kullanım için ilgili klasörün README'sine bak.

## Öne çıkan

- **Buybox kontrolü** — buybox bizde mi, kaçıncı sıradayız, üstteki satıcı kaçtan satıyor.
  Buybox'ın yalnız fiyatla belirlenmediği durum (biz daha ucuzken sıra 2) ayrıca işaretlenir.
- **Kâr kabul aralığı** — kategori bazlı min/tavan marj. Marj aralık içindeyse fiyat doğrudur;
  yalnız altına düştüğünde yükseltme önerilir, hedef fiyat rakibin %1 altıyla sınırlanır.
- **Komisyon baremi** — Trendyol'da fiyat bir baremin altına inince komisyon düşer.
  Motor 4 bandın da en kârlı noktasını tarar; bazen fiyatı düşürmek kârı artırır.
- **Gerçek kesintiler** — komisyon · %1 e-ticaret stopajı · desi bazlı kargo
  (gerçek desi kargo faturalarından geri yazılır) · gönderi başına platform hizmet bedeli.

## Veri ve gizlilik

Repoda iş verisi yoktur: çalışma kitabı boş şablondur, API anahtarları kullanıcının kendi
makinesinde oluşur (`.gitignore` içinde). Satıcı numarası koda gömülü değildir, API
dosyasından okunur.

## Uyarı

Panelde **📤 Trendyol'a Gönder** düğmesi, onay diyaloğundan sonra canlı satış fiyatını
değiştirir. Salt-okunur kullanım için o düğme kaldırılmalıdır.
