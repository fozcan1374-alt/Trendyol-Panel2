#!/usr/bin/env python3
"""AYAR — paketteki bütün yollar burada, PAKETİN KENDİ KONUMUNDAN türetilir.
Klasörü nereye taşırsan taşı çalışır; Windows/macOS/Linux fark etmez.

İki çalışma biçimi desteklenir:
  • Taşınabilir (zip'ten çıkarılmış)  → veri, uygulamanın yanındaki `veri/` klasöründe
  • Kurulumla gelmiş (Setup.exe)      → uygulama klasörü yazılabilir değilse
                                        veri `%LOCALAPPDATA%\\Pazaryeri Paneli\\` altına gider
Program Files yazma korumalı olduğu için bu ayrım şart.
"""
import pathlib, json, sys, os

# ── İKİ AYRI KÖK (exe'de bunlar farklı yerlerdir!)
#    KOK    = exe'nin durduğu klasör  → VERİ buraya yazılır
#    KAYNAK = .py dosyalarının yeri   → exe'de geçici çıkarma klasörü (_MEIPASS)
#    Bu ayrım yapılmazsa exe, içine gömülü trendyol_panel.py'yi bulamaz ve
#    hiçbir şey söylemeden kapanır.
if getattr(sys, 'frozen', False):          # PyInstaller ile .exe olmuş
    KOK    = pathlib.Path(sys.executable).resolve().parent
    KAYNAK = pathlib.Path(getattr(sys, '_MEIPASS', KOK))
else:
    KOK = KAYNAK = pathlib.Path(__file__).resolve().parent


def _yazilabilir(klasor):
    try:
        klasor.mkdir(parents=True, exist_ok=True)
        d = klasor / '.yazma_testi'
        d.write_text('x', encoding='utf-8')
        d.unlink()
        return True
    except Exception:
        return False


# ── veri kökü: uygulama klasörü yazılabilirse orası, değilse kullanıcı klasörü
if _yazilabilir(KOK):
    VERI_KOK = KOK
else:
    VERI_KOK = pathlib.Path(os.environ.get('LOCALAPPDATA') or pathlib.Path.home()) / 'Pazaryeri Paneli'
    VERI_KOK.mkdir(parents=True, exist_ok=True)

VERI = VERI_KOK / 'veri'
VERI.mkdir(parents=True, exist_ok=True)
(VERI / 'Tarifeler').mkdir(exist_ok=True)

# ── veri dosyaları
XLSX         = str(VERI / 'Trendyol_Karlilik.xlsx')  # ana çalışma kitabı
TARIFE_ARSIV = str(VERI / 'Tarifeler')               # haftalık export arşivi
DESI_DEFTER  = str(VERI / '_kargo_desi_gozlem.json') # otomatik oluşur
HB_DEFTER    = str(VERI / '_hb_siparisler.json')     # otomatik oluşur

# ── API bilgileri (yazılabilir kökte)
TY_API = str(VERI_KOK / 'trendyol_api.json')
HB_API = str(VERI_KOK / 'hb_api.json')

# ── scriptler (exe'de gömülü kaynak klasöründen)
PANEL_PY    = str(KAYNAK / 'trendyol_panel.py')
SYNC_PY     = str(KAYNAK / 'trendyol_sync.py')
GUNCELLE_PY = str(KAYNAK / 'trendyol_guncelle.py')
FATURA_PY   = str(KAYNAK / 'fatura_maliyet.py')
DONMUS      = bool(getattr(sys, 'frozen', False))   # exe olarak mı çalışıyoruz

# ── indirilenler klasörü (Windows'ta da bu yol doğrudur)
INDIRILENLER = str(pathlib.Path.home() / 'Downloads')

# ── kurulumla gelen boş çalışma kitabını ilk açılışta veri klasörüne taşı
_tohum = KOK / 'sablon' / 'Trendyol_Karlilik.xlsx'
if not _tohum.is_file():
    _tohum = KAYNAK / 'sablon' / 'Trendyol_Karlilik.xlsx'
if _tohum.is_file() and not pathlib.Path(XLSX).exists():
    try:
        import shutil
        shutil.copy2(_tohum, XLSX)
    except Exception:
        pass


def satici_no(varsayilan='*'):
    """Trendyol export dosyaları `<saticiNo>-tarih-saat.xlsx` diye iniyor.
    Satıcı numarası API dosyasından okunur, koda gömülmez."""
    try:
        return str(json.load(open(TY_API, encoding='utf-8'))['seller_id'])
    except Exception:
        return varsayilan
