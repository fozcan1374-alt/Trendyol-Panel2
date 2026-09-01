#!/usr/bin/env python3
"""GİRİŞ NOKTASI — panel açılmadan önce kurulum tamam mı diye bakar.

Doğrudan trendyol_panel.py açılsaydı Excel yokken çöker, .exe de hiçbir şey
söylemeden kapanırdı. Burada önce kontrol edilir:

  • API bilgileri girilmemişse → PANELDEN GİRİLECEK form açılır
    (bağlantı test edilir, doğruysa trendyol_api.json'a yazılır)
  • Ana Excel yoksa            → ne yapılacağını anlatan pencere açılır

Ayarları sonradan değiştirmek için:  AYARLAR.bat   (ya da  baslat.py --ayar)
"""
import sys, os, json, pathlib, base64
import ayar

# ── PyInstaller görünürlüğü ────────────────────────────────────────────
# trendyol_panel.py çalışma anında runpy ile yükleniyor; PyInstaller statik
# analizde onun importlarını GÖREMEZ ve exe'ye koymaz. Burada bir kez import
# edilirlerse pakete dahil olurlar. Silme — exe eksik modülle çıkar.
import math, warnings, subprocess, datetime, glob, time, io, re, shutil        # noqa: F401
import contextlib, statistics, importlib, importlib.util, runpy, collections   # noqa: F401
import openpyxl, openpyxl.styles, openpyxl.utils, requests                     # noqa: F401
try:
    import webview                                                             # noqa: F401
except Exception:
    pass
# ───────────────────────────────────────────────────────────────────────

ALANLAR = ('seller_id', 'api_key', 'api_secret')


def api_eksik():
    """trendyol_api.json'da boş/şablon kalan alanlar."""
    try:
        c = json.load(open(ayar.TY_API, encoding='utf-8'))
    except Exception:
        return list(ALANLAR)
    return [k for k in ALANLAR
            if not str(c.get(k, '')).strip() or str(c.get(k)).startswith('BURAYA')]


def excel_var():
    return pathlib.Path(ayar.XLSX).exists()


class Kurulum:
    """Kurulum penceresinin arka ucu."""

    def __init__(self):
        self.bitti = False

    def mevcut(self):
        def oku(y):
            try: return json.load(open(y, encoding='utf-8'))
            except Exception: return {}
        ty, hb = oku(ayar.TY_API), oku(ayar.HB_API)
        g = lambda v: '' if not v or str(v).startswith('BURAYA') else str(v)
        return json.dumps(dict(
            seller_id=g(ty.get('seller_id')), api_key=g(ty.get('api_key')),
            api_secret=g(ty.get('api_secret')),
            hb_merchant=g(hb.get('merchant_id')), hb_key=g(hb.get('service_key')),
            hb_agent=g(hb.get('user_agent')),
            excel=excel_var(), excel_yolu=str(pathlib.Path(ayar.XLSX)),
            klasor=str(ayar.KOK)), ensure_ascii=False)

    def test(self, seller_id, api_key, api_secret):
        """Girilen bilgilerle Trendyol'a bağlanıp ürün sayısını döndürür."""
        import requests
        sid, k, s = str(seller_id).strip(), str(api_key).strip(), str(api_secret).strip()
        if not (sid and k and s):
            return json.dumps({'ok': False, 'mesaj': 'Üç alanı da doldur.'}, ensure_ascii=False)
        tok = base64.b64encode(f'{k}:{s}'.encode()).decode()
        H = {'Authorization': f'Basic {tok}', 'User-Agent': f'{sid} - SelfIntegration'}
        try:
            r = requests.get(
                f'https://apigw.trendyol.com/integration/product/sellers/{sid}/products/approved',
                params={'page': 0, 'size': 1}, headers=H, timeout=25)
        except Exception as e:
            return json.dumps({'ok': False, 'mesaj': f'Bağlanılamadı: {e}'}, ensure_ascii=False)
        if r.status_code in (401, 403):
            return json.dumps({'ok': False, 'mesaj':
                'Yetki reddedildi (401/403). API anahtarı veya gizli anahtar yanlış.'},
                ensure_ascii=False)
        if r.status_code == 404:
            return json.dumps({'ok': False, 'mesaj': 'Satıcı numarası bulunamadı (404).'},
                              ensure_ascii=False)
        if r.status_code != 200:
            return json.dumps({'ok': False, 'mesaj': f'Trendyol {r.status_code} döndü.'},
                              ensure_ascii=False)
        try: n = r.json().get('totalElements')
        except Exception: n = None
        return json.dumps({'ok': True, 'mesaj': f'Bağlantı başarılı — {n} ürün görünüyor.'},
                          ensure_ascii=False)

    def kaydet(self, seller_id, api_key, api_secret, hb_merchant, hb_key, hb_agent):
        try:
            json.dump({'seller_id': str(seller_id).strip(),
                       'api_key': str(api_key).strip(),
                       'api_secret': str(api_secret).strip()},
                      open(ayar.TY_API, 'w', encoding='utf-8'), indent=2, ensure_ascii=False)
            json.dump({'merchant_id': str(hb_merchant or '').strip(),
                       'service_key': str(hb_key or '').strip(),
                       'user_agent': str(hb_agent or '').strip()},
                      open(ayar.HB_API, 'w', encoding='utf-8'), indent=2, ensure_ascii=False)
            try: os.chmod(ayar.TY_API, 0o600); os.chmod(ayar.HB_API, 0o600)
            except Exception: pass
            return json.dumps({'ok': True})
        except Exception as e:
            return json.dumps({'ok': False, 'mesaj': str(e)}, ensure_ascii=False)

    def kapat(self):
        """Kaydettikten sonra pencereyi kapatır — ardından panel açılır."""
        self.bitti = True
        try:
            import webview
            webview.windows[0].destroy()
        except Exception:
            pass
        return '1'


SAYFA = r"""<!DOCTYPE html><html lang="tr"><head><meta charset="utf-8"><style>
:root{--or:#f27a1a;--line:#e6e8ef;--mut:#697386}
*{box-sizing:border-box}
body{margin:0;padding:26px 30px;background:#f6f7fb;color:#101828;
 font:14px/1.6 -apple-system,'Segoe UI',Roboto,sans-serif}
h1{font-size:20px;margin:0 0 3px}
.s{color:var(--mut);font-size:13px;margin-bottom:20px}
.card{background:#fff;border:1px solid var(--line);border-radius:14px;padding:18px 20px;margin-bottom:14px}
.card h2{font-size:12px;text-transform:uppercase;letter-spacing:.6px;color:var(--mut);margin:0 0 14px}
label{display:block;font-size:11px;color:var(--mut);margin:0 0 4px}
input{width:100%;padding:9px 12px;border:1px solid var(--line);border-radius:9px;
 font-size:13px;outline:none;font-family:inherit}
input:focus{border-color:var(--or);box-shadow:0 0 0 3px #f27a1a22}
.row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
.row.tek{grid-template-columns:1fr}
.btn{background:var(--or);color:#fff;border:0;border-radius:10px;padding:10px 18px;
 font-size:13px;font-weight:700;cursor:pointer;font-family:inherit}
.btn:disabled{opacity:.5;cursor:default}
.btn.g{background:#fff;color:#101828;border:1px solid var(--line)}
.alt{display:flex;gap:10px;align-items:center;margin-top:4px}
.d{margin-top:12px;padding:11px 14px;border-radius:10px;font-size:13px;display:none}
.d.ok{display:block;background:#ecfdf5;color:#047857;border:1px solid #a7f3d0}
.d.no{display:block;background:#fef2f2;color:#b91c1c;border:1px solid #fecaca}
.ip{font-size:12px;color:var(--mut);margin-top:10px;line-height:1.7}
code{background:#f1f3f9;padding:2px 6px;border-radius:5px;font-size:12px;
 font-family:ui-monospace,Consolas,monospace}
.uyari{background:#fffbeb;border:1px solid #fde68a;border-left:4px solid #f59e0b}
.uyari .b{font-weight:800;color:#b45309;margin-bottom:5px}
</style></head><body>
<h1>Kurulum</h1>
<div class="s">API bilgilerini gir, bağlantıyı test et, kaydet. Panel hemen ardından açılır.</div>
<div id="excelUyari"></div>
<div class="card">
  <h2>Trendyol</h2>
  <div class="row tek"><div>
    <label>Satıcı numarası (Seller ID)</label>
    <input id="sid" placeholder="örn. 123456" autocomplete="off"></div></div>
  <div class="row">
    <div><label>API Key</label><input id="key" autocomplete="off"></div>
    <div><label>API Secret</label><input id="sec" autocomplete="off"></div>
  </div>
  <div class="alt">
    <button class="btn g" id="btest" onclick="test()">Bağlantıyı test et</button>
    <button class="btn" id="bkay" onclick="kaydet()" disabled>Kaydet ve paneli aç</button>
  </div>
  <div class="d" id="durum"></div>
  <div class="ip">Bilgiler: Trendyol Satıcı Paneli &rarr; Hesabım &rarr; Entegrasyon Bilgileri<br>
  Kaydedilecek klasör: <code id="klasor"></code></div>
</div>
<div class="card">
  <h2>Hepsiburada — istege bağlı</h2>
  <div class="row">
    <div><label>Merchant ID</label><input id="hbm" autocomplete="off"></div>
    <div><label>Servis anahtarı</label><input id="hbk" autocomplete="off"></div>
  </div>
  <div class="row tek"><div>
    <label>Entegratör kullanıcı adı (User-Agent)</label>
    <input id="hba" placeholder="entegratör adı" autocomplete="off"></div></div>
  <div class="ip">Boş bırakırsan Hepsiburada ekranları çalışmaz, Trendyol tarafı etkilenmez.</div>
</div>
<script>
const $=(id)=>document.getElementById(id);
async function yukle(){
  if(window._y) return; window._y=1;
  const d=JSON.parse(await window.pywebview.api.mevcut());
  $('sid').value=d.seller_id; $('key').value=d.api_key; $('sec').value=d.api_secret;
  $('hbm').value=d.hb_merchant; $('hbk').value=d.hb_key; $('hba').value=d.hb_agent;
  $('klasor').textContent=d.klasor;
  if(!d.excel){
    $('excelUyari').innerHTML='<div class="card uyari"><div class="b">Ana Excel dosyası yok</div>'
      +'<div style="font-size:13px">Şu dosya bulunamadı:<br><code>'+d.excel_yolu+'</code><br><br>'
      +'<b>Trendyol_Karlilik.xlsx</b> dosyasını <b>veri</b> klasörüne kopyala. Maliyet, desi, '
      +'kargo tarifesi ve hedef kârlar orada tutulur; onsuz panel açılmaz. API bilgilerini şimdi '
      +'kaydedebilirsin, Excel\'i koyduktan sonra tekrar aç.</div></div>';
  }
  ['sid','key','sec'].forEach(i=>$(i).oninput=()=>{$('bkay').disabled=true;});
}
function bilgi(ok,msg){const d=$('durum');d.className='d '+(ok?'ok':'no');d.textContent=msg;}
async function test(){
  const b=$('btest'); b.disabled=true; b.textContent='Deneniyor...';
  try{
    const r=JSON.parse(await window.pywebview.api.test($('sid').value,$('key').value,$('sec').value));
    bilgi(r.ok,r.mesaj); $('bkay').disabled=!r.ok;
  }catch(e){ bilgi(false,'Beklenmeyen hata: '+e); }
  b.disabled=false; b.textContent='Bağlantıyı test et';
}
async function kaydet(){
  const b=$('bkay'); b.disabled=true; b.textContent='Kaydediliyor...';
  const r=JSON.parse(await window.pywebview.api.kaydet(
    $('sid').value,$('key').value,$('sec').value,$('hbm').value,$('hbk').value,$('hba').value));
  if(r.ok){ bilgi(true,'Kaydedildi. Panel açılıyor...'); await window.pywebview.api.kapat(); }
  else{ bilgi(false,'Kaydedilemedi: '+r.mesaj); b.disabled=false; b.textContent='Kaydet ve paneli aç'; }
}
window.addEventListener('pywebviewready', yukle);
setTimeout(()=>{ if(window.pywebview&&window.pywebview.api) yukle(); }, 500);
</script></body></html>"""


def kurulum_penceresi():
    """Formu açar; kullanıcı kaydedip kapatınca döner."""
    import webview
    k = Kurulum()
    webview.create_window('Pazaryeri Paneli — Kurulum', html=SAYFA, js_api=k,
                          width=780, height=740, min_size=(700, 600))
    webview.start()
    return k.bitti


def excel_uyarisi():
    """API tamam ama Excel yoksa gösterilir."""
    import webview
    html = ('<!DOCTYPE html><html lang="tr"><head><meta charset="utf-8"><style>'
            "body{margin:0;padding:30px;background:#f6f7fb;color:#101828;"
            "font:14px/1.7 -apple-system,'Segoe UI',Roboto,sans-serif}"
            'h1{font-size:19px;margin:0 0 16px}'
            '.k{background:#fff;border:1px solid #fde68a;border-left:4px solid #f59e0b;'
            'border-radius:12px;padding:18px 20px}'
            '.b{font-weight:800;color:#b45309;margin-bottom:6px}'
            'code{background:#f1f3f9;padding:2px 6px;border-radius:5px;font-size:12px;'
            'font-family:ui-monospace,Consolas,monospace}'
            '</style></head><body><h1>Panel açılamadı</h1><div class="k">'
            '<div class="b">Ana Excel dosyası yok</div><div>Şu dosya bulunamadı:<br><code>'
            + str(pathlib.Path(ayar.XLSX)) +
            '</code><br><br><b>Trendyol_Karlilik.xlsx</b> dosyasını <b>veri</b> klasörüne kopyala, '
            'sonra tekrar çift tıkla.<br><br>Maliyet, desi, kargo tarifesi ve hedef kârlar bu '
            'dosyada tutulur.</div></div></body></html>')
    webview.create_window('Pazaryeri Paneli — eksik dosya', html=html, width=700, height=420)
    webview.start()


def tarifeleri_arsivle():
    """Downloads'a inmiş Trendyol export'larını sessizce veri/Tarifeler altına alır.
    Ayrı bir 'arşivle' adımı olmasın diye panel açılmadan önce kendiliğinden çalışır."""
    try:
        sys.path.insert(0, str(ayar.KAYNAK))
        import tarife_arsivle
        tarife_arsivle.arsivle(log=lambda *a, **k: None)
    except Exception:
        pass          # arşivleme başarısız olsa da panel açılmalı


def paneli_ac():
    sys.path.insert(0, str(ayar.KAYNAK))
    # Pencere açan kod `if __name__=='__main__'` içinde; import etmek yetmez.
    import runpy
    runpy.run_path(ayar.PANEL_PY, run_name='__main__')


def main():
    if '--ayar' in sys.argv or api_eksik():
        if not kurulum_penceresi():   # kaydetmeden kapattıysa
            return 1
        if api_eksik():
            return 1
    if not excel_var():
        try:
            excel_uyarisi()
        except Exception:
            print('Ana Excel dosyasi yok:', ayar.XLSX)
            try: input('\nKapatmak icin Enter...')
            except Exception: pass
        return 1
    tarifeleri_arsivle()      # yeni indirilen tarife dosyalarını kendiliğinden yerleştir
    paneli_ac()
    return 0


def hata_penceresi(metin):
    """Konsolsuz (pythonw) çalıştığı için beklenmedik hata görünmez olurdu —
    ekrana bas ve veri/hata.log'a yaz."""
    try:
        (pathlib.Path(ayar.VERI) / 'hata.log').write_text(metin, encoding='utf-8')
    except Exception:
        pass
    try:
        import webview, html as _h
        webview.create_window(
            'Pazaryeri Paneli — hata',
            html='<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
                 "body{margin:0;padding:26px;background:#f6f7fb;color:#101828;"
                 "font:13px/1.7 -apple-system,'Segoe UI',Roboto,sans-serif}"
                 'h1{font-size:18px;margin:0 0 12px;color:#b91c1c}'
                 'pre{background:#fff;border:1px solid #e6e8ef;border-radius:10px;padding:14px;'
                 'white-space:pre-wrap;font:12px ui-monospace,Consolas,monospace;overflow:auto}'
                 '</style></head><body><h1>Panel açılamadı</h1>'
                 '<pre>' + _h.escape(metin) + '</pre>'
                 '<p>Bu metin <b>veri\\hata.log</b> dosyasına da yazıldı.</p></body></html>',
            width=820, height=560)
        webview.start()
    except Exception:
        pass


if __name__ == '__main__':
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        import traceback
        hata_penceresi(traceback.format_exc())
        sys.exit(1)
