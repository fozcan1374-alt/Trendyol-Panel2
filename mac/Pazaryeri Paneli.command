#!/bin/zsh
# Pazaryeri Paneli — macOS başlatıcı. Çift tıkla.
# İlk açılışta eksik paketleri kurar, sonrasında doğrudan açılır.
cd "$(dirname "$0")"
if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 kurulu değil."
  echo "Kur:  https://www.python.org/downloads/macos/   (ya da:  brew install python)"
  echo; echo "Kapatmak için Enter'a bas."; read
  exit 1
fi
if ! python3 -c "import webview, openpyxl, requests" >/dev/null 2>&1; then
  echo "Gerekli paketler kuruluyor (yalnızca ilk açılışta)..."
  python3 -m pip install --user --quiet pywebview openpyxl requests || {
    echo "Paket kurulumu başarısız. Elle dene:"
    echo "  python3 -m pip install --user pywebview openpyxl requests"
    echo; echo "Kapatmak için Enter'a bas."; read; exit 1; }
fi
python3 baslat.py
