@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Setup.exe uret
cls
echo.
echo   ╔══════════════════════════════════════════════════════════════╗
echo   │  KURULUM DOSYASI URETICI                                     │
echo   │                                                              │
echo   │  Bu islem PazaryeriPaneli_Setup.exe dosyasini uretir.        │
echo   │  BIR KERE calistirilir, 5-10 dakika surer.                   │
echo   │  Sonrasinda dagitacagin tek dosya o Setup.exe olur.          │
echo   ╚══════════════════════════════════════════════════════════════╝
echo.
pause

rem ── 1) Python
where python >nul 2>&1
if errorlevel 1 (
  echo.
  echo   [1/4] Python kuruluyor...
  winget install -e --id Python.Python.3.12 --scope user --silent ^
    --accept-source-agreements --accept-package-agreements
  echo.
  echo   Python kuruldu. Bu pencere kapanacak, dosyaya TEKRAR cift tikla.
  timeout /t 4 >nul
  exit
)
echo   [1/4] Python var.

rem ── 2) Paketler + PyInstaller
echo   [2/4] Gerekli bilesenler kuruluyor...
python -m pip install --quiet --disable-pip-version-check --upgrade pip
python -m pip install --quiet --disable-pip-version-check pyinstaller pywebview pythonnet openpyxl requests
if errorlevel 1 (
  echo   ! Bilesenler kurulamadi. Internet baglantisini kontrol et.
  pause
  exit
)

rem ── 3) exe
echo   [3/4] Uygulama derleniyor (birkac dakika)...
python -m PyInstaller --noconfirm --onefile --windowed --name PazaryeriPanel ^
  --collect-all pywebview --collect-all clr_loader ^
  --collect-all openpyxl --collect-submodules requests --hidden-import ayar ^
  --add-data "ayar.py;." --add-data "trendyol_panel.py;." ^
  --add-data "trendyol_sync.py;." --add-data "trendyol_guncelle.py;." ^
  --add-data "tarife_arsivle.py;." --add-data "fatura_maliyet.py;." ^
  baslat.py >nul 2>&1
if not exist "dist\PazaryeriPanel.exe" (
  echo   ! Derleme basarisiz. Ayrinti icin:
  echo     python -m PyInstaller --onefile --windowed baslat.py
  pause
  exit
)
move /y "dist\PazaryeriPanel.exe" "PazaryeriPanel.exe" >nul
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q PazaryeriPanel.spec 2>nul
echo   [3/4] PazaryeriPanel.exe hazir.

rem ── 4) Inno Setup ile kurulum dosyasi
echo   [4/4] Kurulum dosyasi uretiliyor...
set "ISCC="
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC (
  echo         Inno Setup kuruluyor...
  winget install -e --id JRSoftware.InnoSetup --silent ^
    --accept-source-agreements --accept-package-agreements
  if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
  if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
)
if not defined ISCC (
  echo.
  echo   ! Inno Setup bulunamadi. jrsoftware.org/isdl.php adresinden kurup
  echo     bu dosyayi tekrar calistir.
  echo     ^(PazaryeriPanel.exe hazir, onu da dogrudan kullanabilirsin.^)
  pause
  exit
)
"%ISCC%" pazaryeri.iss >nul 2>&1
if exist "PazaryeriPaneli_Setup.exe" (
  cls
  echo.
  echo   ╔══════════════════════════════════════════════════════════════╗
  echo   │  HAZIR                                                       │
  echo   │                                                              │
  echo   │  PazaryeriPaneli_Setup.exe                                   │
  echo   │                                                              │
  echo   │  Dagitacagin tek dosya bu. Karsi taraf cift tiklar,          │
  echo   │  Ileri - Ileri - Kur der; Baslat Menusu ve masaustu          │
  echo   │  kisayolu olusur, Denetim Masasi'ndan kaldirilabilir.        │
  echo   ╚══════════════════════════════════════════════════════════════╝
) else (
  echo   ! Kurulum dosyasi uretilemedi.
  "%ISCC%" pazaryeri.iss
)
echo.
pause
