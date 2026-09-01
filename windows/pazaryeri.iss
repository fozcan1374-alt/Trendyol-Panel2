; ══ Pazaryeri Paneli — Inno Setup kurulum betigi ══════════════════════
; Setup.exe uretir: Ileri > Ileri > Kur, Baslat Menusu + masaustu kisayolu,
; Denetim Masasi'ndan kaldirilabilir.
;
; PrivilegesRequired=lowest → yonetici sorulmaz, uygulama
; %LOCALAPPDATA%\Programs altina kurulur; orasi YAZILABILIR oldugu icin
; veri klasoru uygulamanin yaninda kalir.

#define Ad       "Pazaryeri Paneli"
#define Surum    "1.0.0"
#define Uretici  "Pazaryeri"
#define ExeAdi   "PazaryeriPanel.exe"

[Setup]
AppId={{8F3C2A91-4D7E-4B15-9C6A-2E8B7D4F1A03}
AppName={#Ad}
AppVersion={#Surum}
AppPublisher={#Uretici}
DefaultDirName={autopf}\{#Ad}
DefaultGroupName={#Ad}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=.
OutputBaseFilename=PazaryeriPaneli_Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#Ad}
UninstallDisplayIcon={app}\{#ExeAdi}

[Languages]
Name: "tr"; MessagesFile: "compiler:Languages\Turkish.isl"

[Tasks]
Name: "masaustu"; Description: "Masaustunde kisayol olustur"; GroupDescription: "Kisayollar:"

[Files]
Source: "PazaryeriPanel.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "OKUBENI.txt";        DestDir: "{app}"; Flags: ignoreversion
; Bos calisma kitabi "sablon" altina konur; uygulama ilk acilista
; veri\ klasorune kopyalar. Boylece GUNCELLEME kullanicinin verisini SILMEZ.
Source: "veri\Trendyol_Karlilik.xlsx"; DestDir: "{app}\sablon"; Flags: ignoreversion

[Icons]
Name: "{group}\{#Ad}";          Filename: "{app}\{#ExeAdi}"
; API bilgilerini degistirmek icin ayri kisayol — .bat dosyasina gerek kalmaz
Name: "{group}\Ayarlar";        Filename: "{app}\{#ExeAdi}"; Parameters: "--ayar"
Name: "{group}\{#Ad} kaldir";   Filename: "{uninstallexe}"
Name: "{autodesktop}\{#Ad}";    Filename: "{app}\{#ExeAdi}"; Tasks: masaustu

[Run]
Filename: "{app}\{#ExeAdi}"; Description: "{#Ad} uygulamasini simdi ac"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Kullanici verisi (veri\, api json) BILEREK silinmez — tekrar kurunca yerinde durur.
Type: filesandordirs; Name: "{app}\sablon"
