[Setup]
AppName=VOIS AI Assistant
AppVersion=1.0
DefaultDirName={pf}\VOIS
DefaultGroupName=VOIS
OutputDir=output
OutputBaseFilename=VOIS_Installer
Compression=lzma
SolidCompression=yes
LicenseFile=license.txt
SetupIconFile=logo1.ico
WizardImageFile=logo_banner.bmp
DisableProgramGroupPage=yes

[Files]
Source: "dist\vois v1.0.0.exe"; DestDir: "{app}"; DestName: "VOIS.exe"; Flags: ignoreversion
Source: "logo1.ico"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "{app}\VOIS.exe"; Description: "Launch VOIS"; Flags: nowait postinstall skipifsilent

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
    ValueType: string; ValueName: "VOIS"; ValueData: """{app}\VOIS.exe"""

[Icons]
Name: "{group}\VOIS AI Assistant"; Filename: "{app}\VOIS.exe"; IconFilename: "{app}\logo1.ico"

[CustomMessages]
WelcomeLabel1=Welcome to the VOIS AI Assistant Setup Wizard.
WelcomeLabel2=This wizard will guide you through the installation of VOIS v1.0.0.

[UninstallRun]
Filename: "taskkill"; Parameters: "/F /IM VOIS.exe"; Flags: runhidden

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

