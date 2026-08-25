#define MyAppName "Point"
#define MyAppVersion "8.20.6"
#define MyAppPublisher "Sunil Kumar Peela"
#define MyAppExeName "Point.exe"

[Setup]
AppId={{8D891A1F-4F9B-46CA-9706-37C5CB234125}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\Point
DefaultGroupName=Point
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\build-installer
OutputBaseFilename=Point-v8.20.6-Setup
SetupIconFile=..\point.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
UninstallDisplayName=Point Secure Local Data Workspace
UninstallDisplayIcon={app}\Point.exe
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
LicenseFile=EULA.txt
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"; Flags: unchecked
Name: "sampledata"; Description: "Install sample reports"; GroupDescription: "Optional data:"; Flags: unchecked

[Dirs]
Name: "{app}\Inbox"; Flags: uninsneveruninstall
Name: "{app}\Workspace"; Flags: uninsneveruninstall
Name: "{app}\Exports"; Flags: uninsneveruninstall
Name: "{app}\Logs"; Flags: uninsneveruninstall
Name: "{app}\Fetcher\Staging"; Flags: uninsneveruninstall
Name: "{app}\BrowserFetcher\Staging"; Flags: uninsneveruninstall
Name: "{app}\scripts"

[Files]
Source: "..\build\Point.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\build\PointFetcher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\build\PointBrowserFetcher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\scripts\point_xlsx_to_csv.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "point-security.conf"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\COMPLIANCE.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\sample\*.csv"; DestDir: "{app}\Inbox"; Tasks: sampledata; Flags: onlyifdoesntexist uninsneveruninstall

[Icons]
Name: "{group}\Point"; Filename: "{app}\Point.exe"; WorkingDir: "{app}"
Name: "{group}\Point Fetcher"; Filename: "{app}\PointFetcher.exe"; WorkingDir: "{app}"
Name: "{group}\Point Browser Fetcher"; Filename: "{app}\PointBrowserFetcher.exe"; WorkingDir: "{app}"
Name: "{group}\Point Fetcher Staging"; Filename: "{app}\Fetcher\Staging"
Name: "{group}\Point Inbox"; Filename: "{app}\Inbox"
Name: "{group}\Uninstall Point"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Point"; Filename: "{app}\Point.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\Point.exe"; Description: "Launch Point"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\Point.exe"
Type: files; Name: "{app}\PointFetcher.exe"
Type: files; Name: "{app}\PointBrowserFetcher.exe"
Type: files; Name: "{app}\scripts\point_xlsx_to_csv.ps1"
Type: files; Name: "{app}\README.md"
Type: files; Name: "{app}\COMPLIANCE.md"
Type: files; Name: "{app}\LICENSE.txt"
