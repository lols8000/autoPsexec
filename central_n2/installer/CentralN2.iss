#define MyAppName "Central N2 Workstation"
#define MyAppVersion "5.0.0"
#define MyAppExeName "CentralN2.exe"
[Setup]
AppId={{18E9E450-3509-4A3C-9CB7-6C29C5100D20}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\CentralN2
DefaultGroupName=Central N2
OutputDir=Output
OutputBaseFilename=CentralN2-Setup
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
Compression=lzma2
SolidCompression=yes
[Files]
Source: "..\dist\CentralN2\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
[Icons]
Name: "{group}\Central N2"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Central N2"; Filename: "{app}\{#MyAppExeName}"
[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Executar Central N2"; Flags: nowait postinstall skipifsilent
