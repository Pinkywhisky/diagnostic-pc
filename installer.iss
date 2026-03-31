#define MyAppVersion "1.3"

[Setup]
AppName=Diagnostic PC
AppVersion={#MyAppVersion}
VersionInfoVersion={#MyAppVersion}.0.0
VersionInfoTextVersion={#MyAppVersion}
DefaultDirName={commonpf}\DiagnosticPC
DefaultGroupName=Diagnostic PC
OutputBaseFilename=Setup_Diagnostic_PC
Compression=lzma
SolidCompression=yes

[Files]
Source: "debug_pc.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Diagnostic PC"; Filename: "{app}\debug_pc.exe"
Name: "{commondesktop}\Diagnostic PC"; Filename: "{app}\debug_pc.exe"

[Run]
Filename: "{app}\debug_pc.exe"; Description: "Lancer le diagnostic"; Flags: nowait postinstall skipifsilent