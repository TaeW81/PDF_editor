; Kunhwa PDF Editor - Inno Setup Script
; 건화 PDF 편집기 설치 프로그램

#define MyAppName "Kunhwa PDF Editor"
#define MyAppVersion "3.2"
#define MyAppPublisher "Kunhwa Engineering"
#define MyAppExeName "Kunhwa_PDF_Editor.exe"

[Setup]
AppId={{E8F7A6B5-C4D3-2E1F-0A9B-8C7D6E5F4A3B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=Kunhwa_PDF_Editor_v{#MyAppVersion}_Setup
SetupIconFile=data\kunhwa_logo.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

; 한국어 설치 UI
[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면에 바로가기 생성"; GroupDescription: "추가 옵션:"

[Files]
Source: "dist\Kunhwa_PDF_Editor\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{#MyAppName} 제거"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Kunhwa PDF Editor 실행"; Flags: nowait postinstall skipifsilent
