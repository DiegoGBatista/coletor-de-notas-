; Instalador do Coletor de Notas Fiscais — compile com o Inno Setup 6
#define MyAppName "Coletor de Notas Fiscais"
#define MyAppVersion "3.1"
#define MyAppExeName "Coletor de Notas Fiscais.exe"

[Setup]
; >>> IMPORTANTE: para esta versão ATUALIZAR a que já está instalada, troque o AppId abaixo
; >>> pelo AppId do instalador antigo (veja o LEIA-ME.txt). Se deixar este, ela será instalada
; >>> como um programa novo (desinstale a versão antiga antes).
AppId={{1D42AD85-8DE0-4F18-A5F6-D19F16ACE845}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=Automação 2000
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=Output
OutputBaseFilename=Coletor_de_Notas_Fiscais_Setup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
; o programa cria este mutex ao abrir; assim o instalador percebe que ele está aberto
AppMutex=ColetorNotasFiscaisMutex
CloseApplications=yes

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\Coletor de Notas Fiscais\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent
