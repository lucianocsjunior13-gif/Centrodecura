#define MyAppName "Centro de Cura"
#define MyAppVersion "1.0.2"
#define MyAppPublisher "Centro de Cura"
#define MyAppExeName "CentroDeCura.exe"

[Setup]
AppId={{7C1F7D4F-C316-459F-AAA9-A18846F5A927}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

; Pasta de saída segura dentro do próprio projeto
OutputDir=C:\Users\PC\OneDrive\Documentos\Centrodecura\dist
OutputBaseFilename=Instalador_CentroDeCura
SetupIconFile=C:\Users\PC\OneDrive\Documentos\Centrodecura\logo.ico
SolidCompression=yes
WizardStyle=modern dynamic

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"

[Tasks]
; Caixa marcada por defeito para garantir que cria o atalho no Ambiente de Trabalho
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Copia o executável principal da pasta dist
Source: "C:\Users\PC\OneDrive\Documentos\Centrodecura\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Copia apenas os recursos visuais necessários (imagens e ícones)
Source: "C:\Users\PC\OneDrive\Documentos\Centrodecura\logo.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\PC\OneDrive\Documentos\Centrodecura\logocura.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\PC\OneDrive\Documentos\Centrodecura\logocura2.png"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
; Cria o atalho no Ambiente de Trabalho com o ícone oficial
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\logo.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent