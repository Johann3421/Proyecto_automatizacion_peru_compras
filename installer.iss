; Inno Setup script for Peru Compras Bot
; Requires Inno Setup 6+

#define MyAppName "Peru Compras Bot"
#define MyAppVersion "1.3.5"
#define MyAppPublisher "Peru Compras"
#define MyAppExeName "peru_compras_bot.exe"

[Setup]
AppId={{C47D0A6A-BD3C-4F66-B4B4-6D6F0D0A5011}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=PeruComprasBot_Setup_1_3_5
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=assets\app_mascot.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Accesos directos:"; Flags: unchecked

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "productos.xlsx"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Ejecutar {#MyAppName}"; Flags: nowait postinstall skipifsilent
