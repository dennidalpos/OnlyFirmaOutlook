; Inno Setup script for OnlyFirmaOutlook
; Builds the canonical EXE installer

#ifndef AppVersion
#define AppVersion "1.0.0"
#endif

#define AppName "OnlyFirmaOutlook"
#define AppPublisher "Danny Perondi"
#define AppURL "https://github.com/dennidalpos/OnlyFirmaOutlook"
#define AppExeName "OnlyFirmaOutlook.Launcher.exe"

[Setup]
AppId={{D24FEF06-383E-4541-A893-6F1DE5942FE8}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
LicenseFile=..\..\LICENSE
OutputDir=..\output
OutputBaseFilename=OnlyFirmaOutlook-Setup-{#AppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Dist directory content (Launcher + win-x86 + win-x64)
Source: "..\..\dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Filename: "{app}\{#AppExeName}"; Flags: nowait postinstall skipifsilent
