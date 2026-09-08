; Inno Setup script for SpecCleanse.
;
; Build (from the project root, after PyInstaller has produced dist\SpecCleanse.exe):
;
;     iscc /DAppVersion=1.1.0 packaging\installer.iss
;
; AppVersion is supplied by the caller rather than stored here, because the git
; tag is the version — see the Releases section in CLAUDE.md. The fallback below
; only exists so a local build without the flag still works.

#ifndef AppVersion
  #define AppVersion "0.0.0-dev"
#endif

#define AppName "SpecCleanse"
#define AppPublisher "Abraham Borg"
#define AppURL "https://github.com/Abe-Borg/Spec_Cleanse"
#define AppExeName "SpecCleanse.exe"

[Setup]
AppId={{8B3D1F42-6C7A-4E59-9A21-5F0E7C4D2B18}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}/issues
AppUpdatesURL={#AppURL}/releases
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE.md
; Installs for the current user only, so no administrator prompt and no admin
; rights required. {autopf} resolves to %LOCALAPPDATA%\Programs under this
; setting, which is also writable — the app can seed its own config there.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\dist
OutputBaseFilename={#AppName}-Setup-{#AppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; The GUI needs a desktop session; there is nothing here for older Windows.
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "..\dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\CHANGELOG.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Nothing here removes %APPDATA%\SpecCleanse\patterns.yaml. That file holds the
; user's own detection patterns, which they may well want after a reinstall, and
; silently deleting edited configuration on uninstall is a bad surprise.
