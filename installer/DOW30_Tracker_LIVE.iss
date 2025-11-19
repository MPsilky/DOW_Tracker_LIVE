#define MyAppName "{{APP_NAME}}"
#define MyAppVersion "{{APP_VERSION}}"
#define MyAppPublisher "DOW Tracker"
#define MyAppExeName "{{APP_EXE}}"
#define MyCompanyURL "https://dow30tracker.example"
#define SourceRoot "{{PROJECT_DIR}}"

[Setup]
AppId={{D9C0DA7B-7F96-4F4D-889A-83ABF0A96C9D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyCompanyURL}
AppSupportURL={#MyCompanyURL}
AppUpdatesURL={#MyCompanyURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableDirPage=no
DisableProgramGroupPage=yes
LicenseFile={#SourceRoot}\README.md
OutputDir={#SourceRoot}\dist
OutputBaseFilename=DOW30_Tracker_LIVE-Installer
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile={#SourceRoot}\assets\dow.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a Desktop shortcut"; GroupDescription: "Shortcut options:"; Flags: unchecked
Name: "startmenuicon"; Description: "Add Start Menu shortcut"; GroupDescription: "Shortcut options:"; Flags: checkedonce
Name: "taskbaricon"; Description: "Pin to the taskbar (if supported)"; GroupDescription: "Shortcut options:"

[Files]
Source: "{#SourceRoot}\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\dist\DOW30_Tracker_Console_LIVE.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceRoot}\data\*"; DestDir: "{app}\data"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceRoot}\samples\*"; DestDir: "{app}\samples"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"; Flags: uninsneveruninstall

[Icons]
Name: "{autoprograms}\{#MyAppName}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\assets\dow.ico"; Tasks: startmenuicon
Name: "{autoprograms}\{#MyAppName}\Browse Excels"; Filename: "{app}\{#MyAppExeName}"; Parameters: "--browse"; WorkingDir: "{app}"; Tasks: startmenuicon
Name: "{autoprograms}\{#MyAppName}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\assets\dow.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
Filename: "{cmd}"; Parameters: "/c powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command \"$shell=New-Object -ComObject WScript.Shell;$lnkPath='$env:APPDATA\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar\\{#MyAppName}.lnk';$lnk=$shell.CreateShortcut($lnkPath);$lnk.TargetPath='{app}\\{#MyAppExeName}';$lnk.WorkingDirectory='{app}';$lnk.IconLocation='{app}\\assets\\dow.ico';$lnk.Save()\""; Flags: runhidden shellexec nowait skipifsilent; Tasks: taskbaricon

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\{#MyAppName}.lnk"
