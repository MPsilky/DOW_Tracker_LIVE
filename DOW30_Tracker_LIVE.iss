#define MyAppName "DOW 30 Tracker"
#define MyAppVersion "1.3.0"
#define MyAppPublisher "You"
#define MyAppExeName "DOW30_Tracker_LIVE.exe"

[Setup]
AppId={{D9C0DA7B-7F96-4F4D-889A-83ABF0A96C9D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableDirPage=no
DisableProgramGroupPage=no
AllowNoIcons=yes
OutputDir=dist
OutputBaseFilename=DOW30_Tracker_LIVE-Installer
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=assets\dow.ico
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a Desktop shortcut"; GroupDescription: "Shortcut options:"; Flags: unchecked
Name: "startmenuicon"; Description: "Create a Start Menu shortcut"; GroupDescription: "Shortcut options:"; Flags: checkedonce
Name: "taskbarpin"; Description: "Pin DOW 30 Tracker to the taskbar (current user)"; GroupDescription: "Shortcut options:"; Flags: unchecked
Name: "autostart"; Description: "Enable auto-start at login"; GroupDescription: "Additional options:"; Flags: unchecked

[Dirs]
Name: "{app}\data"; Flags: uninsalwaysuninstall
Name: "{app}\assets"; Flags: uninsalwaysuninstall
Name: "{app}\logs"; Flags: uninsalwaysuninstall

[Files]
Source: "dist\DOW30_Tracker_LIVE.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\DOW30_Tracker_Console_LIVE.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "data\*"; DestDir: "{app}\data"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: startmenuicon
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\{#MyAppName}.lnk"; Filename: "{app}\{#MyAppExeName}"; Tasks: taskbarpin; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppExeName}"; IconIndex: 0; Flags: createonlyiffileexists

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "DOW30Tracker"; ValueData: "\"{app}\{#MyAppExeName}\" --autostart"; Tasks: autostart; Flags: uninsdeletevalue

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{cmd}"; Parameters: "/C taskkill /IM {#MyAppExeName} /F"; Flags: runhidden
