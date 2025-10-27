DOW 30 Tracker — packaged files

What's included:
- DOW30_Tracker_LIVE.py  (the app)
- assets/dow.png, assets/dow.ico  (icons)
- build.ps1  (PowerShell one-click builder for folder, deps, EXE, and .iss)

How to use:
1) Copy all files to C:\DOW30Tracker (keeping assets/ subfolder).
2) Open PowerShell as Admin.
3) Run:
   Set-ExecutionPolicy -Scope CurrentUser Bypass -Force
   cd C:\DOW30Tracker
   .\build.ps1
4) Wait for build to finish — EXEs land in C:\DOW30Tracker\dist
5) Open DOW30Tracker.iss in Inno Setup Compiler → Build.
6) Run your installer → done.
