# DOW 30 Tracker — One-Click Builder (GUI + Console)
# Run in Administrator PowerShell inside the project folder.
# Supports optional flags: -InstallDeps, -MakeInstaller, -Run

param(
  [switch]$InstallDeps = $false,
  [switch]$MakeInstaller = $false,
  [switch]$Run = $false
)

$ErrorActionPreference = "Stop"

# Auto-detect project dir
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectDir

Write-Host ">> Project: $ProjectDir" -ForegroundColor Cyan

# Paths
$Assets = Join-Path $ProjectDir "assets"
$Data   = Join-Path $ProjectDir "data"
$MainPy = Join-Path $ProjectDir "DOW30_Tracker_LIVE.py"

if (-not (Test-Path $MainPy)) {
    throw "Missing DOW30_Tracker_LIVE.py in $ProjectDir"
}
if (-not (Test-Path $Assets)) { New-Item -ItemType Directory -Force -Path $Assets | Out-Null }
if (-not (Test-Path $Data))   { New-Item -ItemType Directory -Force -Path $Data   | Out-Null }

# Optional: refresh deps (use your active venv if any)
if ($InstallDeps) {
    Write-Host ">> Installing/Upgrading build dependencies..." -ForegroundColor Cyan
    & python -V | Out-Null
    & python -m pip install --upgrade pip
    & python -m pip install --upgrade pyinstaller PyQt5 pandas yfinance openpyxl
    # Quiet some older hook weirdness
    & python -m pip install --upgrade typing_extensions
}

# --- FULL CLEAN ---
Write-Host ">> Cleaning previous artifacts..." -ForegroundColor Cyan
$toDelete = @("build","dist")
foreach ($d in $toDelete) { if (Test-Path $d) { Remove-Item $d -Recurse -Force -ErrorAction SilentlyContinue } }
Get-ChildItem $ProjectDir -Filter "*.spec" -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem $ProjectDir -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# Ensure 'data' exists so users have a predictable default folder
if (-not (Test-Path $Data))   { New-Item -ItemType Directory -Force -Path $Data | Out-Null }

$MainPy = (Resolve-Path -LiteralPath $MainPy).Path
$AssetsResolved = (Resolve-Path -LiteralPath $Assets).Path
$DataResolved = (Resolve-Path -LiteralPath $Data).Path
$IconPath = Join-Path $AssetsResolved "dow.ico"
$IconResolved = $null
try { $IconResolved = (Resolve-Path -LiteralPath $IconPath).Path } catch { }
if (-not $IconResolved) {
    Write-Host "WARNING: assets\dow.ico missing (EXE will use default icon)" -ForegroundColor Yellow
    $IconResolved = $IconPath
}

# Kill running copies so we can overwrite EXE safely
Write-Host ">> Stopping any running instances..." -ForegroundColor Cyan
$null = cmd /c "taskkill /IM DOW30_Tracker_LIVE.exe /F 2>nul" 
$null = cmd /c "taskkill /IM DOW30_Tracker_Console_LIVE.exe /F 2>nul"

# --- BUILD COMMANDS ---
# Bundle assets and (empty) data dir so the app has a writable sibling by default.
# Windows add-data format: "src;dest"
$CommonArgs = @(
  "--clean",
  "--noconfirm",
  "--icon", $IconResolved,
  "--add-data", "${AssetsResolved};assets",
  "--add-data", "${DataResolved};data",
  "--workpath", (Join-Path $ProjectDir "build"),
  "--distpath", (Join-Path $ProjectDir "dist"),
  "--specpath", $ProjectDir
)

function Invoke-PyInstaller {
    param(
        [string[]]$Arguments
    )
    Write-Host ("   python -m PyInstaller {0}" -f ($Arguments -join ' ')) -ForegroundColor DarkGray
    & python -m PyInstaller @Arguments
}

# GUI build
Write-Host ">> Building windowed EXE..." -ForegroundColor Cyan
$GuiCmd = $CommonArgs + @("--onefile", "--windowed", "--name", "DOW30_Tracker_LIVE", $MainPy)
Invoke-PyInstaller -Arguments $GuiCmd

# Console build
Write-Host ">> Building console EXE..." -ForegroundColor Cyan
$ConCmd = $CommonArgs + @("--onefile", "--console", "--name", "DOW30_Tracker_Console_LIVE", $MainPy)
Invoke-PyInstaller -Arguments $ConCmd

# Verify outputs
if (-not (Test-Path "dist\DOW30_Tracker_LIVE.exe")) { throw "GUI EXE missing" }
if (-not (Test-Path "dist\DOW30_Tracker_Console_LIVE.exe")) { throw "Console EXE missing" }

# --- OPTIONAL: INNO SETUP BUILD ---
if ($MakeInstaller) {
    $issPath = Join-Path $ProjectDir "DOW30_Tracker_LIVE.iss"
    if (-not (Test-Path $issPath)) {
        Write-Host "!! DOW30_Tracker_LIVE.iss not found. Skipping installer build." -ForegroundColor Yellow
    }
    else {
        $iscc = Get-Command iscc.exe -ErrorAction SilentlyContinue
        if ($null -eq $iscc) {
            Write-Host "!! Inno Setup Compiler (iscc.exe) not found in PATH. Install Inno Setup or launch from the Inno command prompt." -ForegroundColor Yellow
            Write-Host "   You can still run iscc manually against $issPath once it is available." -ForegroundColor Yellow
        }
        else {
            Write-Host ">> Compiling installer with Inno Setup..." -ForegroundColor Cyan
            & $iscc.Path $issPath
        }
    }
}

Write-Host ""
Write-Host ">> Build complete." -ForegroundColor Green
$GuiExe = Join-Path $ProjectDir "dist\DOW30_Tracker_LIVE.exe"
$ConsoleExe = Join-Path $ProjectDir "dist\DOW30_Tracker_Console_LIVE.exe"
Write-Host "   GUI:     $GuiExe"
Write-Host "   Console: $ConsoleExe"

if ($Run) {
    Write-Host ">> Launching GUI build..." -ForegroundColor Cyan
    Start-Process -FilePath $GuiExe
}
