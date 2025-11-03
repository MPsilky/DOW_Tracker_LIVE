# DOW 30 Tracker (V1 Final)

The Dow 30 Tracker desktop app captures eight intraday buckets for all Dow Jones components, renders a live PyQt5 grid, and exports end-of-day Excel workbooks. Version 1 Final moves the network, file-system, and Excel work off the UI thread, adds resilient caching plus a secondary data feed, and guarantees the 4:00 PM export even when the app is closed early.

## What's new in V1 Final

* **Threaded capture pipeline**  `_load_intraday_cache`, `backfill_to_now`, and `refresh_now` now run inside a worker pool. The main window stays responsive while captures happen in the background.
* **Hardened intraday caching**  cache entries persist to `data/intraday_cache`, retry on stale timestamps, and track which tickers still need fresh data. Stale values are flagged in the UI.
* **Fallback price source**  when Yahoo Finance is empty, the app can hydrate missing tickers via Alpha Vantage (set `ALPHAVANTAGE_API_KEY`) or the persisted cache. All fallback usage is logged.
* **Guaranteed 4:00 PM export**  closing the app or hitting the final bucket forces a fresh Excel flush. A process `atexit` hook performs a last-minute export if the app exits unexpectedly.
* **Diagnostics**  `tracker.log` records capture status, fallback usage, and worker lifecycle. A sample excerpt lives in `samples/tracker_log_sample.txt`.

A mock workbook that mirrors the 328 grid is available at `samples/Sheet__Sample.xlsx`.

## Quick start

```bash
python .\DOW30_Tracker_LIVE.py
```

Launches the console build with live logging. Use this command while validating network keys or data folders.

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1 -Run
```

Builds the one-file GUI and console executables (plus the Inno Setup script) and starts the freshly built GUI app. Run from an elevated PowerShell inside the repo.

## Runtime expectations

* The app writes state under `%LOCALAPPDATA%\DOW30Tracker` (Windows) or `~/.dow30tracker` (macOS/Linux).
* Excel exports land in the chosen data directory (`data/` by default). Each capture updates `Sheet__MM_DD_YYYY.xlsx` (Grid sheet plus per-bucket sheet). The 4:00 PM bucket forces a final export; closing the window or exiting from the tray does the same.
* `tracker.log` records every capture. Review `samples/tracker_log_sample.txt` for the log format.
* To enable the Alpha Vantage fallback, set `ALPHAVANTAGE_API_KEY` in the environment before launching the app. Without the key, the persisted cache is still used as a secondary source.

## Building & packaging

1. Optional dependency install:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\build.ps1 -InstallDeps
   ```
   Installs/updates PyInstaller, PyQt5, pandas, yfinance, openpyxl, and requests.
2. Build windowed + console executables (and refresh the data folder structure):
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\build.ps1
   ```
3. To stage the Inno Setup script, add `-MakeInstaller`. The generated script and binaries live in `dist/`.

`build.ps1` ensures `data/intraday_cache/` exists so persisted series survive packaging.

## Continuous build workflow

GitHub Actions runs `.github/workflows/v1-final.yml` on pushes and PRs. The workflow installs dependencies, runs `python -m compileall DOW30_Tracker_LIVE.py`, executes the builder in dry-run mode, and uploads the PyInstaller and installer artifacts.

To trigger manually, use the Run workflow button in the Actions tab and target `codex/recent`.

## Samples & artifacts

* `samples/Sheet__Sample.xlsx`  a 328 grid plus the closing bucket sheet matching the live UI layout.
* `samples/tracker_log_sample.txt`  example log lines showing threaded captures and fallback activity.

These samples are bundled by the installer for quick verification and documentation.
