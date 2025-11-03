# V1 Final  UI Unfreeze, Resilient Intraday, Guaranteed EOD Excel

## Added
- Background capture worker that runs `_load_intraday_cache`, `backfill_to_now`, and `refresh_now` away from the Qt event loop.
- Persisted intraday minute cache under `data/intraday_cache/` plus installer support for the directory.
- Alpha Vantage fallback (optional `ALPHAVANTAGE_API_KEY`) and disk-cache recovery when yfinance is empty.
- Forced 4:00 PM export path with `atexit` safety net and tray-exit hook.
- GitHub Actions workflow (`.github/workflows/v1-final.yml`) that builds, compiles, and stages artifacts.
- `samples/Sheet__Sample.xlsx` workbook and `samples/tracker_log_sample.txt` diagnostics for documentation/testing.

## Changed
- `DOW30_Tracker_LIVE.py` now tracks stale tickers, updates the UI via worker callbacks, and highlights stale cells.
- `build.ps1` creates the intraday cache folder, installs the `requests` dependency, and bumps the version to 1.5.0.
- Inno Setup script packages the console build and samples, and references the Markdown README.
- README rewritten for V1 Final (run commands, fallback configuration, workflow usage).
