# V1 Final — UI Unfreeze, Resilient Intraday, Guaranteed EOD Excel

## Added
- Background capture worker that runs `_load_intraday_cache`, `backfill_to_now`, and `refresh_now` away from the Qt event loop.
- Persisted intraday minute cache under `data/intraday_cache/` plus installer support for the directory.
- Async fallback mesh that layers disk cache with yfinance plus optional Alpha Vantage (`ALPHAVANTAGE_API_KEY`), Twelve Data (`TWELVEDATA_API_KEY`), IEX Cloud (`IEX_CLOUD_API_KEY`), Finnhub (`FINNHUB_API_KEY`), and Polygon (`POLYGON_API_KEY`).
- Forced 4:00 PM export path with `atexit` safety net and tray-exit hook.
- GitHub Actions workflow (`.github/workflows/v1-final.yml`) that builds, compiles, and stages artifacts.
- `samples/tracker_log_sample.txt` diagnostics for documentation/testing (mock Excel artifacts removed per shipping policy).

## Changed
- `DOW30_Tracker_LIVE.py` now tracks stale tickers, updates the UI via worker callbacks, highlights stale cells, and fans out provider fetches concurrently with async/await.
- `build.ps1` creates the intraday cache folder, installs the `requests` dependency, and bumps the version to 1.5.0.
- Inno Setup script packages the console build and samples, and references the Markdown README.
- README rewritten for V1 Final (run commands, fallback configuration, workflow usage).
