# Data directory

Runtime captures, Excel exports, and cached minute bars live here. The repository keeps the folder so the
packaged build can create sibling files without shipping placeholder `.gitkeep` entries or mock Excel workbooks.

`intraday_cache/` is populated at runtime with `{DATE}__{TICKER}.csv` snapshots written by `_persist_intraday_cache`.
