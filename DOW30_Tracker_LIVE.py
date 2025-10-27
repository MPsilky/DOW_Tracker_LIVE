# -*- coding: utf-8 -*-
from __future__ import annotations
"""
DOW 30 Tracker — market-hours only (no after-hours)
- Single instance (second run focuses the first)
- Backfill full day on launch; live capture during session
- Buckets: 9:31 AM, 10:00 AM, 11:00 AM, 12 NOON, 1:00 PM, 2:00 PM, 3:00 PM, 4:00 PM
- Arrows / % / colors are always vs the previous bucket (9:31 vs yesterday's close)
- Excel export: daily file "Sheet__MM_DD_YYYY.xlsx"
  * Grid sheet mirrors the UI (tickers + all buckets)
  * One sheet per bucket with numeric `price` and `pct` columns
- User-selectable data folder (remembered in settings.json)
"""

import os, sys, json, socket, threading, math, traceback, re
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, Optional, List, Tuple, Callable, cast

# -------- optional deps (guarded) --------
try:
    import yfinance as yf   # type: ignore
    import pandas as pd     # type: ignore
    import numpy as np      # type: ignore
except Exception:
    yf = None   # type: ignore[assignment]
    pd = None   # type: ignore[assignment]
    np = None   # type: ignore[assignment]
try:
    import openpyxl  # noqa: F401
except Exception:
    openpyxl = None  # type: ignore[assignment]

# -------- Qt --------
from PyQt5.QtCore import Qt, QTimer, QSize, QUrl
from PyQt5.QtGui import QIcon, QColor, QBrush, QCloseEvent, QFont, QDesktopServices
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QToolBar, QAction,
    QFileDialog, QMessageBox, QSystemTrayIcon, QMenu, QDialog, QVBoxLayout,
    QHBoxLayout, QLabel, QCheckBox, QPushButton, QStyle, QHeaderView, QTextEdit
)

APP_NAME = "DOW 30 Tracker"
APP_PORT = 53121  # single-instance IPC

# Order; rows 31/32 are blue
TICKERS: List[str] = [
    "AAPL","AMGN","AXP","BA","CAT","CRM","CSCO","CVX","DIS","DOW","GS","HD",
    "HON","IBM","JNJ","JPM","KO","MCD","MMM","MRK","MSFT","NKE","PG","TRV",
    "UNH","V","VZ","WMT","NVDA","AMZN","INTC","WBA"
]

# Buckets (no After Hours)
BUCKETS: List[Tuple[str, Tuple[int,int]]] = [
    ("9:31 AM", (9,31)),
    ("10:00 AM", (10,0)),
    ("11:00 AM", (11,0)),
    ("12 NOON", (12,0)),
    ("1:00 PM", (13,0)),
    ("2:00 PM", (14,0)),
    ("3:00 PM", (15,0)),
    ("4:00 PM", (16,0)),
]
TIME_COLS: List[str] = [b[0] for b in BUCKETS]
LABEL_TO_INDEX = {label:i for i,(label,_) in enumerate(BUCKETS)}
INDEX_TO_LABEL = {i:label for i,(label,_) in enumerate(BUCKETS)}
MARKET_OPEN = (9,31)
MARKET_CLOSE = (16,0)

# -------- paths & settings --------
def resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)

def exe_path() -> str:
    return sys.executable if getattr(sys, "frozen", False) else os.path.abspath(sys.argv[0])

SETTINGS_JSON = Path.home() / "AppData" / "Local" / "DOW30Tracker" / "settings.json"
SETTINGS_JSON.parent.mkdir(parents=True, exist_ok=True)

def load_settings() -> Dict[str, Any]:
    try:
        if SETTINGS_JSON.exists():
            return json.loads(SETTINGS_JSON.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def save_settings(s: Dict[str, Any]) -> None:
    SETTINGS_JSON.parent.mkdir(parents=True, exist_ok=True)
    SETTINGS_JSON.write_text(json.dumps(s, indent=2), encoding="utf-8")

def default_data_dir() -> Path:
    try:
        base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
        data = base / "data"
        data.mkdir(exist_ok=True)
        return data
    except Exception:
        d = Path.home() / "Documents" / "Saved DOW Sheets"
        d.mkdir(parents=True, exist_ok=True)
        return d

def get_data_dir() -> Path:
    s = load_settings()
    raw = s.get("data_dir")
    d = Path(raw) if isinstance(raw, str) and raw else default_data_dir()
    try: d.mkdir(parents=True, exist_ok=True)
    except Exception: d = default_data_dir()
    return d

DATA_DIR = get_data_dir()
LOG_PATH = DATA_DIR / "tracker.log"
FEATURES_JSON = DATA_DIR / "features.json"

def log(msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass

# -------- IPC: single instance --------
def start_primary_socket() -> socket.socket:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        s.bind(("127.0.0.1", APP_PORT))
        s.listen(1)
        return s
    except OSError:
        try:
            with socket.create_connection(("127.0.0.1", APP_PORT), timeout=1.5) as c:
                c.sendall(b"SHOW")
        except Exception:
            pass
        sys.exit(0)

def socket_listener(sock: socket.socket, on_show: Callable[[], None]) -> None:
    def _loop() -> None:
        try:
            while True:
                conn, _ = sock.accept()
                data = conn.recv(32)
                conn.close()
                if data and data.strip() == b"SHOW":
                    on_show()
        except Exception as e:
            log(f"ipc listener error: {e}")
    threading.Thread(target=_loop, daemon=True).start()

# -------- features (the toggles remain; explainer copy updated) --------
FEATURE_ITEMS: List[Tuple[str,str]] = [
    ("Mini “Dow Dashboard” with Market Mood", "Breadth, best/worst, concentration."),
    ("Auto News Ping for Movers", "Tray balloon for biggest mover + headline."),
    ("Chart Sparkline View", "Unicode mini-sparkline from intraday 1m data."),
    ("Replay Mode", "Play back saved captures."),
    ("Morning Resume Summary", "Toast with yesterday best/worst at start."),
    ("Historical Echo Mode", "On ±2% intraday, show last similar move & next day."),
    ("Dow “Concentration” Meter", "Top-5 movers’ share of movement."),
    ("Custom Market Sounds", "Beep when |move| ≥ threshold."),
    ("Candle Ghosts", "Show ‘ghost:<yday close>’ baseline."),
    ("Daily Confidence Meter", "Breadth meter (advancers/30)."),
    ("“Dow DNA” Export", "Weekly correlations & vol to CSV."),
]

# -------- finance helpers and intraday caching --------
def _df_empty(x: Any) -> bool:
    try:
        import pandas as _pd
        return not isinstance(x, _pd.DataFrame) or cast("_pd.DataFrame", x).empty
    except Exception:
        return True

def _ny_tz() -> str:
    return "US/Eastern"

# Intraday cache
_intraday_cache: Dict[str, "pd.Series"] = {}

def _load_intraday_cache() -> None:
    """
    Populate `_intraday_cache` with 1‑minute close price series for all
    tickers for the current trading day.  A single call to
    `yfinance.download` fetches data efficiently.  If yfinance or
    pandas are unavailable, the function quietly returns.
    """
    global _intraday_cache
    if yf is None or pd is None:
        return
    try:
        df = yf.download(
            tickers=" ".join(TICKERS),
            period="1d",
            interval="1m",
            prepost=False,
            group_by="ticker",
            auto_adjust=False,
            progress=False,
            threads=True,
        )
        for tk in TICKERS:
            s: Optional[pd.Series] = None
            if isinstance(df, dict):
                sub = df.get(tk)
                if sub is not None and "Close" in sub:
                    s = sub["Close"]
            elif tk in getattr(df, "columns", []) or (
                isinstance(getattr(df, "columns", None), pd.MultiIndex)
                and tk in df.columns.get_level_values(0)
            ):
                try:
                    sub_df = df.xs(tk, level=0) if isinstance(df.columns, pd.MultiIndex) else df[tk]  # type: ignore[index]
                    s = sub_df["Close"]  # type: ignore[index]
                except Exception:
                    s = None
            if s is not None and not s.empty:
                if getattr(s.index, "tz", None) is None:
                    s = s.tz_localize(_ny_tz())  # type: ignore[attr-defined]
                else:
                    s = s.tz_convert(_ny_tz())   # type: ignore[attr-defined]
                _intraday_cache[tk] = s.dropna()
    except Exception as e:
        log(f"_load_intraday_cache err: {e}")

def series_for_day_1m(tk: str) -> Optional["pd.Series"]:
    """
    Return the cached 1‑minute close price series for the given ticker
    if available.  If the cache is empty or the ticker has not yet
    been cached, fall back to a fresh yfinance.Ticker() call.  This
    provides an escape hatch if the cache hasn’t been loaded.
    """
    if yf is None or pd is None:
        return None
    s = _intraday_cache.get(tk)
    if s is not None:
        return s
    try:
        df = yf.Ticker(tk).history(period="1d", interval="1m", prepost=False)
        if _df_empty(df):
            return None
        ser = df["Close"]
        if getattr(ser.index, "tz", None) is None:
            ser = ser.tz_localize(_ny_tz())  # type: ignore[attr-defined]
        else:
            ser = ser.tz_convert(_ny_tz())   # type: ignore[attr-defined]
        return ser.dropna()
    except Exception as e:
        log(f"series_for_day_1m({tk}) err: {e}")
        return None

def price_at_or_before_bucket(tk: str, h: int, m: int) -> Optional[float]:
    if pd is None:
        return None
    s = series_for_day_1m(tk)
    if s is None or s.empty:
        return last_close(tk)
    bkt = pd.Timestamp(datetime.now().year, datetime.now().month, datetime.now().day, h, m, tz=_ny_tz())
    try:
        v = s.asof(bkt)  # type: ignore[attr-defined]
        if v is None and len(s) > 0:
            ss = s.loc[:bkt]
            return float(ss.iloc[-1]) if not ss.empty else last_close(tk)
        return float(v) if v is not None else last_close(tk)
    except Exception:
        ss = s.loc[:bkt]
        return float(ss.iloc[-1]) if not ss.empty else last_close(tk)

def yf_hist(tickers: List[str], period: str, interval: str):
    if yf is None: return None
    try:
        return yf.download(
            tickers=" ".join(tickers),
            period=period, interval=interval,
            group_by="ticker", auto_adjust=False, progress=False, threads=True
        )
    except Exception as e:
        log(f"yfinance download err: {e}"); return None

def last_close(tk: str) -> Optional[float]:
    if yf is None: return None
    try:
        h: Any = yf.Ticker(tk).history(period="5d", interval="1d")
        if _df_empty(h): return None
        return float(cast("pd.DataFrame", h)["Close"].iloc[-1])
    except Exception as e:
        log(f"last_close({tk}) err: {e}"); return None

# -------- UI helpers --------
GREEN = QColor(16,124,16)
RED   = QColor(180,32,32)
NEUT  = QColor(60,60,60)

def make_cell(text: str, pct: Optional[float]) -> QTableWidgetItem:
    it = QTableWidgetItem(text)
    col = GREEN if (isinstance(pct,(int,float)) and pct>0) else RED if (isinstance(pct,(int,float)) and pct<0) else NEUT
    it.setForeground(QBrush(col))
    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
    return it

def safe_sheet_name(label: str) -> str:
    s = label.replace(":", " ").replace("\\","_").replace("/","_").replace("?","_").replace("*","_").replace("[","(").replace("]",")")
    s = s.strip() or "Capture"
    return s[:31]

def day_file_name(dt: datetime) -> str:
    return f"Sheet__{dt.strftime('%m_%d_%Y')}.xlsx"

_price_re = re.compile(r"[-+]?\d+(?:\.\d+)?")
_pct_re   = re.compile(r"\(([+-]?\d+(?:\.\d+)?)%")

def parse_price_pct(cell_text: str) -> Tuple[Optional[float], Optional[float]]:
    s = cell_text.strip()
    s = s.lstrip("▲▼• ").replace("\xa0", " ")
    m1 = _price_re.search(s)
    px = float(m1.group(0)) if m1 else None
    m2 = _pct_re.search(s)
    pct = float(m2.group(1)) if m2 else None
    return px, pct

# -------- Main Window --------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME); self.setWindowIcon(self._icon()); self.resize(1400, 860)

        self.features = self._load_features()
        self.last_pct: Dict[str, float] = {}

        # table
        self.table = QTableWidget(len(TICKERS), 1 + len(TIME_COLS))
        self.table.setHorizontalHeaderLabels(["Ticker"] + TIME_COLS)
        self.table.verticalHeader().setVisible(False)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeToContents)

        # bigger UI
        base_font = QFont(); base_font.setPointSize(12)
        self.setFont(base_font)
        self.table.setFont(base_font)
        self.table.setStyleSheet("""
            QTableWidget { alternate-background-color:#fafafa; font-size:13px; }
            QHeaderView::section { padding:8px 12px; font-size:13px; }
        """)
        self.table.setAlternatingRowColors(True)
        for r in range(len(TICKERS)):
            self.table.setRowHeight(r, 28)

        fbold = QFont(); fbold.setPointSize(12); fbold.setBold(True)
        for i, tk in enumerate(TICKERS, start=1):
            it = QTableWidgetItem(f"{i}. {tk}")
            it.setFont(fbold)
            it.setForeground(QBrush(QColor("blue") if tk in ("INTC","WBA") else QColor("black")))
            it.setFlags(it.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(i-1, 0, it)
        self.setCentralWidget(self.table)

        # toolbar
        tb = QToolBar("Main"); tb.setIconSize(QSize(22,22)); tb.setMovable(False)
        tb.setStyleSheet("QToolBar{padding:8px; font-size:13px;} QToolButton{padding:8px 12px;}")
        ftb = QFont(); ftb.setPointSize(12); ftb.setBold(True); tb.setFont(ftb)
        self.addToolBar(Qt.TopToolBarArea, tb)

        self.actRefresh = QAction("Refresh", self); self.actRefresh.triggered.connect(self.refresh_now); tb.addAction(self.actRefresh)
        self.actBrowse  = QAction("Browse Excels", self); self.actBrowse.triggered.connect(self.browse_excels); tb.addAction(self.actBrowse)

        tb.addSeparator()
        self.actAutoOn  = QAction("Enable Auto-Start on Login", self);  self.actAutoOn.triggered.connect(self.enable_autostart); tb.addAction(self.actAutoOn)
        self.actAutoOff = QAction("Disable Auto-Start on Login", self); self.actAutoOff.triggered.connect(self.disable_autostart); tb.addAction(self.actAutoOff)

        tb.addSeparator()
        self.optTimes = QAction("Show Times", self, checkable=True, checked=True)
        self.optPct   = QAction("Show %",     self, checkable=True, checked=True)
        self.optArw   = QAction("Arrows",     self, checkable=True, checked=True)
        self.optStripe= QAction("Stripe Rows",self, checkable=True, checked=True)
        for a in (self.optTimes, self.optPct, self.optArw, self.optStripe): tb.addAction(a)

        tb.addSeparator()
        self.actFeatures = QAction("Features", self); self.actFeatures.triggered.connect(self.open_features); tb.addAction(self.actFeatures)
        self.actGuide    = QAction("Where the hell is the thing?", self); self.actGuide.triggered.connect(self.show_guide); tb.addAction(self.actGuide)

        tb.addSeparator()
        self.actSetFolder = QAction("Set Data Folder…", self); self.actSetFolder.triggered.connect(self.set_data_folder); tb.addAction(self.actSetFolder)

        # tray
        self.tray = QSystemTrayIcon(self._icon(), self)
        menu = QMenu()
        menu.addAction("Show").triggered.connect(self.show_main)
        menu.addAction("Exit").triggered.connect(self.exit_app)
        self.tray.setContextMenu(menu); self.tray.setToolTip(APP_NAME); self.tray.show()
        self.tray.activated.connect(
            lambda reason: self.show_main() if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick) else None
        )

        app = QApplication.instance()
        if app: app.setQuitOnLastWindowClosed(False)

        self.live_timer = QTimer(self); self.live_timer.timeout.connect(self.timer_logic); self.live_timer.start(60_000)
        QTimer.singleShot(60, self.initial_sync)

        self.statusBar().showMessage("Ready.")

    # ---------- core rendering ----------
    def _render_cell(self, price: Optional[float], base: Optional[float]) -> Tuple[str, Optional[float], QTableWidgetItem]:
        pct: Optional[float] = None
        if isinstance(price, (int, float)) and isinstance(base, (int, float)) and base:
            pct = ((price / base) - 1.0) * 100.0

        text = "--" if not isinstance(price, (int, float)) else f"{price:.2f}"
        if self.optPct.isChecked() and isinstance(pct, (int, float)):
            text += f"  ({pct:+.2f}%)"
        if self.optArw.isChecked() and isinstance(pct, (int, float)):
            text = ("▲ " if pct > 0 else ("▼ " if pct < 0 else "• ")) + text

        it = QTableWidgetItem(text)
        col = GREEN if (isinstance(pct,(int,float)) and pct>0) else RED if (isinstance(pct,(int,float)) and pct<0) else NEUT
        it.setForeground(QBrush(col))
        it.setFlags(it.flags() & ~Qt.ItemIsEditable)
        return text, pct, it

    # ---------- settings / features ----------
    def _load_features(self) -> Dict[str,bool]:
        if FEATURES_JSON.exists():
            try: return json.loads(FEATURES_JSON.read_text(encoding="utf-8"))
            except Exception: pass
        return {label: False for (label,_) in FEATURE_ITEMS}

    def _save_features(self) -> None:
        try: FEATURES_JSON.write_text(json.dumps(self.features, indent=2), encoding="utf-8")
        except Exception as e: log(f"save-features error: {e}")

    def set_data_folder(self) -> None:
        global DATA_DIR, LOG_PATH, FEATURES_JSON
        d = QFileDialog.getExistingDirectory(self, "Choose data folder for Excel & outputs", str(DATA_DIR))
        if d:
            s = load_settings(); s["data_dir"] = d; save_settings(s)
            DATA_DIR = Path(d)
            LOG_PATH = DATA_DIR / "tracker.log"
            FEATURES_JSON = DATA_DIR / "features.json"
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            self.tray.showMessage(APP_NAME, f"Data folder set to:\n{d}", QSystemTrayIcon.Information, 2500)

    # ---------- app logic ----------
    def initial_sync(self) -> None:
        """
        Perform the first data population when the application starts.  We
        synchronously load the intraday cache and then backfill the grid
        through the current bucket and capture it.  If any exception
        occurs during loading or rendering, it is logged but does not
        crash the UI.
        """
        try:
            _load_intraday_cache()
            self.backfill_to_now()
            self.refresh_now()
        except Exception as e:
            log(f"initial_sync err: {e}")

    def timer_logic(self) -> None:
        h, m = datetime.now().hour, datetime.now().minute
        if (MARKET_OPEN <= (h,m) <= MARKET_CLOSE):
            self.refresh_now()
        else:
            self.statusBar().showMessage("Outside market hours — next capture at next session open.")

    def browse_excels(self) -> None:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        path, _ = QFileDialog.getOpenFileName(self, "Open captured workbook", str(DATA_DIR), "Excel (*.xlsx)")
        try:
            if path:
                if sys.platform.startswith("win"): os.startfile(path)  # type: ignore[attr-defined]
                else: QDesktopServices.openUrl(QUrl.fromLocalFile(path))
            else:
                if sys.platform.startswith("win"): os.startfile(str(DATA_DIR))  # type: ignore[attr-defined]
                else: QDesktopServices.openUrl(QUrl.fromLocalFile(str(DATA_DIR)))
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Could not open: {e}")

    # ---------- capture ----------
    def _bucket_index_now(self) -> Optional[int]:
        hm = (datetime.now().hour, datetime.now().minute)
        idx = -1
        for i,(_, (h,m)) in enumerate(BUCKETS):
            if hm >= (h,m): idx = i
        return idx if idx >= 0 else None

    def backfill_to_now(self) -> None:
        if pd is None: return
        try:
            idx_now = self._bucket_index_now()
            if idx_now is None: return

            prev_closes = {tk: last_close(tk) for tk in TICKERS}

            for b in range(0, idx_now + 1):
                col = 1 + b
                h, m = BUCKETS[b][1]
                for r, tk in enumerate(TICKERS):
                    price = price_at_or_before_bucket(tk, h, m)

                    if b == 0:
                        base = prev_closes.get(tk)
                    else:
                        base = None
                        prev_item = self.table.item(r, col - 1)
                        if prev_item and prev_item.text():
                            try:
                                base = parse_price_pct(prev_item.text())[0]
                            except Exception:
                                base = None
                        if base is None:
                            base = prev_closes.get(tk)

                    _txt, pct, item = self._render_cell(price, base)
                    self.table.setItem(r, col, item)
                    if isinstance(pct, (int, float)):
                        self.last_pct[tk] = pct

            self.statusBar().showMessage(f"Backfilled through {INDEX_TO_LABEL[idx_now]}")
            self._export_excel_grid_and_bucket(INDEX_TO_LABEL[idx_now])

        except Exception as e:
            log(f"backfill_to_now err: {e}")

    def refresh_now(self) -> None:
        """
        Capture the current bucket.  Before computing new values, we
        refresh the intraday cache so that price_at_or_before_bucket()
        uses up‑to‑date minute data.  This call is synchronous and may
        take a few seconds, but it eliminates stale data between
        buckets.  Any exceptions are logged and surfaced via a
        message box.
        """
        try:
            _load_intraday_cache()

            idx = self._bucket_index_now()
            label = INDEX_TO_LABEL.get(idx) if idx is not None else None
            if label is None:
                self.statusBar().showMessage("Pre-market — will capture starting 9:31 AM", 4000)
                return

            col = 1 + LABEL_TO_INDEX[label]
            prev_col = col - 1

            for r, tk in enumerate(TICKERS):
                h, m = BUCKETS[LABEL_TO_INDEX[label]][1]
                price = price_at_or_before_bucket(tk, h, m)

                base = None
                if prev_col >= 1:
                    prev_it = self.table.item(r, prev_col)
                    if prev_it and prev_it.text():
                        base = parse_price_pct(prev_it.text())[0]
                if base is None:
                    base = last_close(tk)

                _txt, pct, item = self._render_cell(price, base)
                self.table.setItem(r, col, item)
                if isinstance(pct, (int, float)):
                    self.last_pct[tk] = pct

            self.statusBar().showMessage(f"Captured {label}")
            self._export_excel_grid_and_bucket(label)

        except Exception:
            log("refresh_now fatal:\n" + traceback.format_exc())
            QMessageBox.warning(self, APP_NAME, "Refresh failed. See tracker.log for details.")

    # ---------- Excel export (Grid + bucket) ----------
    def _excel_color_for(self, pct: Optional[float]) -> str:
        if isinstance(pct,(int,float)):
            if pct > 0:  return "107C10"
            if pct < 0:  return "B42020"
        return "505050"

    def _export_excel_grid_and_bucket(self, label: str) -> None:
        if pd is None or openpyxl is None:
            return
        try:
            day = datetime.now()
            fn = DATA_DIR / day_file_name(day)
            grid_name = "Grid"
            bucket_name = safe_sheet_name(label)

            headers = ["Ticker"] + TIME_COLS
            rows: List[List[str]] = []
            for r, tk in enumerate(TICKERS):
                row = [f"{r+1}. {tk}"]
                for c in range(1, 1 + len(TIME_COLS)):
                    it = self.table.item(r, c)
                    row.append(it.text() if it else "")
                rows.append(row)
            df_grid = pd.DataFrame(rows, columns=headers)

            col_idx = 1 + LABEL_TO_INDEX[label]
            prices: List[Optional[float]] = []
            pcts: List[Optional[float]] = []
            for r in range(len(TICKERS)):
                it = self.table.item(r, col_idx)
                px, pc = parse_price_pct(it.text() if it else "")
                prices.append(px); pcts.append(pc)
            df_bucket = pd.DataFrame({"price": prices, "pct": pcts},
                                     index=[f"{i+1}. {t}" for i,t in enumerate(TICKERS)])

            mode = "a" if fn.exists() else "w"
            with pd.ExcelWriter(fn, engine="openpyxl", mode=mode) as xw:
                try:
                    wb = xw.book  # type: ignore[attr-defined]
                    if grid_name in wb.sheetnames:
                        wb.remove(wb[grid_name])
                except Exception:
                    pass
                df_grid.to_excel(xw, sheet_name=grid_name, index=False)

                try:
                    wb = xw.book  # type: ignore[attr-defined]
                    if bucket_name in wb.sheetnames:
                        wb.remove(wb[bucket_name])
                except Exception:
                    pass
                df_bucket.to_excel(xw, sheet_name=bucket_name)

            from openpyxl import load_workbook
            from openpyxl.styles import PatternFill, Font

            wb = load_workbook(fn)
            ws = wb["Grid"]
            for cell in next(ws.iter_rows(min_row=1, max_row=1)):
                cell.font = Font(bold=True)
            for r, tk in enumerate(TICKERS, start=2):
                c = ws.cell(row=r, column=1)
                c.font = Font(bold=True, color="000000FF" if tk in ("INTC","WBA") else "00000000")
            for r in range(2, 2 + len(TICKERS)):
                for c in range(2, 2 + len(TIME_COLS) + 0):
                    val = ws.cell(row=r, column=c).value
                    if isinstance(val, str):
                        _, p = parse_price_pct(val)
                        fill = PatternFill("solid", fgColor=self._excel_color_for(p))
                        ws.cell(row=r, column=c).fill = fill

            bws = wb[bucket_name]
            for r in range(2, 2 + len(TICKERS)):
                pval = bws.cell(row=r, column=3).value  # pct column
                try:
                    p = float(pval) if pval not in (None,"") else None
                except Exception:
                    p = None
                fill = PatternFill("solid", fgColor=self._excel_color_for(p))
                bws.cell(row=r, column=2).fill = fill  # price col
                bws.cell(row=r, column=3).fill = fill  # pct col
            wb.save(fn)

        except Exception as e:
            log(f"excel export err: {e}")

    # ---------- Features ----------
    def open_features(self) -> None:
        dlg = QDialog(self); dlg.setWindowTitle("Features"); dlg.setWindowIcon(self._icon())
        v = QVBoxLayout(dlg)
        self.boxes: Dict[str,QCheckBox] = {}
        for label, tip in FEATURE_ITEMS:
            h = QHBoxLayout()
            cb = QCheckBox(label); cb.setChecked(self.features.get(label, False)); cb.setToolTip(tip)
            h.addWidget(cb, 1)
            info = QLabel("ⓘ"); info.setToolTip(tip); info.setStyleSheet("color:#3b82f6; font-weight:600;"); h.addWidget(info, 0, Qt.AlignRight)
            v.addLayout(h)
            self.boxes[label] = cb

        row = QHBoxLayout(); btnC = QPushButton("Cancel"); btnS = QPushButton("Save"); row.addStretch(1); row.addWidget(btnC); row.addWidget(btnS)
        v.addSpacing(6); v.addLayout(row)

        help_row = QHBoxLayout()
        help_btn = QPushButton("Click this if you are confused about a feature above")
        help_btn.setStyleSheet("QPushButton{padding:8px 10px; font-weight:600;}")
        help_btn.clicked.connect(self._open_feature_explainer)
        help_row.addStretch(1); help_row.addWidget(help_btn)
        v.addSpacing(8); v.addLayout(help_row)

        btnC.clicked.connect(dlg.reject); btnS.clicked.connect(dlg.accept)
        if dlg.exec_() == QDialog.Accepted:
            for k, cb in self.boxes.items(): self.features[k] = cb.isChecked()
            self._save_features(); self.statusBar().showMessage("Features updated.", 1500)

    def _open_feature_explainer(self) -> None:
        html = """
        <h2>Feature Explanations</h2>
        <ol>
          <li><b>Mini “Dow Dashboard” with Market Mood</b> — breadth, best/worst, and concentration for quick context.</li>
          <li><b>Auto News Ping for Movers</b> — tray alert when a component leads the move; shows headline if available.</li>
          <li><b>Chart Sparkline View</b> — tiny trend spark from 1-minute data for slope at a glance.</li>
          <li><b>Replay Mode</b> — scrub the day from saved buckets for post-mortems.</li>
          <li><b>Morning Resume Summary</b> — yesterday’s best/worst at start for context.</li>
          <li><b>Historical Echo Mode</b> — on ±2% move, surface the closest analog and next-day result.</li>
          <li><b>Dow “Concentration” Meter</b> — share of the move from the top five names.</li>
          <li><b>Custom Market Sounds</b> — tone when |move| ≥ your threshold.</li>
          <li><b>Candle Ghosts</b> — print `ghost:&lt;yday close&gt;` as a quick anchor.</li>
          <li><b>Daily Confidence Meter</b> — smoothed advancers/30 gauge.</li>
          <li><b>“Dow DNA” Export</b> — weekly correlations and realized vol to CSV.</li>
        </ol>
        """
        dlg = QDialog(self); dlg.setWindowTitle("Feature Explainer"); dlg.setWindowIcon(self._icon())
        lay = QVBoxLayout(dlg)
        te = QTextEdit(); te.setReadOnly(True); te.setHtml(html)
        lay.addWidget(te)
        btn = QPushButton("Close"); btn.clicked.connect(dlg.close); lay.addWidget(btn)
        dlg.resize(720, 560)
        dlg.exec_()

    # ---------- guide ----------
    def show_guide(self) -> None:
        dlg = QDialog(self)
        dlg.setWindowTitle("Where the hell is the thing?")
        dlg.setWindowIcon(self._icon())
        layout = QVBoxLayout(dlg)

        html = """
        <h2>DOW 30 Tracker — Quick Tour</h2>
        <ul>
          <li><b>Toolbar:</b> Refresh · Browse Excels · Auto-Start toggles · view options (Show Times, Show %, Arrows, Stripe Rows) · Features · this guide.</li>
          <li><b>Table:</b> Tickers on the left (1–32); columns are market-hour buckets. Each cell = price vs previous bucket (9:31 compares to yesterday’s close).</li>
          <li><b>Status bar:</b> Messages like “Captured 3:00 PM”.</li>
          <li><b>System tray:</b> Close hides to tray; launching app again brings this window forward.</li>
          <li><b>Excel:</b> Every capture writes to <code>Sheet__MM_DD_YYYY.xlsx</code> in your data folder (Grid + one sheet per bucket).</li>
        </ul>
        <p>Use “Set Data Folder…” to choose where workbooks are saved. Remembered in <code>settings.json</code>.</p>
        """
        te = QTextEdit(); te.setReadOnly(True); te.setHtml(html)
        layout.addWidget(te)
        btn = QPushButton("Got it"); btn.clicked.connect(dlg.accept); layout.addWidget(btn)
        dlg.resize(720, 560)
        dlg.exec_()

    # ---------- tray / window ----------
    def show_main(self) -> None:
        self.showNormal(); self.activateWindow(); self.raise_()

    def exit_app(self) -> None:
        QApplication.quit()

    def closeEvent(self, e: QCloseEvent) -> None:
        e.ignore(); self.hide()
        self.tray.showMessage(APP_NAME, "Still running in tray. Launch again to bring the window forward.", QSystemTrayIcon.Information, 1800)

    # ---------- autostart ----------
    def enable_autostart(self) -> None:
        if not sys.platform.startswith("win"):
            QMessageBox.information(self, APP_NAME, "Auto-start is supported on Windows only."); return
        try:
            import winreg as wr  # type: ignore
            with wr.OpenKey(wr.HKEY_CURRENT_USER, r"Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, wr.KEY_SET_VALUE) as k:
                wr.SetValueEx(k, APP_NAME, 0, wr.REG_SZ, exe_path())
            self.tray.showMessage(APP_NAME, "Auto-start enabled.", QSystemTrayIcon.Information, 1500)
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Failed to enable auto-start:\\n{e}")

    def disable_autostart(self) -> None:
        if not sys.platform.startswith("win"):
            QMessageBox.information(self, APP_NAME, "Auto-start is supported on Windows only."); return
        try:
            import winreg as wr  # type: ignore
            with wr.OpenKey(wr.HKEY_CURRENT_USER, r"Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, wr.KEY_SET_VALUE) as k:
                try: wr.DeleteValue(k, APP_NAME)
                except FileNotFoundError: pass
            self.tray.showMessage(APP_NAME, "Auto-start disabled.", QSystemTrayIcon.Information, 1500)
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Failed to disable auto-start:\\n{e}")

    # ---------- misc ----------
    def _icon(self) -> QIcon:
        for rel in ("assets/dow.ico","assets/dow.png"):
            p = resource_path(rel)
            if os.path.exists(p): return QIcon(p)
        inst = QApplication.instance()
        return inst.style().standardIcon(QStyle.SP_ComputerIcon) if inst else QIcon()

# -------- main --------
def main() -> None:
    sock = start_primary_socket()
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME); app.setQuitOnLastWindowClosed(False)
    w = MainWindow()
    socket_listener(sock, on_show=w.show_main)
    w.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"FATAL: {e}\\n{traceback.format_exc()}")
        raise
