# -*- coding: utf-8 -*-
from __future__ import annotations
"""
DOW 30 Tracker — market-hours only (no after-hours)
- Single instance (second run focuses the first)
- Backfill full day on launch; live capture during session
- Buckets: 9:31 AM, 10:00 AM, 11:00 AM, 12 NOON, 1:00 PM, 2:00 PM, 3:00 PM, 4:00 PM
- Arrows / % / colors compare each bucket to the immediately prior bucket (9:31 uses the prior session's 4:00 PM capture, then
  falls back to the latest close if needed)
- Excel export: daily file "Sheet__MM_DD_YYYY.xlsx"
  * Grid sheet mirrors the UI (tickers + all buckets)
  * One sheet per bucket with numeric `price` and `pct` columns
- User-selectable data folder (remembered in settings.json)
"""

import os, sys, json, socket, threading, traceback, re, atexit, io
from pathlib import Path
from datetime import datetime, date
from typing import Any, Dict, Optional, List, Tuple, Callable, cast

try:
    from zoneinfo import ZoneInfo  # type: ignore
except Exception:
    ZoneInfo = None  # type: ignore[assignment]

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

if sys.platform.startswith("win"):
    local_appdata = os.environ.get("LOCALAPPDATA")
    if local_appdata:
        STATE_DIR = Path(local_appdata) / "DOW30Tracker"
    else:
        STATE_DIR = Path.home() / "AppData" / "Local" / "DOW30Tracker"
else:
    STATE_DIR = Path.home() / ".dow30tracker"
STATE_DIR.mkdir(parents=True, exist_ok=True)
SETTINGS_JSON = STATE_DIR / "settings.json"
INSTANCE_LOCK = STATE_DIR / "tracker.lock"
_LOCK_HANDLE: Optional["io.TextIOWrapper"] = None  # kept alive for the process lifetime

class _SingleInstanceError(RuntimeError):
    pass

def _acquire_lock() -> None:
    global _LOCK_HANDLE
    if _LOCK_HANDLE is not None:
        return
    try:
        fh = INSTANCE_LOCK.open("a+")
        if os.name == "nt":
            import msvcrt  # type: ignore
            try:
                msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError:
                fh.close()
                raise _SingleInstanceError
        else:
            import fcntl  # type: ignore
            try:
                fcntl.lockf(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError:
                fh.close()
                raise _SingleInstanceError
        _LOCK_HANDLE = fh
    except FileNotFoundError:
        STATE_DIR.mkdir(parents=True, exist_ok=True)
        _acquire_lock()

def _release_lock() -> None:
    global _LOCK_HANDLE
    if _LOCK_HANDLE is None:
        return
    try:
        if os.name == "nt":
            import msvcrt  # type: ignore
            try:
                _LOCK_HANDLE.seek(0)
                msvcrt.locking(_LOCK_HANDLE.fileno(), msvcrt.LK_UNLCK, 1)
            except OSError:
                pass
        else:
            import fcntl  # type: ignore
            try:
                fcntl.lockf(_LOCK_HANDLE.fileno(), fcntl.LOCK_UN)
            except OSError:
                pass
    finally:
        try:
            _LOCK_HANDLE.close()
        except Exception:
            pass
        _LOCK_HANDLE = None

atexit.register(_release_lock)

def _notify_existing_instance() -> None:
    try:
        with socket.create_connection(("127.0.0.1", APP_PORT), timeout=1.5) as c:
            c.sendall(b"SHOW")
    except Exception:
        pass

def ensure_single_instance() -> None:
    try:
        _acquire_lock()
    except _SingleInstanceError:
        _notify_existing_instance()
        sys.exit(0)

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
    doc_candidates: List[Path] = []
    home = Path.home()
    doc_candidates.append(home / "Documents" / "Saved DOW Sheets")
    if sys.platform.startswith("win"):
        # Honour OneDrive documents if present
        for env_key in ("OneDriveCommercial", "OneDriveConsumer", "ONEDRIVE", "OneDrive"):
            onedrive = os.environ.get(env_key)
            if onedrive:
                doc_candidates.insert(0, Path(onedrive) / "Documents" / "Saved DOW Sheets")
    for candidate in doc_candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            return candidate
        except Exception:
            continue
    fallback = STATE_DIR / "Saved"
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback

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

_WORKBOOK_RE = re.compile(r"Sheet__(\d{2})_(\d{2})_(\d{4})\.xlsx$", re.IGNORECASE)
_PRIOR_BUCKET_CACHE: Dict[Tuple[str, str], Dict[str, Optional[float]]]
_PRIOR_BUCKET_CACHE = {}


def _parse_workbook_date(path: Path) -> Optional[date]:
    m = _WORKBOOK_RE.match(path.name)
    if not m:
        return None
    try:
        month, day, year = map(int, m.groups())
        return date(year, month, day)
    except ValueError:
        return None


def _latest_workbook_before(session_day: date, data_dir: Path) -> Optional[Path]:
    best: Optional[Tuple[date, Path]] = None
    try:
        for path in data_dir.glob("Sheet__*.xlsx"):
            dt = _parse_workbook_date(path)
            if dt is None or dt >= session_day:
                continue
            if best is None or dt > best[0]:
                best = (dt, path)
    except Exception as e:
        log(f"scan workbooks err: {e}")
        return None
    return best[1] if best else None


def _load_bucket_from_workbook(path: Path, bucket_label: str) -> Dict[str, Optional[float]]:
    cache_key = (str(path.resolve()), bucket_label)
    cached = _PRIOR_BUCKET_CACHE.get(cache_key)
    if cached is not None:
        return dict(cached)
    if openpyxl is None:
        return {}
    try:
        from openpyxl import load_workbook

        sheet_name = safe_sheet_name(bucket_label)
        wb = load_workbook(path, data_only=True, read_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return {}
        ws = wb[sheet_name]
        result: Dict[str, Optional[float]] = {}
        for row_idx, tk in enumerate(TICKERS, start=2):
            val = ws.cell(row=row_idx, column=2).value
            result[tk] = float(val) if isinstance(val, (int, float)) else None
        wb.close()
        _PRIOR_BUCKET_CACHE[cache_key] = dict(result)
        return result
    except Exception as e:
        log(f"load workbook bucket err: {e}")
        return {}


def prior_session_bucket(session_day: date, bucket_label: str = "4:00 PM", data_dir: Optional[Path] = None) -> Dict[str, Optional[float]]:
    directory = data_dir if data_dir is not None else DATA_DIR
    path = _latest_workbook_before(session_day, directory)
    if path is None:
        return {}
    bucket = _load_bucket_from_workbook(path, bucket_label)
    if not bucket:
        log(f"no prior bucket data for {bucket_label} in {path.name}")
    return bucket


def clear_prior_bucket_cache() -> None:
    _PRIOR_BUCKET_CACHE.clear()

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
        _notify_existing_instance()
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
    ("Mini \"Dow Dashboard\" Pulse", "Bright banner showing advancers vs decliners, top movers, and an instant mood check."),
    ("Auto News Ping for Movers", "Tray balloon with the day's biggest mover plus its freshest headline."),
    ("Chart Sparkline View", "Unicode sparkline built from 1-minute closes so you can spot slope without leaving the grid."),
    ("Replay Mode", "Scrub through saved captures to relive the session or clip highlights."),
    ("Morning Resume Summary", "Morning toast recapping yesterday's leaders, laggards, and index finish."),
    ("Historical Echo Mode", "On ±2% swings, surface the closest historical analog and what happened next."),
    ("Dow \"Concentration\" Meter", "Gauge how much of the move the top five components are driving."),
    ("Custom Market Sounds", "Play chimes whenever a component crosses your chosen percentage threshold."),
    ("Candle Ghosts", "Overlay a ghost line at yesterday's close for quick anchoring."),
    ("Daily Confidence Meter", "Animated breadth bar that tracks advancers out of 30 throughout the day."),
    ("\"Dow DNA\" Export", "Weekly CSV drop with correlations, realized vol, and stat blocks."),
]

# -------- finance helpers and intraday caching --------
def _df_empty(x: Any) -> bool:
    try:
        import pandas as _pd
        return not isinstance(x, _pd.DataFrame) or cast("_pd.DataFrame", x).empty
    except Exception:
        return True

def _ny_tz() -> str:
    return "America/New_York"


def now_eastern() -> datetime:
    if ZoneInfo is not None:
        try:
            return datetime.now(tz=ZoneInfo("America/New_York"))
        except Exception:
            pass
    try:
        import pytz  # type: ignore

        return datetime.now(tz=pytz.timezone("US/Eastern"))
    except Exception:
        return datetime.now()


def format_et(dt: datetime) -> str:
    try:
        text = dt.strftime("%I:%M %p")
        if text.startswith("0"):
            text = text[1:]
        return f"{text} ET"
    except Exception:
        return dt.strftime("%H:%M ET")

# Intraday cache
_intraday_cache: Dict[str, "pd.Series"] = {}

def _load_intraday_cache() -> None:
    """
    Populate `_intraday_cache` with 1-minute close price series for all
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
        new_cache: Dict[str, "pd.Series"] = {}
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
                new_cache[tk] = s.dropna()
        if new_cache:
            _intraday_cache = new_cache
    except Exception as e:
        log(f"_load_intraday_cache err: {e}")

def series_for_day_1m(tk: str) -> Optional["pd.Series"]:
    """
    Return the cached 1-minute close price series for the given ticker
    if available.  If the cache is empty or the ticker has not yet
    been cached, fall back to a fresh yfinance.Ticker() call.  This
    provides an escape hatch if the cache has not been loaded.
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

def price_at_or_before_bucket(tk: str, session_day: date, h: int, m: int) -> Optional[float]:
    if pd is None:
        return None
    s = series_for_day_1m(tk)
    if s is None or s.empty:
        return last_close(tk)
    bkt = pd.Timestamp(session_day.year, session_day.month, session_day.day, h, m, tz=_ny_tz())
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
GREEN = QColor("#22c55e")
RED   = QColor("#f87171")
NEUT  = QColor("#e2e8f0")

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
        self.session_date = now_eastern().date()
        self.prev_close: Dict[str, Optional[float]] = {}
        self.prev_session_bucket: Dict[str, Optional[float]] = prior_session_bucket(self.session_date)
        self.bucket_prices: Dict[str, Dict[str, Optional[float]]] = {tk: {} for tk in TICKERS}
        self._exported_states: Dict[str, Tuple[Optional[float], ...]] = {}
        self._last_timer_key: Optional[Tuple[int, int]] = None

        # table
        self.table = QTableWidget(len(TICKERS), 1 + len(TIME_COLS))
        self.table.setHorizontalHeaderLabels(["Ticker"] + TIME_COLS)
        self.table.verticalHeader().setVisible(False)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeToContents)
        hh.setDefaultAlignment(Qt.AlignCenter)
        hh.setSectionResizeMode(0, QHeaderView.Fixed)

        base_font = QFont(); base_font.setPointSize(13)
        self.setFont(base_font)
        self.table.setFont(base_font)
        self.table.setFocusPolicy(Qt.NoFocus)
        self.table.setAlternatingRowColors(True)
        for r in range(len(TICKERS)):
            self.table.setRowHeight(r, 34)

        fbold = QFont(); fbold.setPointSize(13); fbold.setBold(True)
        first_col_bg = QBrush(QColor("#1e293b"))
        for i, tk in enumerate(TICKERS, start=1):
            it = QTableWidgetItem(f"{i}. {tk}")
            it.setFont(fbold)
            color = QColor("#38bdf8") if tk in ("INTC","WBA") else QColor("#f8fafc")
            it.setForeground(QBrush(color))
            it.setBackground(first_col_bg)
            it.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            it.setFlags(it.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(i-1, 0, it)
        self.table.setColumnWidth(0, 150)
        self.setCentralWidget(self.table)

        # toolbar
        tb = QToolBar("Main"); tb.setIconSize(QSize(22,22)); tb.setMovable(False)
        ftb = QFont(); ftb.setPointSize(12); ftb.setBold(True); tb.setFont(ftb)
        tb.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.addToolBar(Qt.TopToolBarArea, tb)

        self.setStyleSheet(
            "QMainWindow { background-color: #0f172a; }\n"
            "QToolBar { background: transparent; padding: 12px; spacing: 6px; }\n"
            "QToolButton { color: #f8fafc; padding: 8px 14px; border-radius: 6px; background-color: rgba(37, 99, 235, 0.18); }\n"
            "QToolButton:hover { background-color: #2563eb; color: #f8fafc; }\n"
            "QToolButton:checked { background-color: #1d4ed8; color: #f8fafc; }\n"
            "QTableWidget { background-color: #0b1220; color: #e2e8f0; gridline-color: #1f2a40; alternate-background-color: #131d33; selection-background-color: #2563eb; selection-color: #f8fafc; border: 1px solid #1f2a40; }\n"
            "QTableWidget::item { padding: 6px 10px; }\n"
            "QHeaderView::section { background-color: #1e293b; color: #f8fafc; font-size: 13px; padding: 10px 12px; border: 0; }\n"
            "QStatusBar { background: #111827; color: #f8fafc; font-weight: 600; padding: 4px 12px; }\n"
            "QMenu { background: #0f172a; color: #f8fafc; border: 1px solid #1f2a40; }\n"
            "QMenu::item:selected { background: #2563eb; }\n"
            "QDialog { background-color: #0f172a; color: #f1f5f9; }\n"
            "QLabel { color: #f1f5f9; }\n"
            "QTextEdit { background-color: #0b1220; color: #e2e8f0; border: 1px solid #1f2a40; }\n"
            "QPushButton { background-color: #1d4ed8; border-radius: 6px; padding: 8px 16px; color: #f8fafc; font-weight: 600; }\n"
            "QPushButton:hover { background-color: #2563eb; }\n"
            "QCheckBox { color: #f1f5f9; font-size: 13px; }"
        )

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
        self.optTimes.toggled.connect(self._toggle_times)
        self.optPct.toggled.connect(lambda _: self._rerender_table())
        self.optArw.toggled.connect(lambda _: self._rerender_table())
        self.optStripe.toggled.connect(self._toggle_stripes)

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

        status_font = QFont(); status_font.setPointSize(11); status_font.setBold(True)
        self.statusBar().setFont(status_font)
        self.statusBar().setSizeGripEnabled(False)
        self.statusBar().showMessage("Ready — waiting for the first capture at 9:31 AM ET.")

    # ---------- core rendering ----------
    def _reset_for_new_session(self, new_date: Optional[date] = None) -> None:
        self.session_date = new_date or now_eastern().date()
        self.prev_close = {}
        clear_prior_bucket_cache()
        self.prev_session_bucket = prior_session_bucket(self.session_date)
        self.bucket_prices = {tk: {} for tk in TICKERS}
        self._exported_states = {}
        self._last_timer_key = None
        for r in range(len(TICKERS)):
            for c in range(1, 1 + len(TIME_COLS)):
                blank = QTableWidgetItem("")
                blank.setFlags(blank.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(r, c, blank)

    def _ensure_session(self, current_date: Optional[date] = None) -> None:
        target = current_date or now_eastern().date()
        if target != self.session_date:
            self._reset_for_new_session(target)

    def _ensure_prev_close(self, tk: str) -> Optional[float]:
        if tk not in self.prev_close or self.prev_close[tk] is None:
            self.prev_close[tk] = last_close(tk)
        return self.prev_close.get(tk)

    def _base_for(self, tk: str, label: str) -> Optional[float]:
        idx = LABEL_TO_INDEX[label]
        if idx == 0:
            prior_val = self.prev_session_bucket.get(tk)
            if isinstance(prior_val, (int, float)):
                return prior_val
            return self._ensure_prev_close(tk)

        bucket_dict = self.bucket_prices.get(tk, {})
        for prev_idx in range(idx - 1, -1, -1):
            prev_label = INDEX_TO_LABEL[prev_idx]
            prev_val = bucket_dict.get(prev_label)
            if isinstance(prev_val, (int, float)):
                return prev_val

        return self._ensure_prev_close(tk)

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

    def _rerender_table(self) -> None:
        for label in TIME_COLS:
            col = 1 + LABEL_TO_INDEX[label]
            for row, tk in enumerate(TICKERS):
                price = self.bucket_prices.get(tk, {}).get(label)
                base = self._base_for(tk, label)
                _txt, _pct, item = self._render_cell(price, base)
                self.table.setItem(row, col, item)

    def _toggle_times(self, checked: bool) -> None:
        self.table.horizontalHeader().setVisible(checked)

    def _toggle_stripes(self, checked: bool) -> None:
        self.table.setAlternatingRowColors(checked)
        self.table.viewport().update()

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
            clear_prior_bucket_cache()
            self.prev_session_bucket = prior_session_bucket(self.session_date)
            self._rerender_table()
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
        now_dt = now_eastern()
        hm = (now_dt.hour, now_dt.minute)
        if self._last_timer_key == hm:
            return
        self._last_timer_key = hm
        if MARKET_OPEN <= hm <= MARKET_CLOSE:
            self.refresh_now()
        else:
            self.statusBar().showMessage("Outside market hours — next capture at next session open.", 4000)

    def browse_excels(self) -> None:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(DATA_DIR))  # type: ignore[attr-defined]
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(DATA_DIR)))
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Could not open: {e}")

    # ---------- capture ----------
    def _bucket_index_now(self, hm: Optional[Tuple[int, int]] = None) -> Optional[int]:
        hour_min = hm or (now_eastern().hour, now_eastern().minute)
        idx = -1
        for i,(_, (h,m)) in enumerate(BUCKETS):
            if hour_min >= (h,m): idx = i
        return idx if idx >= 0 else None

    def backfill_to_now(self) -> None:
        if pd is None: return
        try:
            now_dt = now_eastern()
            hm = (now_dt.hour, now_dt.minute)
            self._ensure_session(now_dt.date())
            idx_now = self._bucket_index_now(hm)
            if idx_now is None: return

            for tk in TICKERS:
                self._ensure_prev_close(tk)

            for b in range(0, idx_now + 1):
                col = 1 + b
                h, m = BUCKETS[b][1]
                label = INDEX_TO_LABEL[b]
                for r, tk in enumerate(TICKERS):
                    price = price_at_or_before_bucket(tk, self.session_date, h, m)
                    base = self._base_for(tk, label)
                    _txt, pct, item = self._render_cell(price, base)
                    self.table.setItem(r, col, item)
                    self.bucket_prices[tk][label] = price if isinstance(price, (int, float)) else None
                self._maybe_export(label)

            last_label = INDEX_TO_LABEL[idx_now]
            self.statusBar().showMessage(f"Backfilled through {last_label} @ {format_et(now_dt)}", 4000)

        except Exception as e:
            log(f"backfill_to_now err: {e}")

    def refresh_now(self) -> None:
        """
        Capture the current bucket.  Before computing new values, we
        refresh the intraday cache so that price_at_or_before_bucket()
        uses up-to-date minute data.  This call is synchronous and may
        take a few seconds, but it eliminates stale data between
        buckets.  Any exceptions are logged and surfaced via a
        message box.
        """
        try:
            _load_intraday_cache()

            now_dt = now_eastern()
            hm = (now_dt.hour, now_dt.minute)
            self._ensure_session(now_dt.date())
            idx = self._bucket_index_now(hm)
            label = INDEX_TO_LABEL.get(idx) if idx is not None else None
            if label is None:
                self.statusBar().showMessage("Pre-market — will capture starting 9:31 AM", 4000)
                return

            label_idx = LABEL_TO_INDEX[label]
            col = 1 + label_idx

            for r, tk in enumerate(TICKERS):
                h, m = BUCKETS[label_idx][1]
                price = price_at_or_before_bucket(tk, self.session_date, h, m)
                self._ensure_prev_close(tk)
                base = self._base_for(tk, label)
                _txt, pct, item = self._render_cell(price, base)
                self.table.setItem(r, col, item)
                self.bucket_prices[tk][label] = price if isinstance(price, (int, float)) else None

            self._maybe_export(label)
            self.statusBar().showMessage(f"Captured {label} @ {format_et(now_dt)}", 4000)

        except Exception:
            log("refresh_now fatal:\n" + traceback.format_exc())
            QMessageBox.warning(self, APP_NAME, "Refresh failed. See tracker.log for details.")

    # ---------- Excel export (Grid + bucket) ----------
    def _snapshot_for_label(self, label: str) -> Tuple[Optional[float], ...]:
        snap: List[Optional[float]] = []
        for tk in TICKERS:
            val = self.bucket_prices.get(tk, {}).get(label)
            snap.append(val if isinstance(val, (int, float)) else None)
        return tuple(snap)

    def _maybe_export(self, label: str, force: bool = False) -> None:
        if pd is None or openpyxl is None:
            return
        snap = self._snapshot_for_label(label)
        if force or self._exported_states.get(label) != snap:
            self._export_excel_grid_and_bucket(label)
            self._exported_states[label] = snap

    def _excel_color_for(self, pct: Optional[float]) -> str:
        if isinstance(pct,(int,float)):
            if pct > 0:  return "22C55E"
            if pct < 0:  return "F87171"
        return "94A3B8"

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

            prices: List[Optional[float]] = []
            pcts: List[Optional[float]] = []
            for tk in TICKERS:
                self._ensure_prev_close(tk)
                px = self.bucket_prices.get(tk, {}).get(label)
                base = self._base_for(tk, label)
                pct_val: Optional[float] = None
                if isinstance(px, (int, float)) and isinstance(base, (int, float)) and base:
                    pct_val = ((px / base) - 1.0) * 100.0
                prices.append(px if isinstance(px, (int, float)) else None)
                pcts.append(pct_val)
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
          <li><b>Mini \"Dow Dashboard\" Pulse</b> — bright banner showing advancers vs decliners, top movers, and a quick market mood dial.</li>
          <li><b>Auto News Ping for Movers</b> — tray alert that pairs the lead mover with its freshest Yahoo headline.</li>
          <li><b>Chart Sparkline View</b> — Unicode sparkline using 1-minute closes so you can see the slope without leaving the grid.</li>
          <li><b>Replay Mode</b> — scrub saved captures like a DVR to replay the session or clip highlights.</li>
          <li><b>Morning Resume Summary</b> — sunrise toast recapping yesterday's leaders, laggards, and index finish.</li>
          <li><b>Historical Echo Mode</b> — on ±2% swings, surface the closest historical analog and what happened next day.</li>
          <li><b>Dow \"Concentration\" Meter</b> — gauge showing how much of the move is driven by the top five components.</li>
          <li><b>Custom Market Sounds</b> — playful chimes whenever a component crosses the percentage threshold you set.</li>
          <li><b>Candle Ghosts</b> — overlay a ghost line at yesterday's close so you instantly know where price is anchored.</li>
          <li><b>Daily Confidence Meter</b> — animated breadth bar tracking advancers out of 30 as the session unfolds.</li>
          <li><b>\"Dow DNA\" Export</b> — weekly CSV drop full of correlations, realized volatility, and stat blocks.</li>
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
          <li><b>Table:</b> Tickers on the left (1–32); columns are market-hour buckets. Each cell = price vs the previous bucket (9:31 compares to the prior session's 4:00 PM capture, falling back to the latest close when needed).</li>
          <li><b>Status bar:</b> Messages like "Captured 3:00 PM".</li>
          <li><b>System tray:</b> Close hides to tray; launching app again brings this window forward.</li>
          <li><b>Excel:</b> Every capture writes to <code>Sheet__MM_DD_YYYY.xlsx</code> in your data folder (Grid + one sheet per bucket).</li>
        </ul>
        <p>Use "Set Data Folder…" to choose where workbooks are saved. Remembered in <code>settings.json</code>.</p>
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
    ensure_single_instance()
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
