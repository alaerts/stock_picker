#!/usr/bin/env python3
"""
Daily multi-index stock report — S&P 500, Nikkei 225, FTSE 100, DAX 40, CAC 40, BEL 20.

Outputs a single .xlsx with one row per stock containing:
  - Current price + prices 1d / 1w / 1mo / 6mo / 1y / 5y ago, all in EUR at TODAY's FX rate
  - Trailing P/E and Forward P/E
  - Native currency
  - Comma-separated list of Yahoo watchlists the stock appears in, from:
      Largest 52-Week Gains, Crowded Hedge Fund Positions, Berkshire Hathaway Portfolio,
      Smart Money Stocks, Most Sold by Activist Hedge Funds,
      Most Bought by Activist Hedge Funds, Most Bought by Hedge Funds,
      Activist Hedge Fund Positions

Sources: yfinance (Yahoo Finance), Wikipedia for index constituents, Yahoo watchlist pages.

USAGE
-----
First-time setup:
    pip install -r requirements.txt

Smoke test (BEL 20 only, ~20 tickers, finishes in ~1 min):
    python stocks_report.py --indexes BEL20

Full run (all indexes, ~925 tickers, 10–20 min):
    python stocks_report.py

Custom output path:
    python stocks_report.py --output ~/reports/stocks.xlsx

CAVEATS
-------
- yfinance is unofficial; Yahoo can break it. If a previously-working run fails,
  try: pip install --upgrade yfinance
- Watchlist scraping depends on Yahoo's HTML structure and may break silently.
  The report still completes; failed lists log a warning.
- Tickers younger than 5y will have a blank "5Y ago" cell — expected.
- FX conversion uses TODAY's rate for all historical prices, as requested.
  This mixes price movement with FX movement for non-EUR stocks.
"""

from __future__ import annotations

import argparse
import datetime as dt
import io
import logging
import re
import sys
import time
from pathlib import Path
from typing import Optional

import pandas as pd
import requests
import yfinance as yf
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("stocks_report")

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HTTP_HEADERS = {"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"}

# Wikipedia URLs per index. Constituent table is auto-detected by column names.
INDEX_WIKI = {
    "SP500":     "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies",
    "NIKKEI225": "https://en.wikipedia.org/wiki/Nikkei_225",
    "FTSE100":   "https://en.wikipedia.org/wiki/FTSE_100_Index",
    "DAX":       "https://en.wikipedia.org/wiki/DAX",
    "CAC40":     "https://en.wikipedia.org/wiki/CAC_40",
    "BEL20":     "https://en.wikipedia.org/wiki/BEL_20",
    "ESTOXX50":  "https://en.wikipedia.org/wiki/EURO_STOXX_50",
}

# Native currency fallback if yfinance .info doesn't give us one.
INDEX_DEFAULT_CCY = {
    "SP500": "USD", "NIKKEI225": "JPY", "FTSE100": "GBP",
    "DAX": "EUR", "CAC40": "EUR", "BEL20": "EUR", "ESTOXX50": "EUR",
}

# Watchlists: (display label, source, reference).
# Sources:
#   "yahoo_screener"     → ref is a Yahoo predefined screener scrIds value.
#   "dataroma_url"       → ref is a dataroma path (e.g. "/m/g/portfolio.php").
#   "dataroma_activists" → aggregate across DATAROMA_ACTIVIST_CODES; ref ignored.
#
# Yahoo deprecated /u/yahoo-finance/watchlists/ pages in 2026; the original
# 8-list spec from CLAUDE.md is mostly gone. See MEMORY.md "2026-05-06" entries
# for the migration path. This is the slim 6-list replacement (Option B).
WATCHLISTS: list[tuple[str, str, str]] = [
    ("Recent 52-Week Highs",                  "yahoo_screener",     "recent_52_week_highs"),
    ("Berkshire Hathaway Portfolio",          "yahoo_screener",     "top_stocks_owned_by_warren_buffet"),
    ("Top Quarterly Buys (Super Investors)",  "dataroma_url",       "/m/g/portfolio_b.php?q=q&o=c"),
    ("Top Quarterly Sells (Super Investors)", "dataroma_url",       "/m/g/portfolio_s.php?q=q&o=c"),
    ("Most-Held by Super Investors",          "dataroma_url",       "/m/g/portfolio.php"),
    ("Activist Hedge Fund Positions",         "dataroma_activists", ""),
]

# Dataroma manager codes for the 7 activist funds we aggregate. Codes are
# case-sensitive in the URL ?m=... param.
DATAROMA_ACTIVIST_CODES = [
    "ic",   # Carl Icahn / Icahn Capital
    "psc",  # Bill Ackman / Pershing Square
    "VA",   # ValueAct Capital
    "TF",   # Nelson Peltz / Trian Fund Management
    "ENG",  # Glenn Welling / Engaged Capital
    "tp",   # Daniel Loeb / Third Point
    "tci",  # Christopher Hohn / TCI Fund Management
]
DATAROMA_BASE = "https://www.dataroma.com"

# FX: 1 EUR = X foreign currency, so foreign_in_eur = foreign_price / rate
FX_PAIRS = {"USD": "EURUSD=X", "JPY": "EURJPY=X", "GBP": "EURGBP=X"}

LOOKBACKS = {
    "Today":   dt.timedelta(days=0),
    "1D ago":  dt.timedelta(days=1),
    "1W ago":  dt.timedelta(days=7),
    "1M ago":  dt.timedelta(days=30),
    "6M ago":  dt.timedelta(days=182),
    "1Y ago":  dt.timedelta(days=365),
    "5Y ago":  dt.timedelta(days=int(5 * 365.25)),
}

# ---------------------------------------------------------------------------
# Ticker normalization
# ---------------------------------------------------------------------------

def normalize_ticker(symbol: str, index_name: str) -> str:
    """Convert a Wikipedia-listed symbol into Yahoo's ticker format."""
    s = str(symbol).strip().upper()
    s = s.split()[0]  # strip footnote markers like "AAA[1]"
    s = re.sub(r"\[.*?\]", "", s)
    if index_name == "SP500":
        return s.replace(".", "-")             # BRK.B -> BRK-B
    if index_name == "NIKKEI225":
        return s if s.endswith(".T") else s + ".T"
    if index_name == "FTSE100":
        s = s.rstrip(".")                      # "RR." -> "RR"
        return s if s.endswith(".L") else s + ".L"
    if index_name == "DAX":
        return s if s.endswith(".DE") else s + ".DE"
    if index_name == "CAC40":
        return s if s.endswith(".PA") else s + ".PA"
    if index_name == "BEL20":
        # BEL 20 includes one Amsterdam-listed name (APAM); leave .AS alone too.
        if s.endswith(".BR") or s.endswith(".AS"):
            return s
        return s + ".BR"
    if index_name == "ESTOXX50":
        # Wikipedia's ESTOXX 50 table already lists fully-qualified Yahoo
        # tickers (.DE / .PA / .AS / .MI / .MC / .HE / .IR), spanning multiple
        # Eurozone exchanges. Pass through as-is.
        return s
    return s

# ---------------------------------------------------------------------------
# Wikipedia: index constituents
# ---------------------------------------------------------------------------

_SYMBOL_HINTS = ("symbol", "ticker", "code", "epic")
_NAME_HINTS   = ("company", "security", "name", "issuer")

# BEL 20 Wikipedia formats tickers as "Euronext Brussels:\xa0ABI" — exchange is
# embedded in the cell. Map exchange word → Yahoo suffix.
_BEL20_EXCHANGE_SUFFIX = {"brussels": ".BR", "amsterdam": ".AS"}
_BEL20_TICKER_RE = re.compile(
    r"Euronext\s+(\w+)[\s:\xa0]+([A-Z][A-Z0-9]{0,9})", re.IGNORECASE
)

def _parse_bel20_ticker(cell: str) -> str:
    """Extract 'SYMBOL.BR' or 'SYMBOL.AS' from a BEL 20 Wikipedia ticker cell.

    Returns "" if the cell doesn't match the expected format.
    """
    m = _BEL20_TICKER_RE.search(str(cell))
    if not m:
        return ""
    exchange = m.group(1).lower()
    sym = m.group(2).upper()
    suffix = _BEL20_EXCHANGE_SUFFIX.get(exchange, ".BR")
    return sym + suffix

def _find_constituent_table(tables: list[pd.DataFrame], min_rows: int) -> Optional[tuple[pd.DataFrame, str, str]]:
    """Heuristic: pick the table that has a symbol-like and a name-like column."""
    for tbl in tables:
        cols = [str(c) for c in tbl.columns]
        sym_col = next((c for c in cols if any(h in c.lower() for h in _SYMBOL_HINTS)), None)
        name_col = next((c for c in cols if any(h in c.lower() for h in _NAME_HINTS)), None)
        if sym_col and name_col and len(tbl) >= min_rows:
            return tbl, sym_col, name_col
    return None

def get_index_constituents(index_name: str) -> pd.DataFrame:
    """Returns DataFrame with columns [Symbol, Name, Index]. Symbol is Yahoo-formatted."""
    url = INDEX_WIKI[index_name]
    log.info(f"Fetching {index_name} constituents from Wikipedia")
    # We need a custom UA for Wikipedia too on some networks
    resp = requests.get(url, headers=HTTP_HEADERS, timeout=30)
    resp.raise_for_status()
    tables = pd.read_html(io.StringIO(resp.text))
    expected_min = {"BEL20": 15, "CAC40": 30, "DAX": 30, "FTSE100": 80,
                    "NIKKEI225": 150, "SP500": 400, "ESTOXX50": 45}.get(index_name, 10)
    found = _find_constituent_table(tables, expected_min)
    if not found:
        # Fallback: lower threshold
        found = _find_constituent_table(tables, 10)
    if not found:
        raise RuntimeError(f"Could not locate constituent table for {index_name}")
    tbl, sym_col, name_col = found
    df = tbl[[sym_col, name_col]].copy()
    df.columns = ["Symbol", "Name"]
    df = df.dropna()
    df["Symbol"] = df["Symbol"].astype(str).str.strip()
    df["Name"] = df["Name"].astype(str).str.strip()
    if index_name == "BEL20":
        # Wikipedia formats BEL 20 tickers as "Euronext Brussels:\xa0SYMBOL"
        # or "Euronext Amsterdam:\xa0SYMBOL". Extract symbol + correct Yahoo suffix.
        df["Symbol"] = df["Symbol"].apply(_parse_bel20_ticker)
        df = df[df["Symbol"].astype(bool)]
    df = df[df["Symbol"].str.len().between(1, 12)]
    df["Symbol"] = df["Symbol"].apply(lambda s: normalize_ticker(s, index_name))
    df["Index"] = index_name
    df = df.drop_duplicates(subset=["Symbol"]).reset_index(drop=True)
    log.info(f"  {index_name}: {len(df)} constituents")
    return df

# ---------------------------------------------------------------------------
# Watchlists — Yahoo predefined screener API
# ---------------------------------------------------------------------------
#
# Yahoo's legacy watchlist pages (/u/yahoo-finance/watchlists/) are deprecated
# and don't render content. The same data (where it still exists) is now served
# by the predefined-screener API. Auth requires:
#   1. A session cookie (acquired from any Yahoo Finance page or fc.yahoo.com).
#   2. A "crumb" token from /v1/test/getcrumb — passed back as ?crumb=...
# Without both, /v1/finance/screener/predefined/saved returns total>0 but 0 records.
#
# The endpoint caps responses at 5 records per call regardless of count param,
# so we paginate via &start=N.

_SCREENER_API = "https://query1.finance.yahoo.com/v1/finance/screener/predefined/saved"
_SCREENER_PAGE_STEP = 5    # observed per-call cap
_SCREENER_MAX_RECORDS = 250  # safety stop per screener

def _yahoo_screener_session() -> Optional[tuple[requests.Session, str]]:
    """Acquire a session + crumb for the Yahoo screener API.

    Returns (session, crumb) on success, or None if auth fails (caller should
    skip Yahoo screeners and continue).
    """
    s = requests.Session()
    s.headers.update(HTTP_HEADERS)
    try:
        # fc.yahoo.com seeds the A1/A3 consent cookie used by yfinance too.
        s.get("https://fc.yahoo.com", timeout=10)
        s.get("https://finance.yahoo.com/research-hub/screener/", timeout=20)
        crumb = s.get(
            "https://query1.finance.yahoo.com/v1/test/getcrumb?lang=en-US&region=US",
            timeout=10,
        ).text.strip()
    except Exception as e:
        log.warning(f"Yahoo screener auth failed: {e}")
        return None
    if not crumb or len(crumb) > 64:
        log.warning(f"Yahoo screener auth: unexpected crumb {crumb!r}")
        return None
    return s, crumb

def fetch_yahoo_screener(scrid: str, sess: requests.Session, crumb: str) -> set[str]:
    """Fetch all tickers for a Yahoo predefined screener id (paginated).

    Returns a set of uppercase Yahoo symbols. Failures log a warning and return
    a partial / empty set rather than raising.
    """
    tickers: set[str] = set()
    start = 0
    while start < _SCREENER_MAX_RECORDS:
        try:
            r = sess.get(_SCREENER_API, params={
                "count": 50, "formatted": "true", "scrIds": scrid,
                "sortField": "", "sortType": "", "start": start,
                "useRecordsResponse": "true", "betaFeatureFlag": "true",
                "lang": "en-US", "region": "US", "crumb": crumb,
            }, timeout=20)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            log.warning(f"Yahoo screener '{scrid}' page start={start} failed: {e}")
            break
        err = (data.get("finance") or {}).get("error")
        if err:
            log.warning(f"Yahoo screener '{scrid}' error: {err.get('description', err)}")
            break
        result = (data.get("finance") or {}).get("result") or []
        if not result:
            break
        records = result[0].get("records") or result[0].get("quotes") or []
        if not records:
            break
        for rec in records:
            sym = (rec.get("ticker") or rec.get("symbol") or "").strip().upper()
            if sym and not sym.startswith("^") and not sym.endswith("=F") and not sym.endswith("=X"):
                tickers.add(sym)
        if len(records) < _SCREENER_PAGE_STEP:
            break
        start += _SCREENER_PAGE_STEP
        time.sleep(0.25)
    return tickers

# ---------------------------------------------------------------------------
# Watchlists — Dataroma ("super investor" 13F aggregator)
# ---------------------------------------------------------------------------
#
# Dataroma serves plain HTML; ticker links have the form
#     <a href="/m/stock.php?sym=AAPL">AAPL</a>
# regardless of which list page (portfolio.php, portfolio_b.php, holdings.php).
# A single regex covers all of them.

_DATAROMA_SYM_RE = re.compile(r'/m/stock\.php\?sym=([A-Z][A-Z0-9\.]{0,11})', re.IGNORECASE)

def _extract_dataroma_tickers(html: str) -> set[str]:
    """Parse ticker symbols out of a Dataroma page's HTML.

    Dataroma uses dot-style for SP500 sub-classes (BRK.B, BF.B); Yahoo uses
    hyphen-style (BRK-B, BF-B). We emit BOTH so the membership dict matches
    whichever form the index constituents table contains.
    """
    out: set[str] = set()
    for raw in _DATAROMA_SYM_RE.findall(html):
        sym = raw.upper()
        out.add(sym)
        if "." in sym:
            out.add(sym.replace(".", "-"))
    return out

def fetch_dataroma_tickers(path: str) -> set[str]:
    """Fetch tickers from a Dataroma list page. ``path`` is the URL path or full URL."""
    url = DATAROMA_BASE + path if path.startswith("/") else path
    try:
        r = requests.get(url, headers=HTTP_HEADERS, timeout=20)
        r.raise_for_status()
    except Exception as e:
        log.warning(f"Dataroma fetch '{path}' failed: {e}")
        return set()
    return _extract_dataroma_tickers(r.text)

def fetch_dataroma_activist_aggregate() -> set[str]:
    """Aggregate holdings across the 7 tracked activist hedge fund managers."""
    out: set[str] = set()
    for code in DATAROMA_ACTIVIST_CODES:
        out |= fetch_dataroma_tickers(f"/m/holdings.php?m={code}")
        time.sleep(0.25)
    return out

# ---------------------------------------------------------------------------

def fetch_all_watchlists() -> dict[str, list[str]]:
    """Returns ticker -> [list of watchlist labels it belongs to]."""
    log.info("Fetching watchlists")
    membership: dict[str, list[str]] = {}

    def _record(label: str, tickers: set[str]) -> None:
        log.info(f"  {label}: {len(tickers)} tickers")
        for t in tickers:
            membership.setdefault(t, []).append(label)

    # Yahoo screener: share one auth session across all entries.
    yahoo_entries = [(label, ref) for label, src, ref in WATCHLISTS if src == "yahoo_screener"]
    if yahoo_entries:
        auth = _yahoo_screener_session()
        if auth is None:
            log.warning("Skipping Yahoo screener watchlists (auth unavailable)")
        else:
            sess, crumb = auth
            for label, scrid in yahoo_entries:
                _record(label, fetch_yahoo_screener(scrid, sess, crumb))
                time.sleep(0.5)

    # Dataroma sources (no shared session needed).
    for label, src, ref in WATCHLISTS:
        if src == "dataroma_url":
            _record(label, fetch_dataroma_tickers(ref))
            time.sleep(0.5)
        elif src == "dataroma_activists":
            _record(label, fetch_dataroma_activist_aggregate())
            time.sleep(0.5)

    return membership

# ---------------------------------------------------------------------------
# FX rates
# ---------------------------------------------------------------------------

def get_fx_rates() -> dict[str, float]:
    """Returns {currency_code: units of currency per 1 EUR}. EUR maps to 1.0."""
    rates: dict[str, float] = {"EUR": 1.0}
    log.info("Fetching FX rates (today)")
    for ccy, symbol in FX_PAIRS.items():
        try:
            hist = yf.Ticker(symbol).history(period="5d", auto_adjust=False)
            if hist.empty:
                raise RuntimeError("empty FX history")
            rate = float(hist["Close"].dropna().iloc[-1])
            rates[ccy] = rate
            log.info(f"  EUR/{ccy} = {rate:.4f}")
        except Exception as e:
            log.error(f"  EUR/{ccy} fetch failed: {e}")
            rates[ccy] = float("nan")
    # GBp (UK pence) shares same rate as GBP, but values must be /100 first
    rates["GBp"] = rates.get("GBP", float("nan"))
    return rates

def to_eur(price: float, currency: str, fx: dict[str, float]) -> Optional[float]:
    if price is None or pd.isna(price):
        return None
    if currency == "EUR":
        return float(price)
    if currency == "GBp":
        rate = fx.get("GBP")
        if rate is None or pd.isna(rate):
            return None
        return (float(price) / 100.0) / rate
    rate = fx.get(currency)
    if rate is None or pd.isna(rate):
        return None
    return float(price) / rate

# ---------------------------------------------------------------------------
# Price history & ticker metadata
# ---------------------------------------------------------------------------

def fetch_close_prices(tickers: list[str], start: dt.date, end: dt.date,
                       chunk_size: int = 80) -> pd.DataFrame:
    """Returns a wide DataFrame of close prices: date index, ticker columns."""
    log.info(f"Downloading price history for {len(tickers)} tickers ({start} → {end})")
    all_closes: list[pd.DataFrame] = []
    chunks = [tickers[i:i + chunk_size] for i in range(0, len(tickers), chunk_size)]
    for i, chunk in enumerate(chunks, 1):
        log.info(f"  Price chunk {i}/{len(chunks)} ({len(chunk)} tickers)")
        try:
            df = yf.download(
                chunk, start=start, end=end,
                progress=False, auto_adjust=False, group_by="column", threads=True,
            )
        except Exception as e:
            log.warning(f"    chunk failed: {e}")
            continue
        if df.empty:
            continue
        if isinstance(df.columns, pd.MultiIndex):
            # Columns: (field, ticker). Pull just Close.
            if "Close" in df.columns.get_level_values(0):
                closes = df["Close"]
            else:
                continue
        else:
            # Single ticker case
            if "Close" not in df.columns:
                continue
            closes = df[["Close"]].rename(columns={"Close": chunk[0]})
        all_closes.append(closes)
        time.sleep(1.0)
    if not all_closes:
        return pd.DataFrame()
    combined = pd.concat(all_closes, axis=1)
    # Drop duplicate columns (can happen if a ticker spans chunks somehow)
    combined = combined.loc[:, ~combined.columns.duplicated()]
    return combined

def fetch_ticker_info(ticker: str) -> dict:
    """Returns currency, P/E, name, sector, and business description. Best-effort.

    All fields come from a single ``yf.Ticker(t).info`` call, so adding sector
    and description has zero extra cost over fetching P/E.
    """
    try:
        info = yf.Ticker(ticker).info or {}
    except Exception as e:
        log.debug(f"info('{ticker}') raised: {e}")
        info = {}
    return {
        "currency":    info.get("currency") or "",
        "trailingPE":  info.get("trailingPE"),
        "forwardPE":   info.get("forwardPE"),
        "longName":    info.get("longName") or info.get("shortName") or "",
        "sector":      info.get("sector") or "",
        "description": info.get("longBusinessSummary") or "",
    }

def fetch_all_info(tickers: list[str], delay: float = 0.25) -> dict[str, dict]:
    log.info(f"Fetching .info for {len(tickers)} tickers (delay={delay}s)")
    out: dict[str, dict] = {}
    for i, t in enumerate(tickers, 1):
        if i % 50 == 0 or i == 1:
            log.info(f"  info {i}/{len(tickers)}")
        out[t] = fetch_ticker_info(t)
        time.sleep(delay)
    return out

# ---------------------------------------------------------------------------
# Price lookup
# ---------------------------------------------------------------------------

def price_at_or_before(series: pd.Series, target: dt.date) -> Optional[float]:
    """Last available close on or before target date."""
    if series is None or series.empty:
        return None
    s = series.dropna()
    if s.empty:
        return None
    target_ts = pd.Timestamp(target)
    # series may be tz-aware
    if s.index.tz is not None:
        target_ts = target_ts.tz_localize(s.index.tz)
    valid = s[s.index <= target_ts]
    if valid.empty:
        return None
    return float(valid.iloc[-1])

# ---------------------------------------------------------------------------
# Report assembly
# ---------------------------------------------------------------------------

def aggregate_constituents(per_index: list[pd.DataFrame]) -> pd.DataFrame:
    """Combine per-index constituent DataFrames into one row per Symbol.

    Input rows: ``[Symbol, Name, Index]`` (one row per Symbol per Index).
    Output rows: ``[Symbol, Name, Indexes]`` where Indexes is a comma-separated
    list preserving the order in which indexes were fetched. A stock appearing
    in DAX and EuroStoxx 50 yields one row with ``Indexes = "DAX, ESTOXX50"``.
    """
    if not per_index:
        return pd.DataFrame(columns=["Symbol", "Name", "Indexes"])
    combined = pd.concat(per_index, ignore_index=True)
    if combined.empty:
        return pd.DataFrame(columns=["Symbol", "Name", "Indexes"])

    # Preserve first-seen Name for each Symbol.
    name_map = combined.drop_duplicates(subset=["Symbol"], keep="first").set_index("Symbol")["Name"]
    # Aggregate indexes per Symbol, preserving order, deduping.
    def _join_unique(seq):
        seen = []
        for x in seq:
            if x not in seen:
                seen.append(x)
        return ", ".join(seen)
    indexes_map = combined.groupby("Symbol", sort=False)["Index"].apply(_join_unique)

    out = pd.DataFrame({
        "Symbol":  indexes_map.index,
        "Name":    [name_map.get(s, "") for s in indexes_map.index],
        "Indexes": indexes_map.values,
    })
    return out.reset_index(drop=True)


def build_report(indexes: list[str], info_delay: float) -> tuple[pd.DataFrame, dict]:
    # 1. constituents — one row per Symbol with all index memberships joined
    constituents = aggregate_constituents([get_index_constituents(idx) for idx in indexes])
    log.info(f"Total unique tickers: {len(constituents)}")

    # 2. watchlist memberships
    wl_membership = fetch_all_watchlists()

    # 3. FX
    fx = get_fx_rates()

    # 4. price history (5y + buffer)
    today = dt.date.today()
    start = today - dt.timedelta(days=int(5.2 * 365))
    end   = today + dt.timedelta(days=1)
    closes = fetch_close_prices(list(constituents["Symbol"]), start, end)

    # 5. metadata
    info_map = fetch_all_info(list(constituents["Symbol"]), delay=info_delay)

    # 6. assemble rows
    log.info("Assembling rows")
    rows = []
    for _, c in constituents.iterrows():
        sym = c["Symbol"]
        info = info_map.get(sym, {})
        # Pick the first listed index as the currency fallback (their default
        # currencies disagree only across regions; intersection cases are EUR-EUR).
        first_index = c["Indexes"].split(",")[0].strip()
        ccy = info.get("currency") or INDEX_DEFAULT_CCY.get(first_index, "")
        series = closes[sym] if sym in closes.columns else pd.Series(dtype=float)

        row = {
            "Indexes":  c["Indexes"],
            "Symbol":   sym,
            "Name":     info.get("longName") or c["Name"],
            "Sector":   info.get("sector") or "",
            "Currency": ccy,
        }
        for label, delta in LOOKBACKS.items():
            target = today - delta
            native = price_at_or_before(series, target)
            row[f"{label} (EUR)"] = to_eur(native, ccy, fx)

        row["P/E (TTM)"]   = info.get("trailingPE")
        row["Forward P/E"] = info.get("forwardPE")

        # Watchlist membership: try sym, sym without suffix, raw US ticker
        sym_root = re.sub(r"\.[A-Z]{1,3}$", "", sym)
        labels: list[str] = []
        for key in {sym, sym_root}:
            labels.extend(wl_membership.get(key, []))
        seen = set()
        labels_dedup = [x for x in labels if not (x in seen or seen.add(x))]
        row["Watchlists"]  = ", ".join(labels_dedup)
        row["Description"] = info.get("description") or ""

        rows.append(row)

    df = pd.DataFrame(rows)
    return df, fx

# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

def write_excel(df: pd.DataFrame, fx: dict, output_path: Path) -> None:
    log.info(f"Writing {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All", index=False)
        meta = pd.DataFrame({
            "Field": [
                "Run timestamp (UTC)", "Total tickers",
                "EUR/USD (1 EUR = X USD)", "EUR/JPY (1 EUR = X JPY)", "EUR/GBP (1 EUR = X GBP)",
                "Note",
            ],
            "Value": [
                dt.datetime.now(dt.UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
                len(df),
                fx.get("USD"), fx.get("JPY"), fx.get("GBP"),
                "Historical prices converted using TODAY's FX rate.",
            ],
        })
        meta.to_excel(writer, sheet_name="Metadata", index=False)

        ws = writer.sheets["All"]
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # Auto-size columns
        for col_idx, col in enumerate(df.columns, 1):
            sample = df[col].head(200)
            max_len = max([len(str(col))] + [len(str(s)) for s in sample])
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 40)

        # Number formats
        numeric_cols = [c for c in df.columns if "(EUR)" in c or c in ("P/E (TTM)", "Forward P/E")]
        for col_name in numeric_cols:
            col_idx = df.columns.get_loc(col_name) + 1
            for row_idx in range(2, len(df) + 2):
                ws.cell(row=row_idx, column=col_idx).number_format = "#,##0.00"

# ---------------------------------------------------------------------------
# Persistent workbook — Main + Market sheets
# ---------------------------------------------------------------------------
#
# The "two jobs" architecture stores everything in a single .xlsx (later .xlsm
# with xlwings buttons). Two sheets:
#
#   Main   — user-maintained portfolio plus job controls, status, metadata.
#            Survives every job run; rebuild_inventory and get_quotes never
#            modify rows the user owns (portfolio area), only the named cells.
#   Market — full inventory of stocks across all tracked indexes. Job 1
#            (rebuild_inventory) overwrites this in full; Job 2 (get_quotes)
#            updates the quote/PE/timestamp columns by Symbol key.

DEFAULT_WORKBOOK_PATH = Path("stocks.xlsx")

# Market sheet column order. Fixed by index — every part of the code that
# reads/writes Market locates columns by name via this list.
MARKET_COLUMNS = [
    "Symbol", "Name", "Owned?", "Indexes", "Sector", "Watchlists", "Currency",
    "Today (EUR)", "1D ago (EUR)", "1W ago (EUR)", "1M ago (EUR)",
    "6M ago (EUR)", "1Y ago (EUR)", "5Y ago (EUR)",
    "P/E (TTM)", "Forward P/E",
    "Description", "Last update (UTC)", "Last error",
]

# Main sheet — named cells. Code reads/writes these by name; layout can
# move within the sheet as long as the dict tracks it.
MAIN_CELLS: dict[str, str] = {
    "TestMode":       "B21",  # TRUE / FALSE
    "Status":         "B22",  # live progress text written during job runs
    "LastRebuildAt":  "B25",
    "LastQuotesAt":   "B26",
    "MarketRowCount": "B27",
    "EurUsd":         "B28",
    "EurJpy":         "B29",
    "EurGbp":         "B30",
}

# Portfolio area on Main: header row + N rows for user entries.
PORTFOLIO_HEADER_ROW = 4
PORTFOLIO_FIRST_ROW = 5
PORTFOLIO_LAST_ROW = 100  # user can have up to 96 manual entries


def _bold(cell, size: int = 11) -> None:
    cell.font = Font(bold=True, size=size)


def _layout_main_sheet(ws: Worksheet) -> None:
    """Write the static layout of the Main sheet — title, portfolio header,
    section labels, metadata field labels. Does NOT touch user data cells."""
    ws["A1"] = "Stock Picker"
    _bold(ws["A1"], size=16)

    ws["A3"] = "Portfolio (manual — list every Symbol you own)"
    _bold(ws["A3"], size=12)

    ws.cell(row=PORTFOLIO_HEADER_ROW, column=1, value="Symbol")
    ws.cell(row=PORTFOLIO_HEADER_ROW, column=2, value="Quantity")
    ws.cell(row=PORTFOLIO_HEADER_ROW, column=3, value="Notes")
    for col in (1, 2, 3):
        _bold(ws.cell(row=PORTFOLIO_HEADER_ROW, column=col))

    ws["A19"] = "Jobs"
    _bold(ws["A19"], size=12)
    ws["A20"] = "(buttons wired via xlwings — see README)"
    ws["A20"].font = Font(italic=True, color="666666")

    ws["A21"] = "Test mode (only refresh BEL20 + 1 quote):"
    ws["A22"] = "Status:"

    ws["A24"] = "Metadata"
    _bold(ws["A24"], size=12)

    ws["A25"] = "Last rebuild_inventory:"
    ws["A26"] = "Last get_quotes:"
    ws["A27"] = "Total rows in Market:"
    ws["A28"] = "EUR/USD (1 EUR = X USD):"
    ws["A29"] = "EUR/JPY (1 EUR = X JPY):"
    ws["A30"] = "EUR/GBP (1 EUR = X GBP):"

    # Column widths
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 40


def _layout_market_sheet(ws: Worksheet) -> None:
    """Write Market sheet headers, freeze pane, column widths. Does not touch
    data rows."""
    for col_idx, name in enumerate(MARKET_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        _bold(cell)
        cell.fill = PatternFill("solid", fgColor="F2F2F2")
    ws.freeze_panes = "A2"

    # Per-column widths tuned for readability. Description is wide; quote
    # columns stay narrow.
    widths = {
        "Symbol": 10, "Name": 28, "Owned?": 8, "Indexes": 18, "Sector": 22,
        "Watchlists": 38, "Currency": 10,
        "Today (EUR)": 12, "1D ago (EUR)": 12, "1W ago (EUR)": 12, "1M ago (EUR)": 12,
        "6M ago (EUR)": 12, "1Y ago (EUR)": 12, "5Y ago (EUR)": 12,
        "P/E (TTM)": 10, "Forward P/E": 10,
        "Description": 60, "Last update (UTC)": 20, "Last error": 30,
    }
    for col_idx, name in enumerate(MARKET_COLUMNS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = widths.get(name, 12)


def init_workbook(path: Path) -> Path:
    """Create or refresh the persistent workbook layout.

    Idempotent: if ``path`` exists, this function preserves all user data and
    only re-asserts headers, section labels, and column widths. It will never
    blow away the portfolio area or the Market data rows.
    """
    path = Path(path)
    if path.exists():
        log.info(f"Refreshing workbook layout in {path} (data preserved)")
        wb = load_workbook(path)
        main_ws = wb["Main"] if "Main" in wb.sheetnames else wb.create_sheet("Main", 0)
        market_ws = wb["Market"] if "Market" in wb.sheetnames else wb.create_sheet("Market", 1)
    else:
        log.info(f"Creating workbook {path}")
        wb = Workbook()
        # Default sheet "Sheet" → rename to Main; add Market.
        main_ws = wb.active
        main_ws.title = "Main"
        market_ws = wb.create_sheet("Market")

    _layout_main_sheet(main_ws)
    _layout_market_sheet(market_ws)

    # Seed Test mode = FALSE if blank (don't overwrite user choice).
    if main_ws[MAIN_CELLS["TestMode"]].value is None:
        main_ws[MAIN_CELLS["TestMode"]] = "FALSE"
    if main_ws[MAIN_CELLS["Status"]].value is None:
        main_ws[MAIN_CELLS["Status"]] = "(idle — never run)"

    wb.save(path)
    log.info(f"Workbook ready: {path.resolve()}")
    return path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _cmd_legacy_run(args) -> int:
    """Legacy single-shot mode: builds a fresh stocks_report_YYYY-MM-DD.xlsx."""
    if args.indexes.upper() == "ALL":
        indexes = list(INDEX_WIKI.keys())
    else:
        indexes = [s.strip().upper() for s in args.indexes.split(",")]
        unknown = [i for i in indexes if i not in INDEX_WIKI]
        if unknown:
            log.error(f"Unknown index(es): {unknown}")
            log.error(f"Valid: {list(INDEX_WIKI.keys())}")
            return 2

    output = Path(args.output) if args.output else Path(f"stocks_report_{dt.date.today()}.xlsx")

    t0 = time.time()
    df, fx = build_report(indexes, info_delay=args.info_delay)
    write_excel(df, fx, output)
    log.info(f"Done in {time.time() - t0:.0f}s. Rows: {len(df)}. Output: {output}")
    return 0


def _cmd_init_workbook(args) -> int:
    """Create or refresh stocks.xlsx with Main + Market sheet templates."""
    init_workbook(Path(args.workbook))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    sub = parser.add_subparsers(dest="cmd")

    # init-workbook — create or refresh the persistent workbook
    p_init = sub.add_parser(
        "init-workbook",
        help="Create or refresh stocks.xlsx with Main + Market sheet templates.",
    )
    p_init.add_argument(
        "--workbook", default=str(DEFAULT_WORKBOOK_PATH),
        help=f"Workbook path (default: {DEFAULT_WORKBOOK_PATH})",
    )

    # run — legacy single-shot all-in-one report (default when no subcommand)
    p_run = sub.add_parser(
        "run",
        help="Legacy single-shot: write a fresh dated xlsx with everything.",
    )
    p_run.add_argument(
        "--indexes", default="ALL",
        help="Comma-separated: SP500,NIKKEI225,FTSE100,DAX,CAC40,BEL20,ESTOXX50 (default: ALL)",
    )
    p_run.add_argument(
        "--output", default=None,
        help="Output xlsx path (default: stocks_report_YYYY-MM-DD.xlsx in current dir)",
    )
    p_run.add_argument(
        "--info-delay", type=float, default=0.25,
        help="Seconds between yfinance .info calls (default: 0.25)",
    )

    # Legacy compat: when called as `stocks_report.py --indexes BEL20` with
    # no subcommand, fall back to the run subparser. Detect by looking for
    # a known top-level flag.
    raw_args = sys.argv[1:]
    if raw_args and raw_args[0] not in {"init-workbook", "run", "-h", "--help"}:
        raw_args = ["run", *raw_args]
    args = parser.parse_args(raw_args)

    if args.cmd is None or args.cmd == "run":
        # Provide defaults for legacy invocation (`stocks_report.py` with no args).
        for attr, default in (("indexes", "ALL"), ("output", None), ("info_delay", 0.25)):
            if not hasattr(args, attr):
                setattr(args, attr, default)
        return _cmd_legacy_run(args)
    if args.cmd == "init-workbook":
        return _cmd_init_workbook(args)

    parser.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
