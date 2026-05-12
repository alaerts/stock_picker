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
from typing import Callable, Optional

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

_NIKKEI_TICKER_RE = re.compile(r"\(\s*TYO\s*:\s*([0-9A-Z]{4,5})\s*\)")


_NIKKEI225_MIN_COUNT = 150  # safety floor; real value is ~225


def _parse_nikkei225_components(html: str, min_count: int = _NIKKEI225_MIN_COUNT) -> pd.DataFrame:
    """Extract (Symbol, Name, Index) rows from the Components section of the
    English Wikipedia Nikkei 225 page.

    Wikipedia stopped publishing a constituent table for Nikkei 225 sometime
    in 2025. The 225 stocks now live in bullet lists under sector
    subheadings, each entry formatted "Company Name (TYO: 9202)" with
    occasional trailing parenthetical annotations like
    "(Holding company for X)". Tickers can be 4 digits or 4 chars including
    letters (e.g. "543A").

    Raises RuntimeError if the result is implausibly small — protects
    against a future Wikipedia restructure silently degrading the workbook
    instead of surfacing as a visible error.
    """
    start = html.find('id="Components"')
    end = html.find('id="Statistics"')
    if start < 0 or end < 0:
        raise RuntimeError("Could not locate Components section on Nikkei 225 Wikipedia page")
    body = html[start:end]
    rows: list[tuple[str, str]] = []
    for raw in re.findall(r"<li[^>]*>(.+?)</li>", body, re.DOTALL):
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", raw)).strip()
        m = _NIKKEI_TICKER_RE.search(text)
        if not m:
            continue
        ticker = m.group(1)
        # Name is the prefix before "(TYO:" — strip stray trailing punctuation
        name = text[: m.start()].rstrip().rstrip("(").strip()
        if name and ticker:
            rows.append((ticker, name))
    df = pd.DataFrame(rows, columns=["Symbol", "Name"])
    df["Symbol"] = df["Symbol"].apply(lambda s: normalize_ticker(s, "NIKKEI225"))
    df["Index"] = "NIKKEI225"
    df = df.drop_duplicates(subset=["Symbol"]).reset_index(drop=True)
    if len(df) < min_count:
        raise RuntimeError(
            f"Nikkei 225 parser yielded only {len(df)} constituents "
            f"(expected ~225, floor {min_count}). Wikipedia layout likely changed."
        )
    return df


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

    if index_name == "NIKKEI225":
        # Wikipedia restructured the page; no constituent table any more.
        df = _parse_nikkei225_components(resp.text)
        log.info(f"  {index_name}: {len(df)} constituents")
        return df

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

def fetch_all_info(tickers: list[str], delay: float = 0.25,
                   progress: Optional[Callable[[int, int, str], None]] = None) -> dict[str, dict]:
    """Fetch yfinance .info for each ticker. Optional ``progress`` callback is
    invoked as ``progress(done, total, current_symbol)`` before each fetch — used
    by rebuild_inventory to push live ticker count into the workbook's Status cell.
    """
    log.info(f"Fetching .info for {len(tickers)} tickers (delay={delay}s)")
    out: dict[str, dict] = {}
    for i, t in enumerate(tickers, 1):
        if i % 50 == 0 or i == 1:
            log.info(f"  info {i}/{len(tickers)}")
        if progress is not None:
            try:
                progress(i, len(tickers), t)
            except Exception as e:
                log.debug(f"progress callback failed: {e}")
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


def _resolve_workbook(arg_path: str) -> Path:
    """Auto-upgrade: if the user kept the default 'stocks.xlsx' but
    'stocks.xlsm' exists (i.e. setup-buttons has been run), use the xlsm.
    Saves the user from having to pass --workbook stocks.xlsm everywhere.
    """
    p = Path(arg_path)
    if p == Path("stocks.xlsx") and Path("stocks.xlsm").exists():
        return Path("stocks.xlsm")
    return p

# Market sheet column order. Fixed by index — every part of the code that
# reads/writes Market locates columns by name via this list.
MARKET_COLUMNS = [
    "Symbol", "Name", "Owned?", "Indexes", "Sector", "Watchlists", "Currency",
    "Today (EUR)", "1D ago (EUR)", "1W ago (EUR)", "1M ago (EUR)",
    "6M ago (EUR)", "1Y ago (EUR)", "5Y ago (EUR)",
    "P/E (TTM)", "Forward P/E",
    "Description", "Last update (UTC)", "Last error",
]

# Main sheet — named cells. All controls/metadata live ABOVE the portfolio
# area so the portfolio can grow freely without colliding with them.
MAIN_CELLS: dict[str, str] = {
    "TestMode":       "B5",   # TRUE / FALSE
    "Status":         "B6",   # live progress text written during job runs
    "LastRebuildAt":  "B9",
    "LastQuotesAt":   "B10",
    "MarketRowCount": "B11",
    "EurUsd":         "B12",
    "EurJpy":         "B13",
    "EurGbp":         "B14",
}

# Portfolio area on Main: header row + many rows for user entries.
PORTFOLIO_HEADER_ROW = 17
PORTFOLIO_FIRST_ROW = 18
PORTFOLIO_LAST_ROW = 999  # plenty of room; never collides with controls above


def _bold(cell, size: int = 11) -> None:
    cell.font = Font(bold=True, size=size)


def _layout_main_sheet(ws: Worksheet, *, overwrite: bool = True) -> None:
    """Write the static layout of the Main sheet — title, controls + metadata
    at the top, then portfolio header. Does NOT touch user data cells.

    With overwrite=False, only fills cells that are currently blank — used on
    re-runs of init-workbook so user cosmetic edits / renames survive.

    Layout:
      Row 1   : title
      Row 3   : "Jobs" section
      Row 4   : (xlwings button hint)
      Row 5   : Test mode label + checkbox cell (B5)
      Row 6   : Status label + live progress cell (B6)
      Row 8   : "Metadata" section
      Row 9-14: timestamps + FX rates
      Row 16  : "Portfolio" section
      Row 17  : Portfolio header (Symbol / Notes)
      Row 18+ : user portfolio entries
    """
    def _set(addr: str, value, *, bold: bool = False, size: int = 11,
             italic: bool = False, color: Optional[str] = None) -> None:
        cell = ws[addr]
        if overwrite or cell.value is None:
            cell.value = value
        # Style only when the value matches what we just wrote — leaves
        # user-edited labels with their own styling.
        if cell.value == value:
            font_kwargs = {"size": size}
            if bold: font_kwargs["bold"] = True
            if italic: font_kwargs["italic"] = True
            if color: font_kwargs["color"] = color
            cell.font = Font(**font_kwargs)

    _set("A1", "Stock Picker", bold=True, size=16)

    _set("A3", "Jobs", bold=True, size=12)
    _set("A4", "(buttons wired via xlwings — see README)", italic=True, color="666666")
    _set("A5", "Test mode (only refresh BEL20 + 1 quote):")
    _set("A6", "Status: (live progress while a job runs; last-run summary when idle)")

    _set("A8", "Metadata", bold=True, size=12)
    _set("A9",  "Last rebuild_inventory:")
    _set("A10", "Last get_quotes:")
    _set("A11", "Total rows in Market:")
    _set("A12", "EUR/USD (1 EUR = X USD):")
    _set("A13", "EUR/JPY (1 EUR = X JPY):")
    _set("A14", "EUR/GBP (1 EUR = X GBP):")

    _set("A16", "Portfolio (manual — list every Symbol you own)", bold=True, size=12)

    _set(ws.cell(row=PORTFOLIO_HEADER_ROW, column=1).coordinate, "Symbol", bold=True)
    _set(ws.cell(row=PORTFOLIO_HEADER_ROW, column=2).coordinate, "Notes", bold=True)

    # Column widths: only reset on fresh creation (overwrite=True) to respect
    # any custom widths the user has set.
    if overwrite:
        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 60


def _layout_market_sheet(ws: Worksheet, *, overwrite: bool = True) -> None:
    """Write Market sheet headers, freeze pane, column widths. Does not touch
    data rows.

    With overwrite=False (re-run path), header text is always re-asserted
    (the script's column-name lookups depend on it) but bold + fill styling
    and column widths are left alone so user cosmetic choices survive.
    """
    for col_idx, name in enumerate(MARKET_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx)
        if overwrite or cell.value is None or cell.value != name:
            cell.value = name
        if overwrite:
            _bold(cell)
            cell.fill = PatternFill("solid", fgColor="F2F2F2")
    ws.freeze_panes = "A2"

    if overwrite:
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


def _market_col(name: str) -> int:
    """1-based column index in Market for a header name. Raises if unknown."""
    return MARKET_COLUMNS.index(name) + 1


_TEST_MODE_TRUTHY = {"TRUE", "T", "YES", "Y", "1"}


def read_test_mode(main_ws: Worksheet) -> bool:
    """Interpret the Main!TestMode cell as a boolean.

    Tolerates the cell being TRUE/FALSE strings (default seed), a literal
    Python bool (some xlwings paths return that), 1/0, "Yes"/"No", or blank
    (treated as False).
    """
    raw = main_ws[MAIN_CELLS["TestMode"]].value
    if raw is None:
        return False
    if isinstance(raw, bool):
        return raw
    return str(raw).strip().upper() in _TEST_MODE_TRUTHY


def read_portfolio_symbols(main_ws: Worksheet) -> set[str]:
    """Read the Main sheet's manual portfolio area, returning normalized symbols."""
    out: set[str] = set()
    for r in range(PORTFOLIO_FIRST_ROW, PORTFOLIO_LAST_ROW + 1):
        v = main_ws.cell(row=r, column=1).value
        if v is None:
            continue
        sym = str(v).strip().upper()
        if sym:
            out.add(sym)
    return out


def _owned_for(symbol: str, portfolio: set[str]) -> str:
    """Return "Yes" if the Market symbol matches a portfolio entry, else "No".

    Tolerates suffix mismatch: portfolio entry "KBC" matches Market "KBC.BR".
    """
    if symbol in portfolio:
        return "Yes"
    root = re.sub(r"\.[A-Z]{1,3}$", "", symbol)
    if root in portfolio:
        return "Yes"
    return "No"


def _load_xl(path: Path):
    """load_workbook with keep_vba=True for .xlsm (preserves user-imported VBA)."""
    if path.suffix.lower() == ".xlsm":
        return load_workbook(path, keep_vba=True)
    return load_workbook(path)


def init_workbook(path: Path) -> Path:
    """Create or refresh the persistent workbook layout.

    Idempotent: if ``path`` exists, this function preserves all user data and
    only re-asserts headers, section labels, and column widths. It will never
    blow away the portfolio area, the Market data rows, or any VBA the user
    has imported into the workbook.
    """
    path = Path(path)
    if path.exists():
        log.info(f"Refreshing workbook layout in {path} (data + cosmetics preserved)")
        wb = _load_xl(path)
        main_ws = wb["Main"] if "Main" in wb.sheetnames else wb.create_sheet("Main", 0)
        market_ws = wb["Market"] if "Market" in wb.sheetnames else wb.create_sheet("Market", 1)
        overwrite = False
    else:
        log.info(f"Creating workbook {path}")
        wb = Workbook()
        # Default sheet "Sheet" → rename to Main; add Market.
        main_ws = wb.active
        main_ws.title = "Main"
        market_ws = wb.create_sheet("Market")
        overwrite = True

    _layout_main_sheet(main_ws, overwrite=overwrite)
    _layout_market_sheet(market_ws, overwrite=overwrite)

    # Seed Test mode = FALSE if blank (don't overwrite user choice). The
    # cell is replaced by a real Excel checkbox in setup-buttons; until then
    # the user can edit this text cell directly to flip the mode.
    if main_ws[MAIN_CELLS["TestMode"]].value is None:
        main_ws[MAIN_CELLS["TestMode"]] = "FALSE"
    if main_ws[MAIN_CELLS["Status"]].value is None:
        main_ws[MAIN_CELLS["Status"]] = "Idle. Click a button to refresh."

    wb.save(path)
    log.info(f"Workbook ready: {path.resolve()}")
    return path


# ---------------------------------------------------------------------------
# Job 1: rebuild_inventory — refresh Market structural data
# ---------------------------------------------------------------------------
#
# Fetches the universe of stocks (constituents + watchlist memberships +
# company info) and overwrites the Market sheet in place. Quote columns are
# left blank for get_quotes to fill.

def rebuild_inventory(
    workbook_path: Path,
    indexes: Optional[list[str]] = None,
    info_delay: float = 0.25,
    status: Optional[Callable[[str], None]] = None,
    test_mode: Optional[bool] = None,
) -> int:
    """Rebuild the Market sheet from scratch. Preserves Main (portfolio + cells)
    other than the metadata cells this job owns.

    ``indexes``: subset of INDEX_WIKI keys; defaults to all.
    ``test_mode``: True / False / None. When None, reads Main!TestMode cell.
                   When True, restricts to BEL20 regardless of ``indexes``.
    ``status``: optional one-arg callback for live progress text (Phase 8
                wires this to Main!Status via xlwings; CLI mode passes None
                and progress goes to the log instead).
    """
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        log.info(f"Workbook {workbook_path} does not exist — running init-workbook first")
        init_workbook(workbook_path)

    def _say(msg: str) -> None:
        log.info(msg)
        if status is not None:
            try:
                status(msg)
            except Exception as e:
                log.debug(f"status callback raised: {e}")

    t0 = time.time()
    _say("Reading portfolio from Main")
    wb = _load_xl(workbook_path)
    if test_mode is None:
        test_mode = read_test_mode(wb["Main"])
    if test_mode:
        indexes = ["BEL20"]
        _say("Test mode ON — BEL20 only")
    elif indexes is None:
        indexes = list(INDEX_WIKI.keys())
    _say(f"rebuild_inventory: starting ({', '.join(indexes)})")

    portfolio = read_portfolio_symbols(wb["Main"])
    _say(f"  Portfolio entries: {len(portfolio)}")

    _say("Fetching index constituents")
    constituents = aggregate_constituents([get_index_constituents(idx) for idx in indexes])
    _say(f"  Total unique tickers: {len(constituents)}")

    _say("Fetching watchlists")
    wl_membership = fetch_all_watchlists()

    _say(f"Fetching .info for {len(constituents)} tickers — this is the slow part")

    def info_progress(done: int, total: int, sym: str) -> None:
        # Throttle status updates: report every 25 tickers + at boundaries.
        if status is not None and (done == 1 or done == total or done % 25 == 0):
            status(f"info {done}/{total} — {sym}")

    info_map = fetch_all_info(list(constituents["Symbol"]), delay=info_delay,
                              progress=info_progress)

    _say("Writing Market sheet")
    market_ws = wb["Market"]
    # Clear existing data rows (preserve row 1 headers).
    if market_ws.max_row > 1:
        market_ws.delete_rows(2, market_ws.max_row - 1)

    sym_col = _market_col("Symbol")
    cols = {name: i + 1 for i, name in enumerate(MARKET_COLUMNS)}

    for r_offset, (_, c) in enumerate(constituents.iterrows(), 0):
        row = 2 + r_offset
        sym = c["Symbol"]
        info = info_map.get(sym, {})

        # Watchlist membership: try sym + suffix-stripped root
        sym_root = re.sub(r"\.[A-Z]{1,3}$", "", sym)
        labels: list[str] = []
        for key in {sym, sym_root}:
            labels.extend(wl_membership.get(key, []))
        seen = set()
        labels_dedup = [x for x in labels if not (x in seen or seen.add(x))]

        first_index = c["Indexes"].split(",")[0].strip()
        ccy = info.get("currency") or INDEX_DEFAULT_CCY.get(first_index, "")

        market_ws.cell(row=row, column=cols["Symbol"],      value=sym)
        market_ws.cell(row=row, column=cols["Name"],        value=info.get("longName") or c["Name"])
        market_ws.cell(row=row, column=cols["Owned?"],      value=_owned_for(sym, portfolio))
        market_ws.cell(row=row, column=cols["Indexes"],     value=c["Indexes"])
        market_ws.cell(row=row, column=cols["Sector"],      value=info.get("sector") or "")
        market_ws.cell(row=row, column=cols["Watchlists"],  value=", ".join(labels_dedup))
        market_ws.cell(row=row, column=cols["Currency"],    value=ccy)
        market_ws.cell(row=row, column=cols["Description"], value=info.get("description") or "")
        # Quote columns and Last update/error are left blank for get_quotes.

    # Update Main metadata cells
    main_ws = wb["Main"]
    now_iso = dt.datetime.now(dt.UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    main_ws[MAIN_CELLS["LastRebuildAt"]] = now_iso
    main_ws[MAIN_CELLS["MarketRowCount"]] = len(constituents)

    wb.save(workbook_path)
    duration = time.time() - t0
    _say(f"rebuild_inventory: done in {duration:.0f}s. {len(constituents)} rows.")
    return 0


# ---------------------------------------------------------------------------
# xlwings entry points — called from VBA buttons in stocks.xlsm
# ---------------------------------------------------------------------------
#
# These functions:
#   * Use xlwings to read state from / write cells to the OPEN workbook
#     (openpyxl can't write to a file Excel has open).
#   * Stream "Status:" cell updates so the user sees live progress.
#   * Share the network fetchers (get_index_constituents, fetch_all_watchlists,
#     fetch_all_info) with the CLI rebuild_inventory / get_quotes paths.
#
# Wiring: VBA module stocks_picker.bas (see vba/) calls
#   RunPython "import stocks_report; stocks_report.button_rebuild_inventory()"
# from each macro. The user imports stocks_picker.bas once via Alt+F11, then
# assigns macros to button shapes via right-click → Assign Macro.

def _xw_status(main_sheet) -> Callable[[str], None]:
    """Status writer that updates Main!Status live. Truncated to fit one cell."""
    addr = MAIN_CELLS["Status"]
    def _set(msg: str) -> None:
        main_sheet.range(addr).value = msg[:200]
    return _set


def _xw_read_portfolio(main_sheet) -> set[str]:
    out: set[str] = set()
    for r in range(PORTFOLIO_FIRST_ROW, PORTFOLIO_LAST_ROW + 1):
        v = main_sheet.range((r, 1)).value
        if v is None:
            continue
        sym = str(v).strip().upper()
        if sym:
            out.add(sym)
    return out


def _xw_read_test_mode(main_sheet) -> bool:
    raw = main_sheet.range(MAIN_CELLS["TestMode"]).value
    if raw is None:
        return False
    if isinstance(raw, bool):
        return raw
    return str(raw).strip().upper() in _TEST_MODE_TRUTHY


def button_rebuild_inventory() -> None:
    """Entry point for the Excel "Rebuild Inventory" button."""
    import xlwings as xw
    wb = xw.Book.caller()
    main = wb.sheets["Main"]
    market = wb.sheets["Market"]
    status = _xw_status(main)

    try:
        test_mode = _xw_read_test_mode(main)
        indexes = ["BEL20"] if test_mode else list(INDEX_WIKI.keys())
        status(f"rebuild: starting ({'TEST: BEL20' if test_mode else 'full'})")

        portfolio = _xw_read_portfolio(main)
        status(f"Portfolio entries: {len(portfolio)}")

        status("Fetching index constituents…")
        constituents = aggregate_constituents([get_index_constituents(idx) for idx in indexes])
        status(f"Total unique tickers: {len(constituents)}")

        status("Fetching watchlists…")
        wl_membership = fetch_all_watchlists()

        status(f"Fetching .info for {len(constituents)} tickers (slow step)")

        def info_progress(done: int, total: int, sym: str) -> None:
            if done == 1 or done == total or done % 10 == 0:
                status(f"info {done}/{total} — {sym}")

        info_map = fetch_all_info(list(constituents["Symbol"]),
                                  delay=0.25, progress=info_progress)

        status("Building rows…")
        rows = []
        for _, c in constituents.iterrows():
            sym = c["Symbol"]
            info = info_map.get(sym, {})
            sym_root = re.sub(r"\.[A-Z]{1,3}$", "", sym)
            labels: list[str] = []
            for key in {sym, sym_root}:
                labels.extend(wl_membership.get(key, []))
            seen = set()
            labels_dedup = [x for x in labels if not (x in seen or seen.add(x))]
            first_index = c["Indexes"].split(",")[0].strip()
            ccy = info.get("currency") or INDEX_DEFAULT_CCY.get(first_index, "")
            row = [None] * len(MARKET_COLUMNS)
            row[MARKET_COLUMNS.index("Symbol")]      = sym
            row[MARKET_COLUMNS.index("Name")]        = info.get("longName") or c["Name"]
            row[MARKET_COLUMNS.index("Owned?")]      = _owned_for(sym, portfolio)
            row[MARKET_COLUMNS.index("Indexes")]     = c["Indexes"]
            row[MARKET_COLUMNS.index("Sector")]      = info.get("sector") or ""
            row[MARKET_COLUMNS.index("Watchlists")]  = ", ".join(labels_dedup)
            row[MARKET_COLUMNS.index("Currency")]    = ccy
            row[MARKET_COLUMNS.index("Description")] = info.get("description") or ""
            rows.append(row)

        status(f"Writing {len(rows)} rows to Market…")
        # Clear existing data rows
        last_row = market.used_range.last_cell.row
        if last_row >= 2:
            market.range(f"A2:{get_column_letter(len(MARKET_COLUMNS))}{last_row}").clear_contents()
        # Bulk write — single COM call
        if rows:
            market.range((2, 1)).value = rows

        now_iso = dt.datetime.now(dt.UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
        main.range(MAIN_CELLS["LastRebuildAt"]).value = now_iso
        main.range(MAIN_CELLS["MarketRowCount"]).value = len(rows)
        status(f"rebuild: done. {len(rows)} rows at {now_iso}.")
    except Exception as e:
        status(f"ERROR: {type(e).__name__}: {e}"[:200])
        raise


def button_get_quotes() -> None:
    """Entry point for the Excel "Get Quotes" button."""
    import xlwings as xw
    wb = xw.Book.caller()
    main = wb.sheets["Main"]
    market = wb.sheets["Market"]
    status = _xw_status(main)

    try:
        test_mode = _xw_read_test_mode(main)
        status(f"quotes: starting (test_mode={test_mode})")

        # Read current Market: pull Symbol + Currency + Indexes into a list
        last_row = market.used_range.last_cell.row
        if last_row < 2:
            status("Market is empty — run Rebuild Inventory first")
            return

        sym_col = MARKET_COLUMNS.index("Symbol") + 1
        ccy_col = MARKET_COLUMNS.index("Currency") + 1
        idx_col = MARKET_COLUMNS.index("Indexes") + 1

        market_rows: list[tuple[int, str, str, str]] = []
        # Bulk-read all three columns in one shot
        syms = market.range((2, sym_col), (last_row, sym_col)).value
        ccys = market.range((2, ccy_col), (last_row, ccy_col)).value
        idxs = market.range((2, idx_col), (last_row, idx_col)).value
        # When the range is a single cell xlwings returns a scalar, not a list
        if not isinstance(syms, list): syms = [syms]
        if not isinstance(ccys, list): ccys = [ccys]
        if not isinstance(idxs, list): idxs = [idxs]
        for offset, (s, c, i) in enumerate(zip(syms, ccys, idxs)):
            if not s:
                continue
            market_rows.append((2 + offset, str(s).strip(), str(c or ""), str(i or "")))

        portfolio = _xw_read_portfolio(main)
        if test_mode:
            picked = _pick_test_target(market_rows, portfolio)
            targets = [picked] if picked else []
            if not targets:
                status("Test mode: could not pick a target")
                return
            status(f"Test mode: refreshing {targets[0][1]} only")
        else:
            targets = market_rows

        status("Fetching FX rates")
        fx = get_fx_rates()

        status(f"Downloading price history for {len(targets)} tickers")
        today = dt.date.today()
        start = today - dt.timedelta(days=int(5.2 * 365))
        end   = today + dt.timedelta(days=1)
        closes = fetch_close_prices([t[1] for t in targets], start, end)

        status(f"Fetching .info for {len(targets)} tickers (P/E)")

        def info_progress(done: int, total: int, sym: str) -> None:
            if done == 1 or done == total or done % 10 == 0:
                status(f"quotes {done}/{total} — {sym}")

        info_map = fetch_all_info([t[1] for t in targets], delay=0.25, progress=info_progress)

        status(f"Writing quote columns for {len(targets)} rows…")
        now_iso = dt.datetime.now(dt.UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
        cols = {n: MARKET_COLUMNS.index(n) + 1 for n in MARKET_COLUMNS}

        ok = 0
        fail = 0
        for (row_idx, sym, ccy, _idx) in targets:
            try:
                series = closes[sym] if sym in closes.columns else pd.Series(dtype=float)
                for label, delta in LOOKBACKS.items():
                    target_d = today - delta
                    native = price_at_or_before(series, target_d)
                    market.range((row_idx, cols[f"{label} (EUR)"])).value = to_eur(native, ccy, fx)
                info = info_map.get(sym, {})
                market.range((row_idx, cols["P/E (TTM)"])).value   = info.get("trailingPE")
                market.range((row_idx, cols["Forward P/E"])).value = info.get("forwardPE")
                market.range((row_idx, cols["Last update (UTC)"])).value = now_iso
                market.range((row_idx, cols["Last error"])).value = None
                ok += 1
            except Exception as e:
                market.range((row_idx, cols["Last error"])).value = f"{type(e).__name__}: {e}"[:300]
                fail += 1

        main.range(MAIN_CELLS["LastQuotesAt"]).value = now_iso
        if not test_mode:
            if fx.get("USD") is not None and not pd.isna(fx["USD"]):
                main.range(MAIN_CELLS["EurUsd"]).value = float(fx["USD"])
            if fx.get("JPY") is not None and not pd.isna(fx["JPY"]):
                main.range(MAIN_CELLS["EurJpy"]).value = float(fx["JPY"])
            if fx.get("GBP") is not None and not pd.isna(fx["GBP"]):
                main.range(MAIN_CELLS["EurGbp"]).value = float(fx["GBP"])

        status(f"quotes: done. {ok} ok, {fail} failed at {now_iso}.")
    except Exception as e:
        status(f"ERROR: {type(e).__name__}: {e}"[:200])
        raise


# ---------------------------------------------------------------------------
# Job 2: get_quotes — refresh price and P/E columns daily
# ---------------------------------------------------------------------------

def _read_market_symbols(market_ws: Worksheet) -> list[tuple[int, str, str, str]]:
    """Iterate Market data rows. Returns ``[(row_idx, symbol, currency, indexes), ...]``."""
    sym_c = _market_col("Symbol")
    ccy_c = _market_col("Currency")
    idx_c = _market_col("Indexes")
    out: list[tuple[int, str, str, str]] = []
    for r in range(2, market_ws.max_row + 1):
        sym = market_ws.cell(row=r, column=sym_c).value
        if not sym:
            continue
        ccy = market_ws.cell(row=r, column=ccy_c).value or ""
        idxs = market_ws.cell(row=r, column=idx_c).value or ""
        out.append((r, str(sym).strip(), str(ccy), str(idxs)))
    return out


def _pick_test_target(
    market_rows: list[tuple[int, str, str, str]],
    portfolio: set[str],
) -> Optional[tuple[int, str, str, str]]:
    """Test-mode target: top portfolio entry that exists in Market, else first
    BEL20 row, else first Market row."""
    if portfolio:
        # User may have typed "KBC" (root) or "KBC.BR" (full Yahoo) in Main.
        # Market always stores the full Yahoo form. Match either direction.
        for sym in sorted(portfolio):
            for row in market_rows:
                market_sym = row[1].upper()
                market_root = re.sub(r"\.[A-Z]{1,3}$", "", market_sym)
                if market_sym == sym or market_root == sym:
                    return row
    for row in market_rows:
        if "BEL20" in row[3]:
            return row
    return market_rows[0] if market_rows else None


def get_quotes(
    workbook_path: Path,
    test_mode: Optional[bool] = None,
    info_delay: float = 0.25,
    status: Optional[Callable[[str], None]] = None,
) -> int:
    """Refresh quote columns in Market for every Symbol (or one in test mode).

    Sets per-row "Last update (UTC)" on success and "Last error" on failure;
    successful refresh clears any previous error. ``test_mode=None`` reads
    Main!TestMode from the workbook.
    """
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        log.error(f"Workbook {workbook_path} does not exist. Run init-workbook + rebuild-inventory first.")
        return 1

    def _say(msg: str) -> None:
        log.info(msg)
        if status is not None:
            try:
                status(msg)
            except Exception as e:
                log.debug(f"status callback raised: {e}")

    t0 = time.time()
    wb = _load_xl(workbook_path)
    main_ws = wb["Main"]
    market_ws = wb["Market"]
    if test_mode is None:
        test_mode = read_test_mode(main_ws)
    _say(f"get_quotes: starting (test_mode={test_mode})")

    portfolio = read_portfolio_symbols(main_ws)
    market_rows = _read_market_symbols(market_ws)

    if not market_rows:
        log.error("Market sheet is empty. Run rebuild-inventory first.")
        return 1

    if test_mode:
        picked = _pick_test_target(market_rows, portfolio)
        if picked is None:
            log.error("Could not pick a test target.")
            return 1
        targets = [picked]
        _say(f"Test mode: refreshing {targets[0][1]} only")
    else:
        targets = market_rows

    _say("Fetching FX rates")
    fx = get_fx_rates()

    _say(f"Downloading price history for {len(targets)} tickers")
    today = dt.date.today()
    start = today - dt.timedelta(days=int(5.2 * 365))
    end   = today + dt.timedelta(days=1)
    closes = fetch_close_prices([t[1] for t in targets], start, end)

    _say(f"Fetching .info for {len(targets)} tickers (P/E)")

    def info_progress(done: int, total: int, sym: str) -> None:
        if status is not None and (done == 1 or done == total or done % 25 == 0):
            status(f"quotes {done}/{total} — {sym}")

    info_map = fetch_all_info([t[1] for t in targets], delay=info_delay, progress=info_progress)

    _say("Writing quote columns")
    now_iso = dt.datetime.now(dt.UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    cols = {name: _market_col(name) for name in MARKET_COLUMNS}

    successes = 0
    failures = 0
    for (row_idx, sym, ccy, _indexes) in targets:
        try:
            series = closes[sym] if sym in closes.columns else pd.Series(dtype=float)
            for label, delta in LOOKBACKS.items():
                target = today - delta
                native = price_at_or_before(series, target)
                market_ws.cell(row=row_idx, column=cols[f"{label} (EUR)"],
                               value=to_eur(native, ccy, fx))
            info = info_map.get(sym, {})
            market_ws.cell(row=row_idx, column=cols["P/E (TTM)"],   value=info.get("trailingPE"))
            market_ws.cell(row=row_idx, column=cols["Forward P/E"], value=info.get("forwardPE"))
            market_ws.cell(row=row_idx, column=cols["Last update (UTC)"], value=now_iso)
            market_ws.cell(row=row_idx, column=cols["Last error"], value=None)  # clear previous
            successes += 1
        except Exception as e:
            market_ws.cell(row=row_idx, column=cols["Last error"], value=f"{type(e).__name__}: {e}"[:300])
            failures += 1
            log.warning(f"  {sym}: {e}")

    # Update Main metadata
    main_ws[MAIN_CELLS["LastQuotesAt"]] = now_iso
    if not test_mode:
        # Don't overwrite the headline FX values in test mode — they only describe
        # a 1-symbol run which is rarely interesting.
        if fx.get("USD") is not None and not pd.isna(fx["USD"]):
            main_ws[MAIN_CELLS["EurUsd"]] = float(fx["USD"])
        if fx.get("JPY") is not None and not pd.isna(fx["JPY"]):
            main_ws[MAIN_CELLS["EurJpy"]] = float(fx["JPY"])
        if fx.get("GBP") is not None and not pd.isna(fx["GBP"]):
            main_ws[MAIN_CELLS["EurGbp"]] = float(fx["GBP"])

    wb.save(workbook_path)
    duration = time.time() - t0
    _say(f"get_quotes: done in {duration:.0f}s. {successes} ok, {failures} failed.")
    return 0


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


def _cmd_setup_buttons(args) -> int:
    """One-time: convert the workbook to .xlsm and add buttons + xlwings.conf sheet.

    Requires Excel + xlwings. The user still needs to import vba/stocks_picker.bas
    once via Alt+F11 — programmatic VBA injection needs Excel's "Trust access to
    the VBA project object model" setting, which we can't toggle here.

    Flow:
      - Source: whichever workbook --workbook points at (default stocks.xlsx).
      - Target: same basename with .xlsm extension.
      - On success, the source .xlsx is removed (unless source IS already .xlsm).
    """
    import xlwings as xw

    src = Path(args.workbook).resolve()
    if not src.exists():
        log.error(f"Workbook {src} does not exist. Run init-workbook first.")
        return 1
    target = src.with_suffix(".xlsm")

    log.info(f"Opening {src} in Excel (headless)")
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(src))
        main = wb.sheets["Main"]

        # 1. Add xlwings.conf sheet (used by xlwings to find Python interpreter).
        if "xlwings.conf" not in [s.name for s in wb.sheets]:
            conf = wb.sheets.add("xlwings.conf", after=wb.sheets[wb.sheets.count - 1])
        else:
            conf = wb.sheets["xlwings.conf"]
        interpreter = (Path.cwd() / ".venv" / "Scripts" / "python.exe").resolve()
        pythonpath = Path.cwd().resolve()
        conf_rows = [
            ["INTERPRETER_WIN", str(interpreter)],
            ["PYTHONPATH",      str(pythonpath)],
            ["SHOW CONSOLE",    "False"],
        ]
        conf.range("A1").value = conf_rows
        try:
            conf.api.Visible = 2  # xlSheetVeryHidden (only visible via VBA editor)
        except Exception:
            pass

        # 2. Remove any existing buttons we added previously (idempotent).
        for shp in list(main.shapes):
            if shp.name.startswith("StockPicker_"):
                shp.delete()

        # 3. Add two buttons + a Test-mode checkbox in the Jobs area to the
        #    right of column B. Column A (~38 chars ≈ 270 px) and column B
        #    (~60 chars ≈ 430 px) together reach left=700, so buttons live
        #    at left=720+ to avoid hiding either.
        #    Shapes.AddFormControl(type, left, top, w, h):
        #      0 = msoFormControlButton
        #      1 = msoFormControlCheckBox
        sheet_api = main.api
        btn1 = sheet_api.Shapes.AddFormControl(0, 720, 10, 130, 28)
        btn1.Name = "StockPicker_Rebuild"
        btn1.TextFrame.Characters().Text = "Rebuild Inventory"
        btn1.OnAction = "RebuildInventory"
        btn2 = sheet_api.Shapes.AddFormControl(0, 720, 42, 130, 28)
        btn2.Name = "StockPicker_GetQuotes"
        btn2.TextFrame.Characters().Text = "Get Quotes"
        btn2.OnAction = "GetQuotes"

        # 4. SaveAs .xlsm (FileFormat 52 = xlOpenXMLWorkbookMacroEnabled).
        if target.exists() and target != src:
            target.unlink()  # Excel SaveAs refuses to overwrite without prompts in headless mode
        wb.api.SaveAs(str(target), FileFormat=52)
        wb.close()

        # Remove the source .xlsx if it's distinct from the .xlsm target.
        if src != target and src.exists():
            src.unlink()

        log.info(f"setup-buttons: workbook saved as {target}")
        log.info("")
        log.info("Final manual step (one-time):")
        log.info(f"  1. Open {target.name} in Excel.")
        log.info("  2. Press Alt+F11 -> File -> Import File -> select vba/stocks_picker.bas.")
        log.info("  3. Save (the buttons will now respond to clicks).")
        log.info("")
        log.info(r"Also run once: `.\.venv\Scripts\xlwings.exe addin install`")
        return 0
    finally:
        app.quit()


def _cmd_get_quotes(args) -> int:
    """Job 2 — refresh quote columns (prices + P/E) for every Market row."""
    # --test forces test mode; without it, the workbook's TestMode cell decides.
    return get_quotes(
        workbook_path=_resolve_workbook(args.workbook),
        test_mode=True if args.test else None,
        info_delay=args.info_delay,
    )


def _cmd_rebuild_inventory(args) -> int:
    """Job 1 — refresh the Market sheet structural data (no quotes)."""
    # --test forces BEL20 only; without it, the workbook's TestMode cell decides.
    indexes = None
    if not args.test and args.indexes and args.indexes.upper() != "ALL":
        indexes = [s.strip().upper() for s in args.indexes.split(",")]
        unknown = [i for i in indexes if i not in INDEX_WIKI]
        if unknown:
            log.error(f"Unknown index(es): {unknown}")
            log.error(f"Valid: {list(INDEX_WIKI.keys())}")
            return 2
    return rebuild_inventory(
        workbook_path=_resolve_workbook(args.workbook),
        indexes=indexes,
        info_delay=args.info_delay,
        test_mode=True if args.test else None,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    sub = parser.add_subparsers(dest="cmd")

    # init-workbook — create or refresh the persistent workbook
    p_init = sub.add_parser(
        "init-workbook",
        help="Create or refresh stocks.xlsm with Main + Market sheet templates.",
    )
    p_init.add_argument(
        "--workbook", default=str(DEFAULT_WORKBOOK_PATH),
        help=f"Workbook path (default: {DEFAULT_WORKBOOK_PATH})",
    )

    # setup-buttons — one-time: add Excel buttons + xlwings.conf sheet
    p_setup = sub.add_parser(
        "setup-buttons",
        help="One-time: add Excel buttons + xlwings.conf sheet to the workbook.",
    )
    p_setup.add_argument(
        "--workbook", default=str(DEFAULT_WORKBOOK_PATH),
        help=f"Workbook path (default: {DEFAULT_WORKBOOK_PATH})",
    )

    # get-quotes — Job 2 (refresh prices + P/E + per-row last update/error)
    p_quotes = sub.add_parser(
        "get-quotes",
        help="Job 2: refresh quote columns in the Market sheet (run daily).",
    )
    p_quotes.add_argument(
        "--workbook", default=str(DEFAULT_WORKBOOK_PATH),
        help=f"Workbook path (default: {DEFAULT_WORKBOOK_PATH})",
    )
    p_quotes.add_argument(
        "--info-delay", type=float, default=0.25,
        help="Seconds between yfinance .info calls (default: 0.25)",
    )
    p_quotes.add_argument(
        "--test", action="store_true",
        help="Test mode: refresh 1 symbol only (top of Main portfolio or first BEL20).",
    )

    # rebuild-inventory — Job 1 (no quotes; populates Market structural data)
    p_rebuild = sub.add_parser(
        "rebuild-inventory",
        help="Job 1: refresh the Market sheet's structural data (no quote fetch).",
    )
    p_rebuild.add_argument(
        "--workbook", default=str(DEFAULT_WORKBOOK_PATH),
        help=f"Workbook path (default: {DEFAULT_WORKBOOK_PATH})",
    )
    p_rebuild.add_argument(
        "--indexes", default="ALL",
        help="Subset of indexes to refresh (default: ALL).",
    )
    p_rebuild.add_argument(
        "--info-delay", type=float, default=0.25,
        help="Seconds between yfinance .info calls (default: 0.25)",
    )
    p_rebuild.add_argument(
        "--test", action="store_true",
        help="Test mode: BEL20 only.",
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
    known_subcommands = {"init-workbook", "setup-buttons", "rebuild-inventory", "get-quotes", "run", "-h", "--help"}
    raw_args = sys.argv[1:]
    if raw_args and raw_args[0] not in known_subcommands:
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
    if args.cmd == "setup-buttons":
        return _cmd_setup_buttons(args)
    if args.cmd == "rebuild-inventory":
        return _cmd_rebuild_inventory(args)
    if args.cmd == "get-quotes":
        return _cmd_get_quotes(args)

    parser.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
