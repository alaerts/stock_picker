"""Live-network integration tests.

These tests catch the class of failure where an upstream source changes its
page structure and our parser silently breaks or fails noisily — for
example, Wikipedia removing the Nikkei 225 constituent table (which
produced the rebuild-inventory button error in May 2026 and which the
offline unit tests could not have detected).

Run them with:
    pytest --integration -v

They are SKIPPED by default so the routine pytest run stays offline + fast.
"""
import re

import pytest

from stocks_report import (
    DATAROMA_ACTIVIST_CODES,
    ETF_LIST,
    FX_PAIRS,
    INDEX_WIKI,
    fetch_close_prices,
    fetch_dataroma_activist_aggregate,
    fetch_dataroma_tickers,
    fetch_yahoo_screener,
    get_fx_history,
    get_fx_rates,
    get_index_constituents,
    _yahoo_screener_session,
)


pytestmark = pytest.mark.integration


# ---------------------------------------------------------------------------
# Wikipedia: one parametrized test per index
# ---------------------------------------------------------------------------

# (index, min_count, allowed_suffixes). min_count is a conservative floor —
# real counts fluctuate (index reshuffles, recent additions), so we don't
# pin to an exact number.
INDEX_EXPECTATIONS = [
    ("BEL20",     18,  {".BR", ".AS"}),
    # CAC40 has a few non-Paris primary listings (ArcelorMittal MT.AS, etc.).
    ("CAC40",     35,  {".PA", ".AS"}),
    # DAX has Airbus (AIR.PA) — its primary Yahoo listing is Paris.
    ("DAX",       35,  {".DE", ".PA"}),
    ("FTSE100",   90,  {".L"}),
    ("NIKKEI225", 200, {".T"}),
    ("SP500",     480, {""}),  # SP500 symbols have no exchange suffix (AAPL, BRK-B)
    # ESTOXX50 covers the whole Eurozone, picks up listings from Brussels too.
    ("ESTOXX50",  45,  {".DE", ".PA", ".AS", ".MI", ".MC", ".HE", ".IR", ".BR"}),
]


@pytest.mark.parametrize("index,min_count,suffixes", INDEX_EXPECTATIONS,
                         ids=[x[0] for x in INDEX_EXPECTATIONS])
def test_index_constituents_live(index, min_count, suffixes):
    """Each index's Wikipedia page must still yield enough usable tickers.

    Catches: page restructures (table removed / renamed), regex drift,
    Wikipedia auth/redirect changes.
    """
    assert index in INDEX_WIKI, f"{index} missing from INDEX_WIKI"
    df = get_index_constituents(index)
    assert len(df) >= min_count, (
        f"{index}: got {len(df)} constituents, expected >= {min_count}. "
        f"Did the source page restructure?"
    )
    # Every symbol must end with one of the allowed suffixes (empty string
    # means "no suffix", i.e. SP500's AAPL / BRK-B style).
    bad = []
    for sym in df["Symbol"]:
        if not any((sym.endswith(s) if s else "." not in sym) for s in suffixes):
            bad.append(sym)
    assert not bad, f"{index}: {len(bad)} symbols with unexpected suffix: {bad[:8]}"


def test_all_indexes_yield_unique_symbols():
    """Across the 7 indexes, the aggregate symbol set should be ~975 unique
    after dedup (stocks like SAP.DE appear in DAX + ESTOXX50). If this drops
    drastically, multiple sources broke at once."""
    all_syms: set[str] = set()
    counts = {}
    for index in INDEX_WIKI:
        df = get_index_constituents(index)
        counts[index] = len(df)
        all_syms.update(df["Symbol"])
    assert len(all_syms) >= 900, (
        f"Aggregate symbol count {len(all_syms)} unexpectedly low. "
        f"Per-index counts: {counts}"
    )


# ---------------------------------------------------------------------------
# Yahoo screener API — cookie + crumb auth, paginated screener fetch
# ---------------------------------------------------------------------------

def test_yahoo_screener_session_returns_crumb():
    """Yahoo's auth flow (fc.yahoo.com → /research-hub → /v1/test/getcrumb)
    must still produce a usable session+crumb pair."""
    auth = _yahoo_screener_session()
    assert auth is not None, "Yahoo screener auth failed — cookie/crumb flow broke?"
    sess, crumb = auth
    assert sess is not None
    assert crumb and len(crumb) < 64, f"Suspicious crumb: {crumb!r}"


def test_yahoo_screener_recent_52_week_highs():
    """recent_52_week_highs is one of two screeners we depend on."""
    auth = _yahoo_screener_session()
    assert auth is not None
    sess, crumb = auth
    tickers = fetch_yahoo_screener("recent_52_week_highs", sess, crumb)
    assert len(tickers) >= 50, f"recent_52_week_highs returned {len(tickers)} (expected ~100)"
    # All symbols should look like equity tickers (no =F / =X / ^).
    for t in tickers:
        assert not t.startswith("^"), f"Index symbol leaked through: {t}"
        assert not t.endswith("=F"), f"Futures symbol leaked through: {t}"
        assert not t.endswith("=X"), f"FX symbol leaked through: {t}"


def test_yahoo_screener_berkshire_contains_real_holdings():
    """Berkshire's undocumented scrIds must still return AAPL / BAC / KO."""
    auth = _yahoo_screener_session()
    assert auth is not None
    sess, crumb = auth
    tickers = fetch_yahoo_screener("top_stocks_owned_by_warren_buffet", sess, crumb)
    assert len(tickers) >= 50, f"Berkshire returned {len(tickers)} (expected ~100)"
    # At least 3 of these long-time holdings must be present. If all are gone,
    # the API is returning a different dataset.
    long_term_holdings = {"AAPL", "BAC", "KO", "AXP", "OXY", "CVX", "MCO", "KHC"}
    hits = long_term_holdings & set(tickers)
    assert len(hits) >= 3, (
        f"Only {len(hits)} known Berkshire holdings found in result ({hits}). "
        f"Sample of what came back: {sorted(tickers)[:15]}. "
        "Berkshire scrIds may have changed or been replaced with another dataset."
    )


# ---------------------------------------------------------------------------
# Dataroma — plain HTML scraping for the 4 super-investor lists
# ---------------------------------------------------------------------------

DATAROMA_PAGES = [
    ("/m/g/portfolio.php",                "most-held"),
    ("/m/g/portfolio_b.php?q=q&o=c",      "top-buys"),
    ("/m/g/portfolio_s.php?q=q&o=c",      "top-sells"),
]


@pytest.mark.parametrize("path,label", DATAROMA_PAGES, ids=[x[1] for x in DATAROMA_PAGES])
def test_dataroma_list_pages_return_tickers(path, label):
    tickers = fetch_dataroma_tickers(path)
    assert len(tickers) >= 50, (
        f"dataroma {label} ({path}) returned {len(tickers)} (expected ~100). "
        "URL or HTML structure changed?"
    )
    # Sample should look like equity tickers (letters/digits/.-)
    for t in list(tickers)[:5]:
        assert re.fullmatch(r"[A-Z][A-Z0-9.-]{0,11}", t), f"Bad ticker shape: {t!r}"


def test_dataroma_activist_aggregate_covers_all_seven_managers():
    """The 7 activist-manager holdings.php pages together should yield ~80+
    distinct tickers. If the count collapses, one or more manager codes
    (ic / psc / VA / TF / ENG / tp / tci) are no longer valid."""
    tickers = fetch_dataroma_activist_aggregate()
    assert len(tickers) >= 40, (
        f"Activist aggregate returned {len(tickers)} (expected ~80+). "
        f"One or more of {DATAROMA_ACTIVIST_CODES} may be invalid."
    )


# ---------------------------------------------------------------------------
# ETFs — yfinance has price history for each curated ticker
# ---------------------------------------------------------------------------

def test_etf_tickers_resolve_on_yahoo():
    """Bulk-download last week of prices for every ETF in ETF_LIST. Any
    ticker that doesn't resolve on Yahoo (typo, delisted, renamed) won't
    appear in the returned columns and we surface it as a failure."""
    import datetime as dt
    tickers = [sym for sym, _, _ in ETF_LIST]
    end = dt.date.today() + dt.timedelta(days=1)
    start = end - dt.timedelta(days=14)
    closes = fetch_close_prices(tickers, start, end)
    missing = [t for t in tickers if t not in closes.columns]
    # Allow a small failure budget — Yahoo occasionally drops a ticker for
    # a session. Hard-fail if more than ~10% are missing.
    assert len(missing) <= max(2, len(tickers) // 10), (
        f"{len(missing)}/{len(tickers)} ETF tickers did not resolve on Yahoo: "
        f"{missing}. Curated list may need updating."
    )


# ---------------------------------------------------------------------------
# FX rates — yfinance pulls EURUSD=X, EURJPY=X, EURGBP=X
# ---------------------------------------------------------------------------

def test_fx_rates_are_plausible():
    """Each currency rate (against EUR) must come back as a sane positive
    number. Catches: yfinance breakage, FX symbol rename, network."""
    rates = get_fx_rates()
    bounds = {
        "USD": (0.5, 2.0),
        "JPY": (80.0, 300.0),
        "GBP": (0.5, 1.5),
        "CHF": (0.5, 1.5),
    }
    for ccy, (lo, hi) in bounds.items():
        rate = rates.get(ccy)
        assert rate is not None, f"EUR/{ccy} missing from rates"
        assert lo <= rate <= hi, (
            f"EUR/{ccy} = {rate}, outside plausible range [{lo}, {hi}]"
        )


def test_fx_history_returns_series_for_each_pair():
    """get_fx_history must return a non-empty Close-price Series per pair
    so the Currencies sheet can resolve every lookback."""
    history = get_fx_history()
    for ccy in FX_PAIRS:
        assert ccy in history, f"{ccy} missing from FX history"
        series = history[ccy]
        assert len(series) >= 100, (
            f"EUR/{ccy} returned only {len(series)} datapoints — expected ~1300"
        )
