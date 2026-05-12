"""Offline unit tests for stocks_report. Run with: python -m pytest tests/ -v"""

import datetime as dt
import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from stocks_report import (  # noqa: E402
    MAIN_CELLS,
    MARKET_COLUMNS,
    PORTFOLIO_FIRST_ROW,
    _extract_dataroma_tickers,
    _owned_for,
    _parse_bel20_ticker,
    aggregate_constituents,
    init_workbook,
    normalize_ticker,
    price_at_or_before,
    read_portfolio_symbols,
    to_eur,
)


# ---------------------------------------------------------------------------
# normalize_ticker
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("raw,index_name,expected", [
    # S&P 500: dot becomes hyphen
    ("AAPL",     "SP500",     "AAPL"),
    ("BRK.B",    "SP500",     "BRK-B"),
    ("BF.B",     "SP500",     "BF-B"),
    # DAX
    ("SAP",      "DAX",       "SAP.DE"),
    ("SAP.DE",   "DAX",       "SAP.DE"),  # idempotent
    # CAC 40
    ("AIR",      "CAC40",     "AIR.PA"),
    # BEL 20
    ("KBC",      "BEL20",     "KBC.BR"),
    ("APAM.AS",  "BEL20",     "APAM.AS"),  # Amsterdam-listed, must not be re-suffixed
    # EuroStoxx 50: Wikipedia provides fully-qualified Yahoo tickers — pass through
    ("ADS.DE",   "ESTOXX50",  "ADS.DE"),
    ("ADYEN.AS", "ESTOXX50",  "ADYEN.AS"),
    ("AI.PA",    "ESTOXX50",  "AI.PA"),
    ("ISP.MI",   "ESTOXX50",  "ISP.MI"),  # Milan
    ("IBE.MC",   "ESTOXX50",  "IBE.MC"),  # Madrid
    # FTSE 100: trailing dot stripped, .L appended
    ("ULVR",     "FTSE100",   "ULVR.L"),
    ("RR.",      "FTSE100",   "RR.L"),
    ("BHP.",     "FTSE100",   "BHP.L"),
    # Nikkei 225
    ("7203",     "NIKKEI225", "7203.T"),
    # Robustness: footnote markers, lowercase
    ("AAPL[1]",  "SP500",     "AAPL"),
    ("aapl",     "SP500",     "AAPL"),
    ("  AAPL ",  "SP500",     "AAPL"),
])
def test_normalize_ticker(raw, index_name, expected):
    assert normalize_ticker(raw, index_name) == expected


# ---------------------------------------------------------------------------
# _parse_bel20_ticker — Wikipedia BEL 20 cell format is "Euronext X:\xa0SYMBOL"
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("cell,expected", [
    ("Euronext Brussels:\xa0ABI",     "ABI.BR"),
    ("Euronext Brussels:\xa0ARGX",    "ARGX.BR"),
    ("Euronext Amsterdam:\xa0APAM",   "APAM.AS"),
    ("Euronext Brussels: KBC",        "KBC.BR"),  # plain space
    ("euronext brussels:\xa0umi",     "UMI.BR"),  # case-insensitive
    ("garbage value",                 ""),         # no match
    ("",                              ""),
])
def test_parse_bel20_ticker(cell, expected):
    assert _parse_bel20_ticker(cell) == expected


# ---------------------------------------------------------------------------
# _extract_dataroma_tickers — pulls /m/stock.php?sym=X links from HTML
# ---------------------------------------------------------------------------

def test_extract_dataroma_tickers_basic():
    html = (
        '<tr><td class="sym"><a href="/m/stock.php?sym=AMZN">AMZN</a></td>'
        '<td class="stock"><a href="/m/stock.php?sym=AMZN">Amazon.com Inc.</a></td></tr>'
        '<tr><td class="sym"><a href="/m/stock.php?sym=MSFT">MSFT</a></td></tr>'
    )
    assert _extract_dataroma_tickers(html) == {"AMZN", "MSFT"}


def test_extract_dataroma_tickers_dual_form_for_dot_classes():
    """Dataroma writes BRK.B; SP500 in this report uses BRK-B. Both forms must come back."""
    html = '<a href="/m/stock.php?sym=BRK.B">BRK.B</a> <a href="/m/stock.php?sym=BF.B">BF.B</a>'
    out = _extract_dataroma_tickers(html)
    assert {"BRK.B", "BRK-B", "BF.B", "BF-B"} <= out


def test_extract_dataroma_tickers_empty_on_no_match():
    assert _extract_dataroma_tickers("<html>no links here</html>") == set()


def test_extract_dataroma_tickers_ignores_unrelated_links():
    html = '<a href="/m/managers.php">Managers</a> <a href="/m/stock.php?sym=AAPL">AAPL</a>'
    assert _extract_dataroma_tickers(html) == {"AAPL"}


# ---------------------------------------------------------------------------
# aggregate_constituents — dedups by Symbol, joins Indexes
# ---------------------------------------------------------------------------

def _df(rows):
    return pd.DataFrame(rows, columns=["Symbol", "Name", "Index"])


def test_aggregate_constituents_dedups_by_symbol():
    """A stock in two indexes yields one row with both indexes listed."""
    dax = _df([("SAP.DE", "SAP", "DAX"), ("ALV.DE", "Allianz", "DAX")])
    estoxx = _df([("SAP.DE", "SAP SE", "ESTOXX50"), ("AIR.PA", "Airbus", "ESTOXX50")])
    out = aggregate_constituents([dax, estoxx])
    rows = {r["Symbol"]: r for _, r in out.iterrows()}
    assert rows["SAP.DE"]["Indexes"] == "DAX, ESTOXX50"
    assert rows["ALV.DE"]["Indexes"] == "DAX"
    assert rows["AIR.PA"]["Indexes"] == "ESTOXX50"
    assert len(out) == 3  # SAP collapsed into one row


def test_aggregate_constituents_preserves_first_seen_name():
    dax = _df([("SAP.DE", "SAP", "DAX")])
    estoxx = _df([("SAP.DE", "SAP SE", "ESTOXX50")])  # different Name spelling
    out = aggregate_constituents([dax, estoxx])
    assert out.iloc[0]["Name"] == "SAP"  # DAX listed first → its Name wins


def test_aggregate_constituents_dedups_repeated_index():
    """Defensive: if the same Symbol+Index appears twice, Indexes lists it once."""
    a = _df([("AAPL", "Apple", "SP500"), ("AAPL", "Apple", "SP500")])
    out = aggregate_constituents([a])
    assert out.iloc[0]["Indexes"] == "SP500"


def test_aggregate_constituents_empty_input():
    out = aggregate_constituents([])
    assert list(out.columns) == ["Symbol", "Name", "Indexes"]
    assert len(out) == 0


# ---------------------------------------------------------------------------
# init_workbook — creates Main + Market layout, idempotent
# ---------------------------------------------------------------------------

def test_init_workbook_creates_main_and_market(tmp_path):
    """Fresh-create writes both sheets with the expected headers and named cells."""
    from openpyxl import load_workbook
    path = tmp_path / "stocks.xlsx"
    init_workbook(path)
    assert path.exists()
    wb = load_workbook(path)
    assert wb.sheetnames == ["Main", "Market"]
    # Market headers match MARKET_COLUMNS
    market = wb["Market"]
    headers = [market.cell(row=1, column=i).value for i in range(1, len(MARKET_COLUMNS) + 1)]
    assert headers == MARKET_COLUMNS
    # Main has the title and named cells initialized
    main = wb["Main"]
    assert main["A1"].value == "Stock Picker"
    assert main[MAIN_CELLS["TestMode"]].value == "FALSE"


def test_init_workbook_is_idempotent(tmp_path):
    """Running init twice does not clobber data the user added."""
    from openpyxl import load_workbook
    path = tmp_path / "stocks.xlsx"
    init_workbook(path)

    # User adds a portfolio entry and toggles test mode.
    wb = load_workbook(path)
    main = wb["Main"]
    main.cell(row=PORTFOLIO_FIRST_ROW, column=1, value="KBC.BR")
    main.cell(row=PORTFOLIO_FIRST_ROW, column=2, value=10)
    main[MAIN_CELLS["TestMode"]] = "TRUE"
    market = wb["Market"]
    market["A2"] = "FAKE.XX"   # pretend Market has data already
    wb.save(path)

    # Run init again — user data MUST survive.
    init_workbook(path)
    wb2 = load_workbook(path)
    assert wb2["Main"].cell(row=PORTFOLIO_FIRST_ROW, column=1).value == "KBC.BR"
    assert wb2["Main"].cell(row=PORTFOLIO_FIRST_ROW, column=2).value == 10
    assert wb2["Main"][MAIN_CELLS["TestMode"]].value == "TRUE"
    assert wb2["Market"]["A2"].value == "FAKE.XX"


def test_init_workbook_market_columns_unique():
    """No duplicate column names — otherwise lookups by name would be ambiguous."""
    assert len(MARKET_COLUMNS) == len(set(MARKET_COLUMNS))


# ---------------------------------------------------------------------------
# _owned_for & read_portfolio_symbols
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("symbol,portfolio,expected", [
    ("KBC.BR",  {"KBC.BR"},        "Yes"),
    ("KBC.BR",  {"KBC"},           "Yes"),   # suffix-stripped match
    ("AAPL",    {"AAPL"},          "Yes"),
    ("AAPL",    {"GOOGL"},         "No"),
    ("BRK-B",   {"BRK-B"},         "Yes"),   # hyphen form
    ("KBC.BR",  set(),             "No"),
])
def test_owned_for(symbol, portfolio, expected):
    assert _owned_for(symbol, portfolio) == expected


def test_read_portfolio_symbols_skips_blanks_and_normalizes(tmp_path):
    """Reads col A of the portfolio range, uppercase + strip, skip None/empty."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Main"
    ws.cell(row=PORTFOLIO_FIRST_ROW,     column=1, value="  kbc.br  ")  # whitespace + lower
    ws.cell(row=PORTFOLIO_FIRST_ROW + 1, column=1, value=None)            # blank
    ws.cell(row=PORTFOLIO_FIRST_ROW + 2, column=1, value="AAPL")
    ws.cell(row=PORTFOLIO_FIRST_ROW + 3, column=1, value="")              # empty string
    assert read_portfolio_symbols(ws) == {"KBC.BR", "AAPL"}


def test_read_portfolio_symbols_ignores_main_section_labels(tmp_path):
    """Regression: a fresh init_workbook must have an empty portfolio set.

    Previously the portfolio range overlapped with the Jobs/Status/Metadata
    section labels in column A, so every label was being parsed as a "symbol".
    """
    path = tmp_path / "stocks.xlsx"
    init_workbook(path)
    from openpyxl import load_workbook
    wb = load_workbook(path)
    assert read_portfolio_symbols(wb["Main"]) == set()


# ---------------------------------------------------------------------------
# to_eur
# ---------------------------------------------------------------------------

def test_to_eur_basic_conversions():
    fx = {"USD": 1.07, "JPY": 165.0, "GBP": 0.85, "EUR": 1.0, "GBp": 0.85}
    assert to_eur(100.0, "USD", fx) == pytest.approx(100 / 1.07)
    assert to_eur(1000.0, "JPY", fx) == pytest.approx(1000 / 165.0)
    assert to_eur(50.0, "GBP", fx) == pytest.approx(50 / 0.85)
    assert to_eur(50.0, "EUR", fx) == 50.0


def test_to_eur_pence():
    """GBp = pence; convert to pounds first then to EUR."""
    fx = {"GBP": 0.85, "GBp": 0.85}
    # 200 pence = £2 = 2/0.85 EUR
    assert to_eur(200.0, "GBp", fx) == pytest.approx(2.0 / 0.85)


def test_to_eur_handles_none_price():
    assert to_eur(None, "USD", {"USD": 1.07}) is None


def test_to_eur_handles_missing_rate():
    fx = {"EUR": 1.0}
    assert to_eur(100.0, "JPY", fx) is None


def test_to_eur_handles_nan_rate():
    fx = {"USD": float("nan")}
    assert to_eur(100.0, "USD", fx) is None


# ---------------------------------------------------------------------------
# price_at_or_before
# ---------------------------------------------------------------------------

def _make_series():
    idx = pd.to_datetime([
        "2026-04-27", "2026-04-28", "2026-04-29",
        "2026-04-30", "2026-05-01", "2026-05-04",
    ])
    return pd.Series([100, 101, 102, 103, 104, 107], index=idx)


def test_price_at_or_before_weekend_rollback():
    s = _make_series()
    # Sat 2026-05-02 → fall back to Fri May 1 (104)
    assert price_at_or_before(s, dt.date(2026, 5, 2)) == 104.0
    # Sun 2026-05-03 → also fall back to Fri May 1
    assert price_at_or_before(s, dt.date(2026, 5, 3)) == 104.0


def test_price_at_or_before_exact_match():
    s = _make_series()
    assert price_at_or_before(s, dt.date(2026, 5, 4)) == 107.0
    assert price_at_or_before(s, dt.date(2026, 4, 27)) == 100.0


def test_price_at_or_before_target_after_history():
    s = _make_series()
    # If target is in the future, just return the last known close
    assert price_at_or_before(s, dt.date(2030, 1, 1)) == 107.0


def test_price_at_or_before_target_before_history():
    s = _make_series()
    assert price_at_or_before(s, dt.date(2020, 1, 1)) is None


def test_price_at_or_before_empty_series():
    assert price_at_or_before(pd.Series(dtype=float), dt.date(2026, 5, 4)) is None


def test_price_at_or_before_none_series():
    assert price_at_or_before(None, dt.date(2026, 5, 4)) is None


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
