"""Offline unit tests for stocks_report. Run with: python -m pytest tests/ -v"""

import datetime as dt
import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from stocks_report import (  # noqa: E402
    CURRENCY_HEADERS,
    ETF_LIST,
    FX_PAIRS,
    MAIN_CELLS,
    MARKET_COLUMNS,
    MONTHLY_LOSERS_SHEET_NAME,
    MONTHLY_MOVERS_LEGACY_NAME,
    MONTHLY_WINNERS_SHEET_NAME,
    PORTFOLIO_FIRST_ROW,
    RANKING_HEADERS,
    SCHEMA_VERSION,
    VERSION_HISTORY,
    _append_error_log,
    _compute_monthly_losers,
    _compute_monthly_winners,
    _currency_rows,
    _ensure_help_sheet_versions,
    _extract_dataroma_tickers,
    _migrate_monthly_movers_to_winners_openpyxl,
    _owned_for,
    _parse_bel20_ticker,
    _parse_nikkei225_components,
    _pct_change,
    _pct_column_pairs,
    _pick_test_target,
    _resolve_workbook,
    aggregate_constituents,
    get_etfs,
    init_workbook,
    normalize_ticker,
    price_at_or_before,
    read_portfolio_symbols,
    read_test_mode,
    to_eur,
    yahoo_quote_url,
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
    # Cross-index dedup — regression for the 2026-05-13 batch of bogus rows:
    ("AIR.PA",   "DAX",       "AIR.PA"),   # Airbus in DAX; was → AIR.PA.DE
    ("MT.AS",    "CAC40",     "MT.AS"),    # ArcelorMittal in CAC40; was → MT.AS.PA
    ("ADS.DE",   "ESTOXX50",  "ADS.DE"),   # Adidas in ESTOXX50; suffix already present
    # FTSE 100 sub-class shares use a hyphen on Yahoo (like SP500's BRK-B):
    ("BT.A",     "FTSE100",   "BT-A.L"),   # was → BT.A.L
    ("BT-A.L",   "FTSE100",   "BT-A.L"),   # idempotent
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
    # Nikkei 225 — alphanumeric tickers like 543A.T (Archion) are valid
    ("7203",     "NIKKEI225", "7203.T"),
    ("543A",     "NIKKEI225", "543A.T"),
    ("285A",     "NIKKEI225", "285A.T"),
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
# _parse_nikkei225_components — bullet-list scraping (no table on the page)
# ---------------------------------------------------------------------------

_NIKKEI_FIXTURE = """
<h2 id="Components">Components</h2>
<ul>
  <li>Energy (0.30%)</li>
  <li>Materials (7.00%)</li>
</ul>
<h3>Air transport</h3>
<ul>
  <li>ANA Holdings Inc. (TYO: 9202)</li>
  <li>Japan Airlines Co., Ltd. (TYO: 9201)</li>
</ul>
<h3>Automotive</h3>
<ul>
  <li>Archion Corp. (TYO: 543A)</li>
  <li>Honda Motor Co., Ltd. (TYO: 7267)</li>
  <li>M3 Inc. (TYO: 2413) (Parent company for MDLinx)</li>
</ul>
<h2 id="Statistics">Statistics</h2>
<li>This must not match (TYO: 9999)</li>
"""


def test_parse_nikkei225_components_extracts_pairs():
    # min_count=0 bypasses the production safety floor since the fixture is tiny.
    df = _parse_nikkei225_components(_NIKKEI_FIXTURE, min_count=0)
    pairs = list(zip(df["Symbol"], df["Name"]))
    # 5 stocks; sector summary lines and the post-Statistics li are excluded.
    assert pairs == [
        ("9202.T", "ANA Holdings Inc."),
        ("9201.T", "Japan Airlines Co., Ltd."),
        ("543A.T", "Archion Corp."),
        ("7267.T", "Honda Motor Co., Ltd."),
        ("2413.T", "M3 Inc."),
    ]
    assert (df["Index"] == "NIKKEI225").all()


def test_parse_nikkei225_components_ignores_trailing_annotations():
    """Entries like 'Foo (TYO: 9147)(Holding company for Foo)' must still parse."""
    html = '<h2 id="Components">x</h2><li>Foo Corp. (TYO: 9147)(Holding company for Foo)</li><h2 id="Statistics">x</h2>'
    df = _parse_nikkei225_components(html, min_count=0)
    assert list(df["Symbol"]) == ["9147.T"]
    assert list(df["Name"]) == ["Foo Corp."]


def test_parse_nikkei225_components_raises_when_no_components_section():
    with pytest.raises(RuntimeError, match="Components section"):
        _parse_nikkei225_components("<html>no components header</html>")


def test_parse_nikkei225_components_raises_when_too_few_results():
    """Safety net: if a future page change makes the parser yield only a handful
    of entries, surface as an error rather than silently writing 5 rows."""
    html = '<h2 id="Components">x</h2><li>Foo (TYO: 9001)</li><h2 id="Statistics">x</h2>'
    with pytest.raises(RuntimeError, match="Wikipedia layout likely changed"):
        _parse_nikkei225_components(html)  # uses default min_count=150


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
    assert wb.sheetnames == ["Main", "Market", "Help"]
    # Market headers match MARKET_COLUMNS
    market = wb["Market"]
    headers = [market.cell(row=1, column=i).value for i in range(1, len(MARKET_COLUMNS) + 1)]
    assert headers == MARKET_COLUMNS
    # Main has the title; TestMode cell is intentionally blank on fresh init
    # (the checkbox added by setup-buttons writes TRUE/FALSE into it).
    main = wb["Main"]
    assert main["A1"].value == "Stock Picker"
    assert main[MAIN_CELLS["TestMode"]].value is None


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


def test_init_workbook_preserves_user_cosmetic_edits(tmp_path):
    """User edits to label text and column widths must survive a re-run."""
    from openpyxl import load_workbook
    path = tmp_path / "stocks.xlsx"
    init_workbook(path)

    # Pretend the user renamed the title and widened column A.
    wb = load_workbook(path)
    wb["Main"]["A1"] = "Patrick's Stock Picker"
    wb["Main"].column_dimensions["A"].width = 100
    wb["Main"]["A3"] = "My Custom Jobs Section"  # custom section label
    wb.save(path)

    init_workbook(path)  # re-run on existing file
    wb2 = load_workbook(path)
    assert wb2["Main"]["A1"].value == "Patrick's Stock Picker"
    assert wb2["Main"]["A3"].value == "My Custom Jobs Section"
    assert wb2["Main"].column_dimensions["A"].width == 100


def test_init_workbook_preserves_cleared_labels(tmp_path):
    """If the user CLEARED a label (set to None), a re-run must not restore it.

    This is the trap that caught us on 2026-05-13 — the prior _set helper
    rewrote any cell whose value was None, which silently undid the user's
    deliberate clearances of A3 and A5.
    """
    from openpyxl import load_workbook
    path = tmp_path / "stocks.xlsx"
    init_workbook(path)

    # User clears two section labels.
    wb = load_workbook(path)
    wb["Main"]["A3"] = None
    wb["Main"]["A5"] = None
    wb.save(path)

    init_workbook(path)  # re-run on existing file
    wb2 = load_workbook(path)
    assert wb2["Main"]["A3"].value is None, "A3 was restored after re-run; user cleared it"
    assert wb2["Main"]["A5"].value is None, "A5 was restored after re-run; user cleared it"


# ---------------------------------------------------------------------------
# ETF list + get_etfs()
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# _append_error_log — appends a timestamped block to the log file
# ---------------------------------------------------------------------------

def test_append_error_log_creates_file(tmp_path):
    log = tmp_path / "stocks_errors.log"
    _append_error_log(log, "RuntimeError", "boom", "Traceback line 1\nline 2\n")
    text = log.read_text(encoding="utf-8")
    assert "RuntimeError: boom" in text
    assert "Traceback line 1" in text
    assert "line 2" in text
    # ISO timestamp prefix should be present
    assert "T" in text and "Z" in text


def test_append_error_log_appends_to_existing(tmp_path):
    log = tmp_path / "stocks_errors.log"
    _append_error_log(log, "ValueError", "first", "tb1\n")
    _append_error_log(log, "TypeError", "second", "tb2\n")
    text = log.read_text(encoding="utf-8")
    # Both entries present
    assert "ValueError: first" in text
    assert "TypeError: second" in text
    # Second is after first
    assert text.index("first") < text.index("second")


def test_append_error_log_swallows_io_errors(tmp_path):
    """Log helper must never raise — its caller is already handling an
    exception and a follow-on crash would mask the original error."""
    # Pass a nonexistent parent directory; the call must NOT raise.
    bogus = tmp_path / "no" / "such" / "dir" / "errors.log"
    _append_error_log(bogus, "RuntimeError", "boom", "tb\n")  # should not raise


def test_etf_list_has_one_per_index_plus_sector_and_country():
    """Sanity: 7 index ETFs + 11 sector ETFs + 19+ country ETFs."""
    assert len(ETF_LIST) >= 30, f"ETF_LIST shrank to {len(ETF_LIST)}; was 37+"
    labels = [label for _, _, label in ETF_LIST]
    # Must cover every tracked index
    for idx in ("SP500", "NIKKEI225", "FTSE100", "DAX", "CAC40", "BEL20", "ESTOXX50"):
        assert any(f"ETF — {idx}" == lbl for lbl in labels), f"No ETF for {idx}"
    # Must have at least one sector ETF labelled "Utilities" (user explicitly named it)
    assert any("Sector: Utilities" in lbl for lbl in labels)
    # Must have several country ETFs
    country_count = sum(1 for lbl in labels if lbl.startswith("ETF — Country:"))
    assert country_count >= 15


def test_etf_tickers_are_unique():
    tickers = [sym for sym, _, _ in ETF_LIST]
    assert len(tickers) == len(set(tickers)), "Duplicate ETF tickers"


def test_get_etfs_returns_constituent_shape():
    df = get_etfs()
    assert list(df.columns) == ["Symbol", "Name", "Index"]
    assert len(df) == len(ETF_LIST)
    # Every row's Index label starts with "ETF — "
    assert df["Index"].str.startswith("ETF —").all()


def test_get_etfs_aggregates_with_index_data():
    """ETFs flow through aggregate_constituents alongside index DataFrames."""
    import pandas as pd
    index_df = pd.DataFrame(
        [("AAPL", "Apple", "SP500"), ("KBC.BR", "KBC", "BEL20")],
        columns=["Symbol", "Name", "Index"],
    )
    combined = aggregate_constituents([index_df, get_etfs()])
    syms = set(combined["Symbol"])
    assert "AAPL" in syms
    assert "SPY" in syms  # known ETF
    assert "XLU" in syms  # the Utilities sector ETF specifically


@pytest.mark.parametrize("symbol,expected", [
    ("AAPL",    "https://finance.yahoo.com/quote/AAPL"),
    ("KBC.BR",  "https://finance.yahoo.com/quote/KBC.BR"),
    ("BRK-B",   "https://finance.yahoo.com/quote/BRK-B"),
    ("7203.T",  "https://finance.yahoo.com/quote/7203.T"),
    ("543A.T",  "https://finance.yahoo.com/quote/543A.T"),
])
def test_yahoo_quote_url(symbol, expected):
    assert yahoo_quote_url(symbol) == expected


# ---------------------------------------------------------------------------
# % change helpers
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("today,past,expected", [
    (110, 100, 0.10),       # +10%
    (90,  100, -0.10),      # -10%
    (100, 100, 0.0),
    (100.5, 100, 0.005),
    (None, 100, None),       # missing today
    (100, None, None),       # missing past
    (100, 0,   None),        # divide-by-zero
    (100, "x", None),        # bad type
])
def test_pct_change(today, past, expected):
    out = _pct_change(today, past)
    if expected is None:
        assert out is None
    else:
        assert out == pytest.approx(expected)


def test_pct_column_pairs_covers_every_non_today_lookback():
    """Exactly one (label, column_name) pair per lookback except Today."""
    pairs = _pct_column_pairs()
    expected_labels = ["1D ago", "1W ago", "1M ago", "6M ago", "1Y ago", "5Y ago"]
    assert [lbl for lbl, _ in pairs] == expected_labels
    # Each column name must exist in MARKET_COLUMNS so writes don't KeyError.
    for _label, col in pairs:
        assert col in MARKET_COLUMNS, f"{col!r} missing from MARKET_COLUMNS"


def test_market_columns_includes_pct_columns_interleaved():
    """After every '<X> ago (EUR)' column the next column must be '<X> %'."""
    for i, name in enumerate(MARKET_COLUMNS):
        if name.endswith(" ago (EUR)"):
            label_short = name.replace(" ago (EUR)", "")
            assert MARKET_COLUMNS[i + 1] == f"{label_short} %", (
                f"After {name!r}, expected {label_short!r} %, got "
                f"{MARKET_COLUMNS[i+1]!r}"
            )


# ---------------------------------------------------------------------------
# Schema version, VERSION_HISTORY, _resolve_workbook
# ---------------------------------------------------------------------------

def test_schema_version_string_format():
    """v01, v02, ... — letter v then digits."""
    assert SCHEMA_VERSION.startswith("v")
    assert SCHEMA_VERSION[1:].isdigit()


def test_version_history_includes_current_schema():
    """Every released SCHEMA_VERSION must have a VERSION_HISTORY entry."""
    versions = {v for v, _date, _msg in VERSION_HISTORY}
    assert SCHEMA_VERSION in versions, (
        f"SCHEMA_VERSION={SCHEMA_VERSION!r} missing from VERSION_HISTORY"
    )


def test_version_history_unique_keys():
    versions = [v for v, _date, _msg in VERSION_HISTORY]
    assert len(versions) == len(set(versions)), "Duplicate version in VERSION_HISTORY"


def test_resolve_workbook_uses_explicit_path_when_not_default(tmp_path):
    """A non-default --workbook always wins, even if a versioned file exists."""
    user_path = tmp_path / "my_custom.xlsm"
    assert _resolve_workbook(str(user_path)) == user_path


def test_resolve_workbook_finds_highest_version(tmp_path, monkeypatch):
    """When the caller passes our DEFAULT_WORKBOOK_PATH, pick the existing
    stocks_picker_vNN.xlsm with the highest NN."""
    from stocks_report import DEFAULT_WORKBOOK_PATH
    monkeypatch.chdir(tmp_path)
    (tmp_path / "stocks_picker_v01.xlsm").touch()
    (tmp_path / "stocks_picker_v03.xlsm").touch()
    (tmp_path / "stocks_picker_v02.xlsm").touch()
    out = _resolve_workbook(str(DEFAULT_WORKBOOK_PATH))
    assert out.name == "stocks_picker_v03.xlsm"


def test_resolve_workbook_falls_back_to_legacy_xlsm(tmp_path, monkeypatch):
    from stocks_report import DEFAULT_WORKBOOK_PATH
    monkeypatch.chdir(tmp_path)
    (tmp_path / "stocks.xlsm").touch()
    out = _resolve_workbook(str(DEFAULT_WORKBOOK_PATH))
    assert out.name == "stocks.xlsm"


# ---------------------------------------------------------------------------
# Help-sheet population
# ---------------------------------------------------------------------------

def test_ensure_help_sheet_versions_appends_missing(tmp_path):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Help"
    _ensure_help_sheet_versions(ws, fresh=True)
    # Every entry in VERSION_HISTORY should appear in column A
    col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    for ver, _date, _msg in VERSION_HISTORY:
        assert ver in col_a, f"{ver} missing from Help sheet"


def test_ensure_help_sheet_versions_preserves_user_credit_line(tmp_path):
    """User has a credit line at A1; our entries get appended below."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Developed by Claude"
    _ensure_help_sheet_versions(ws, fresh=False)
    assert ws["A1"].value == "Developed by Claude", "user credit clobbered"
    col_a_below = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    for ver, _date, _msg in VERSION_HISTORY:
        assert ver in col_a_below


def test_ensure_help_sheet_versions_does_not_duplicate(tmp_path):
    """Running twice does not re-append already-mentioned versions."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    _ensure_help_sheet_versions(ws, fresh=True)
    rows_after_first = ws.max_row
    _ensure_help_sheet_versions(ws, fresh=False)
    assert ws.max_row == rows_after_first, "second call appended duplicate entries"


# ---------------------------------------------------------------------------
# Currencies sheet helpers
# ---------------------------------------------------------------------------

def test_chf_is_in_fx_pairs():
    assert "CHF" in FX_PAIRS
    assert FX_PAIRS["CHF"] == "EURCHF=X"


def test_currency_headers_match_lookbacks():
    """Pair + every LOOKBACKS entry."""
    from stocks_report import LOOKBACKS
    assert CURRENCY_HEADERS == ["Pair"] + list(LOOKBACKS.keys())


def test_currency_rows_one_per_pair(tmp_path):
    """_currency_rows produces one row per FX_PAIRS entry, each starting with
    "EUR/{ccy}" and having a value per LOOKBACKS column."""
    from stocks_report import LOOKBACKS
    fake_history = {ccy: pd.Series(dtype=float) for ccy in FX_PAIRS}
    rows = _currency_rows(fake_history)
    assert len(rows) == len(FX_PAIRS)
    for r in rows:
        assert len(r) == 1 + len(LOOKBACKS)
        assert r[0].startswith("EUR/")


# ---------------------------------------------------------------------------
# Monthly movers
# ---------------------------------------------------------------------------

def test_compute_monthly_winners_filters_and_sorts():
    """Filter: 1D >= 0 AND 1W >= 0 AND 1M > 0. Sort: 1M desc."""
    records = [
        {"symbol": "A", "pct_1d": 0.01,  "pct_1w": 0.02,  "pct_1m": 0.10},  # qualifies
        {"symbol": "B", "pct_1d": 0.01,  "pct_1w": 0.02,  "pct_1m": 0.05},  # qualifies
        {"symbol": "C", "pct_1d": -0.01, "pct_1w": 0.05,  "pct_1m": 0.20},  # 1D negative — skip
        {"symbol": "D", "pct_1d": 0.01,  "pct_1w": -0.01, "pct_1m": 0.30},  # 1W negative — skip
        {"symbol": "E", "pct_1d": 0.01,  "pct_1w": 0.01,  "pct_1m": -0.05}, # 1M negative — skip
        {"symbol": "F", "pct_1d": None,  "pct_1w": 0.01,  "pct_1m": 0.10},  # missing — skip
        {"symbol": "G", "pct_1d": 0.0,   "pct_1w": 0.0,   "pct_1m": 0.07},  # zero 1D/1W is ok (>=0)
    ]
    out = _compute_monthly_winners(records, top_n=10)
    # A (10%), G (7%), B (5%) — order by 1M desc.
    assert [r["symbol"] for r in out] == ["A", "G", "B"]


def test_compute_monthly_winners_respects_top_n():
    records = [
        {"symbol": f"X{i}", "pct_1d": 0.01, "pct_1w": 0.01, "pct_1m": i / 100.0}
        for i in range(1, 11)
    ]
    out = _compute_monthly_winners(records, top_n=3)
    assert [r["symbol"] for r in out] == ["X10", "X9", "X8"]


def test_compute_monthly_losers_filters_and_sorts():
    """Filter: 1M < 0 AND 1D <= 0 AND 1W <= 0. Sort: 1M asc (worst first)."""
    records = [
        {"symbol": "A", "pct_1d": -0.01, "pct_1w": -0.02, "pct_1m": -0.10},  # qualifies
        {"symbol": "B", "pct_1d": -0.01, "pct_1w": -0.02, "pct_1m": -0.05},  # qualifies, less bad
        {"symbol": "C", "pct_1d": 0.01,  "pct_1w": -0.05, "pct_1m": -0.20},  # 1D positive — skip (bounce)
        {"symbol": "D", "pct_1d": -0.01, "pct_1w": 0.01,  "pct_1m": -0.30},  # 1W positive — skip
        {"symbol": "E", "pct_1d": -0.01, "pct_1w": -0.01, "pct_1m": 0.05},   # 1M positive — skip
        {"symbol": "F", "pct_1d": None,  "pct_1w": -0.01, "pct_1m": -0.10},  # missing — skip
        {"symbol": "G", "pct_1d": 0.0,   "pct_1w": 0.0,   "pct_1m": -0.07},  # zero 1D/1W ok (<=0)
    ]
    out = _compute_monthly_losers(records, top_n=10)
    # A (-10%), G (-7%), B (-5%) — ascending = most negative first.
    assert [r["symbol"] for r in out] == ["A", "G", "B"]


def test_ranking_headers_includes_owned_and_required_columns():
    for col in ("Symbol", "Owned?", "Today (EUR)", "1D %", "1W %", "1M %"):
        assert col in RANKING_HEADERS


def test_migrate_monthly_movers_renames_legacy_sheet(tmp_path):
    """Workbook with old 'Monthly movers' sheet gets it renamed in place,
    its content preserved."""
    from openpyxl import Workbook, load_workbook
    path = tmp_path / "legacy.xlsx"
    wb = Workbook()
    legacy = wb.active
    legacy.title = MONTHLY_MOVERS_LEGACY_NAME
    legacy["A1"] = "had data"
    wb.save(path)
    wb = load_workbook(path)
    _migrate_monthly_movers_to_winners_openpyxl(wb)
    wb.save(path)
    wb2 = load_workbook(path)
    assert MONTHLY_MOVERS_LEGACY_NAME not in wb2.sheetnames
    assert MONTHLY_WINNERS_SHEET_NAME in wb2.sheetnames
    assert wb2[MONTHLY_WINNERS_SHEET_NAME]["A1"].value == "had data"


def test_migrate_monthly_movers_no_op_when_winners_already_exists(tmp_path):
    """If both legacy and new names exist, migration must NOT clobber the
    new sheet — leave the legacy one alone for the user to delete."""
    from openpyxl import Workbook
    wb = Workbook()
    wb.active.title = MONTHLY_MOVERS_LEGACY_NAME
    wb.create_sheet(MONTHLY_WINNERS_SHEET_NAME)
    _migrate_monthly_movers_to_winners_openpyxl(wb)
    assert MONTHLY_MOVERS_LEGACY_NAME in wb.sheetnames
    assert MONTHLY_WINNERS_SHEET_NAME in wb.sheetnames


def test_init_workbook_portfolio_header_has_no_quantity(tmp_path):
    """Portfolio header should be just Symbol / Notes (Quantity removed)."""
    from openpyxl import load_workbook
    from stocks_report import PORTFOLIO_HEADER_ROW
    path = tmp_path / "stocks.xlsx"
    init_workbook(path)
    wb = load_workbook(path)
    main = wb["Main"]
    assert main.cell(row=PORTFOLIO_HEADER_ROW, column=1).value == "Symbol"
    assert main.cell(row=PORTFOLIO_HEADER_ROW, column=2).value == "Notes"
    # Column 3 should now be blank (Quantity was removed)
    assert main.cell(row=PORTFOLIO_HEADER_ROW, column=3).value is None


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


# ---------------------------------------------------------------------------
# _pick_test_target — get-quotes test-mode selection
# ---------------------------------------------------------------------------

def _market(rows):
    """rows: list of (row_idx, symbol, currency, indexes_csv)"""
    return rows


def test_pick_test_target_prefers_portfolio_symbol_in_market():
    rows = _market([(2, "ABI.BR", "EUR", "BEL20"), (3, "AAPL", "USD", "SP500")])
    assert _pick_test_target(rows, {"AAPL"})[1] == "AAPL"


def test_pick_test_target_handles_suffix_stripped_portfolio_entry():
    """User wrote 'KBC' in Main; Market has 'KBC.BR' — should still match."""
    rows = _market([(2, "ABI.BR", "EUR", "BEL20"), (3, "KBC.BR", "EUR", "BEL20")])
    assert _pick_test_target(rows, {"KBC"})[1] == "KBC.BR"


def test_pick_test_target_falls_back_to_first_bel20():
    rows = _market([(2, "AAPL", "USD", "SP500"), (3, "ABI.BR", "EUR", "BEL20")])
    assert _pick_test_target(rows, set())[1] == "ABI.BR"


def test_pick_test_target_falls_back_to_first_row_if_no_bel20():
    rows = _market([(2, "AAPL", "USD", "SP500"), (3, "GOOGL", "USD", "SP500")])
    assert _pick_test_target(rows, set())[1] == "AAPL"


def test_pick_test_target_returns_none_when_market_empty():
    assert _pick_test_target([], set()) is None


# ---------------------------------------------------------------------------
# read_test_mode — Main!B5 cell interpretation
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("raw,expected", [
    ("TRUE",  True),
    ("True",  True),
    ("true",  True),
    ("Yes",   True),
    ("Y",     True),
    ("1",     True),
    (1,       True),
    (True,    True),
    ("FALSE", False),
    ("No",    False),
    ("0",     False),
    (0,       False),
    (False,   False),
    ("",      False),
    (None,    False),
    ("  TRUE  ", True),  # whitespace tolerated
])
def test_read_test_mode_interpretation(raw, expected):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Main"
    ws[MAIN_CELLS["TestMode"]] = raw
    assert read_test_mode(ws) is expected


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
