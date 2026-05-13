"""xlwings / Excel COM tests.

These exercise the workbook-manipulation code paths that the offline
test suite cannot reach: anything that calls `.api` on an xlwings
sheet/range to invoke Excel COM directly.

They caught the failure class where COM methods on hidden / non-active
sheets blow up — see the 2026-05-13 commits for:
  - "Worksheet.Move method failed" when adding a sheet next to the
    very-hidden xlwings.conf sheet.
  - "AutoFilter method of Range class failed" when re-applying filter
    on a non-active sheet.

Both bugs slipped past the offline suite (which uses openpyxl) and the
integration suite (which exercises network fetchers, not Excel).

Run with: pytest --integration -v test_xlwings.py
Requires: Excel installed + xlwings.
"""
import pandas as pd
import pytest

import stocks_report
from stocks_report import (
    CURRENCIES_SHEET_NAME,
    FX_PAIRS,
    MARKET_COLUMNS,
    MONTHLY_LOSERS_SHEET_NAME,
    MONTHLY_MOVERS_LEGACY_NAME,
    MONTHLY_WINNERS_SHEET_NAME,
    _ensure_currencies_sheet_xlwings,
    _ensure_ranking_sheet_xlwings,
    _migrate_monthly_movers_to_winners_xlwings,
    _xw_last_visible_sheet,
    _xw_reapply_market_autofilter,
    init_workbook,
)


pytestmark = pytest.mark.integration


@pytest.fixture(scope="module")
def xw_app():
    """A shared headless Excel app for the whole module — saves ~3s per test.

    **Safety:** if the user already has Excel running, we skip the entire
    module. Creating a new App while another exists has, on some Excel/COM
    version combos, caused `app.quit()` to also tear down the user's
    sessions — and the cost of a false negative (lost work) dwarfs the
    benefit of running these tests right now. Close Excel before running
    `pytest --integration test_xlwings.py` if you want this layer covered.
    """
    import xlwings as xw
    if list(xw.apps):
        pids = [a.pid for a in xw.apps]
        pytest.skip(
            f"Skipping xlwings tests — Excel is already running (PIDs {pids}). "
            "Close Excel and re-run if you want this layer covered."
        )
    app = xw.App(visible=False, add_book=False)
    yield app
    # Best-effort cleanup; close any leftover workbooks too.
    try:
        for book in list(app.books):
            try:
                book.close()
            except Exception:
                pass
    except Exception:
        pass
    app.quit()


@pytest.fixture
def workbook_with_hidden_trailing_sheet(tmp_path, xw_app):
    """A workbook shaped like the user's setup-buttons output: Main / Help /
    Market + a very-hidden trailing sheet (mimics xlwings.conf)."""
    path = tmp_path / "stocks_test.xlsx"
    init_workbook(path)
    wb = xw_app.books.open(str(path))
    # Inject a very-hidden trailing sheet to reproduce the user's state.
    trailing = wb.sheets.add("trailing_hidden", after=_xw_last_visible_sheet(wb))
    trailing.api.Visible = 2  # xlSheetVeryHidden
    yield wb
    try:
        wb.close()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# _xw_last_visible_sheet
# ---------------------------------------------------------------------------

def test_last_visible_sheet_skips_very_hidden(workbook_with_hidden_trailing_sheet):
    """Regression for 2026-05-13 bug 2: the very-hidden xlwings.conf
    sheet was being returned as the "last sheet" target for new sheet
    additions, which caused Worksheet.Move to fail.
    """
    last_visible = _xw_last_visible_sheet(workbook_with_hidden_trailing_sheet)
    assert last_visible.name != "trailing_hidden", (
        "Helper returned the hidden trailing sheet — bug regressed"
    )
    # Verify it's actually visible
    assert last_visible.api.Visible == -1


def test_last_visible_sheet_returns_sheet_object(xw_app, tmp_path):
    """Sanity: even on a workbook where every sheet is visible, helper
    returns the actually-last sheet (not the second-to-last)."""
    path = tmp_path / "all_visible.xlsx"
    init_workbook(path)
    wb = xw_app.books.open(str(path))
    try:
        last_visible = _xw_last_visible_sheet(wb)
        # init_workbook creates Main, Market, Help — last is Help.
        assert last_visible.name == "Help"
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Sheet creation paths that previously hit "Worksheet.Move failed"
# ---------------------------------------------------------------------------

def test_ensure_currencies_sheet_succeeds_with_hidden_trailing(workbook_with_hidden_trailing_sheet):
    """The bug: sheets.add(after=very_hidden_sheet) raises
    "Move method of Worksheet class failed". The fix routes the
    `after=` argument through _xw_last_visible_sheet."""
    wb = workbook_with_hidden_trailing_sheet
    fake_history = {ccy: pd.Series(dtype=float) for ccy in FX_PAIRS}
    # MUST NOT raise — this is the regression
    _ensure_currencies_sheet_xlwings(wb, fake_history)
    assert CURRENCIES_SHEET_NAME in [s.name for s in wb.sheets]


def test_ensure_currencies_sheet_is_idempotent(workbook_with_hidden_trailing_sheet):
    """Second call shouldn't duplicate the sheet or crash."""
    wb = workbook_with_hidden_trailing_sheet
    fake_history = {ccy: pd.Series(dtype=float) for ccy in FX_PAIRS}
    _ensure_currencies_sheet_xlwings(wb, fake_history)
    _ensure_currencies_sheet_xlwings(wb, fake_history)
    matches = [s.name for s in wb.sheets if s.name == CURRENCIES_SHEET_NAME]
    assert len(matches) == 1


def test_ensure_ranking_sheet_winners_succeeds_with_hidden_trailing(workbook_with_hidden_trailing_sheet):
    """Same bug class as Currencies — used to fail because sheets.add(after=
    very_hidden_sheet) raised Move-method-failed."""
    wb = workbook_with_hidden_trailing_sheet
    _ensure_ranking_sheet_xlwings(wb, MONTHLY_WINNERS_SHEET_NAME, rows=[])
    assert MONTHLY_WINNERS_SHEET_NAME in [s.name for s in wb.sheets]


def test_ensure_ranking_sheet_losers_succeeds_with_hidden_trailing(workbook_with_hidden_trailing_sheet):
    wb = workbook_with_hidden_trailing_sheet
    _ensure_ranking_sheet_xlwings(wb, MONTHLY_LOSERS_SHEET_NAME, rows=[], filter_owned_yes=True)
    assert MONTHLY_LOSERS_SHEET_NAME in [s.name for s in wb.sheets]


def test_ensure_ranking_sheet_writes_rows_with_owned_column(workbook_with_hidden_trailing_sheet):
    """Sheet gets the 9-col header including Owned? at col C."""
    wb = workbook_with_hidden_trailing_sheet
    rows = [
        {"symbol": "AAPL", "name": "Apple", "owned": "Yes", "indexes": "SP500", "sector": "Tech",
         "today": 200.0, "pct_1d": 0.01, "pct_1w": 0.03, "pct_1m": 0.10},
        {"symbol": "MSFT", "name": "Microsoft", "owned": "No", "indexes": "SP500", "sector": "Tech",
         "today": 400.0, "pct_1d": 0.005, "pct_1w": 0.02, "pct_1m": 0.08},
    ]
    _ensure_ranking_sheet_xlwings(wb, MONTHLY_WINNERS_SHEET_NAME, rows=rows)
    sheet = wb.sheets[MONTHLY_WINNERS_SHEET_NAME]
    assert sheet.range("A2").value == "AAPL"
    assert sheet.range("C2").value == "Yes"   # Owned? column
    assert sheet.range("C3").value == "No"


def test_migrate_monthly_movers_xlwings_renames(workbook_with_hidden_trailing_sheet):
    """The xlwings migration helper renames an existing 'Monthly movers'
    sheet without affecting its content."""
    wb = workbook_with_hidden_trailing_sheet
    legacy = wb.sheets.add(MONTHLY_MOVERS_LEGACY_NAME)
    legacy.range("A1").value = "preserve me"
    _migrate_monthly_movers_to_winners_xlwings(wb)
    names = [s.name for s in wb.sheets]
    assert MONTHLY_MOVERS_LEGACY_NAME not in names
    assert MONTHLY_WINNERS_SHEET_NAME in names
    assert wb.sheets[MONTHLY_WINNERS_SHEET_NAME].range("A1").value == "preserve me"


# ---------------------------------------------------------------------------
# AutoFilter re-apply helper
# ---------------------------------------------------------------------------

def test_reapply_market_autofilter_never_raises(workbook_with_hidden_trailing_sheet):
    """Regression for 2026-05-13 bug 1: Range.AutoFilter() blew up the
    whole rebuild. The helper must swallow ANY exception and return a bool
    so the rebuild path stays alive.

    Note: under headless Excel (visible=False), `Range.AutoFilter()` reliably
    fails with "AutoFilter method of Range class failed" — that's an Excel
    quirk, not a code bug. The user runs visible Excel via the button, where
    it normally succeeds. The test asserts the resilience contract regardless.
    """
    wb = workbook_with_hidden_trailing_sheet
    market = wb.sheets["Market"]
    rows = [["SYM_" + str(i)] + [None] * (len(MARKET_COLUMNS) - 1) for i in range(5)]
    market.range((2, 1)).value = rows

    last_col_letter = stocks_report.get_column_letter(len(MARKET_COLUMNS))
    # MUST NOT raise even when Excel rejects the COM call
    ok = _xw_reapply_market_autofilter(market, wb, last_col_letter, 1 + len(rows))
    assert isinstance(ok, bool)


def test_reapply_market_autofilter_restores_active_sheet(workbook_with_hidden_trailing_sheet):
    """The helper temporarily activates Market for the AutoFilter call.
    Whatever sheet the user had active before MUST be restored — otherwise
    every button click silently changes the user's view."""
    wb = workbook_with_hidden_trailing_sheet
    market = wb.sheets["Market"]
    rows = [["SYM_" + str(i)] + [None] * (len(MARKET_COLUMNS) - 1) for i in range(5)]
    market.range((2, 1)).value = rows
    # Make Main the active sheet to ensure it survives the helper.
    wb.sheets["Main"].api.Activate()
    prior_active = wb.api.ActiveSheet.Name
    assert prior_active == "Main"

    last_col_letter = stocks_report.get_column_letter(len(MARKET_COLUMNS))
    _xw_reapply_market_autofilter(market, wb, last_col_letter, 1 + len(rows))

    # Whether AutoFilter itself succeeded or not, the previously-active sheet
    # must be restored — the finally clause owns this contract.
    assert wb.api.ActiveSheet.Name == prior_active


def test_reapply_market_autofilter_leaves_clean_state_on_failure(workbook_with_hidden_trailing_sheet):
    """On failure, AutoFilterMode must be cleanly OFF (not in a half-broken
    state where the user sees ghost dropdowns spanning the wrong range)."""
    wb = workbook_with_hidden_trailing_sheet
    market = wb.sheets["Market"]
    # Pre-existing AutoFilter on a tiny range — we want to verify it ends up
    # cleanly cleared if the re-apply can't run for any reason.
    market.range((2, 1)).value = ["SYM"]  # one data row
    last_col_letter = stocks_report.get_column_letter(len(MARKET_COLUMNS))
    ok = _xw_reapply_market_autofilter(market, wb, last_col_letter, 2)
    if not ok:
        # Failure path: AutoFilterMode must be False
        assert market.api.AutoFilterMode is False
