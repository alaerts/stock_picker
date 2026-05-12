# Stock Picker

A personal stock-tracking workbook driven by Python. Maintains a single Excel
file with your portfolio, your "investable universe" across 7 major indexes,
and daily price + P/E refreshes — all in EUR using today's FX rate.

## What it tracks

**7 indexes (~975 stocks, deduped):** S&P 500, Nikkei 225, FTSE 100, DAX,
CAC 40, BEL 20, EuroStoxx 50.

**6 watchlists** (US tickers only):
- Recent 52-Week Highs (Yahoo screener)
- Berkshire Hathaway Portfolio (Yahoo screener)
- Top Quarterly Buys / Sells across "super investors" (dataroma)
- Most-Held by Super Investors (dataroma)
- Activist Hedge Fund Positions — aggregate across 7 activist managers
  (Icahn, Ackman, ValueAct, Trian, Engaged, Third Point, TCI) (dataroma)

**Per stock:** Symbol, Name, Owned? (from your portfolio), Indexes (multi),
Sector, Watchlists (multi), Currency, prices in EUR for 7 lookbacks
(Today / 1D / 1W / 1M / 6M / 1Y / 5Y), trailing P/E, forward P/E, business
description, last-update timestamp, last-error message.

## Architecture

One persistent workbook (`stocks.xlsm`) with three sheets:

- **Main** — your manual portfolio (Symbol / Quantity / Notes), plus the job
  controls and metadata (last run timestamps, FX rates, total rows). You
  edit this directly in Excel. Test-mode checkbox sits here.
- **Market** — the stock universe with all data. Rebuilt by Job 1; quotes
  refreshed by Job 2. Don't edit by hand.
- **xlwings.conf** — hidden config sheet so xlwings can find Python.

Two jobs:

- **rebuild_inventory** (Job 1, run weekly-ish, ~5–10 min): refreshes the
  structural columns — index memberships, watchlist memberships, sector,
  description, currency, Owned? lookup against your portfolio. Leaves quote
  columns blank.
- **get_quotes** (Job 2, run daily, ~5 min): refreshes prices + P/E for
  every row. Updates per-row Last update / Last error.

Both jobs trigger from Excel buttons on the Main sheet, with live progress
in a Status cell while they run. Or from the command line.

## First-time setup

```powershell
# 1. Install dependencies into a venv.
uv venv .venv
uv pip install -r requirements.txt --python .venv\Scripts\python.exe

# 2. Create the workbook (stocks.xlsx, plain).
.\.venv\Scripts\python.exe stocks_report.py init-workbook

# 3. One-time: add Excel buttons + xlwings.conf sheet and convert to .xlsm.
.\.venv\Scripts\python.exe stocks_report.py setup-buttons

# 4. Install the xlwings Excel add-in (once per machine).
.\.venv\Scripts\xlwings.exe addin install

# 5. In Excel: open stocks.xlsm, press Alt+F11, File -> Import File ->
#    select vba\stocks_picker.bas, save. Buttons now work.

# 6. Still in the VBA editor: Tools -> References -> tick "xlwings" -> OK.
#    Without this the macros compile-error with "Sub or Function not
#    defined" on RunPython (the symbol lives in the xlwings add-in and
#    each workbook's VBA project needs its own reference).
```

Optional polish (one-time):
- Right-click each button → **Edit Text** to confirm labels read "Rebuild
  Inventory" and "Get Quotes".
- Use the Test-mode checkbox on Main to toggle between BEL20-only smoke
  runs (~1 min) and the full universe (~10 min). Or use the `--test` CLI
  flag on either subcommand.

## Troubleshooting

**`'python' is not recognized`** — there's no plain `python` on PATH.
From cmd.exe: `.venv\Scripts\python.exe stocks_report.py ...` (or run
`.venv\Scripts\activate.bat` first, then `python` works for the session).
From PowerShell: `.\.venv\Scripts\python.exe stocks_report.py ...`.

**`Sub or Function not defined` on `RunPython` when clicking a button** —
the workbook's VBA project is missing its reference to the xlwings add-in.
Open the VBA editor (Alt+F11), **Tools → References**, tick **xlwings**.
If "xlwings" doesn't appear there, the add-in itself isn't loaded:
**File → Options → Add-ins → Manage: Excel Add-ins → Go** and tick
**xlwings**. Restart Excel afterwards if needed. This reference is
per-workbook; recreating `stocks.xlsm` from scratch means redoing
this one-time step.

**Truncated error popup** — shouldn't happen any more. Button-triggered
failures land in three places: `stocks_errors.log` next to the workbook
(full traceback, append-only), the `Errors` sheet inside the workbook
(one row per failure), and the `Main!Status` cell (concise summary).
If you're still seeing a truncated VBA MsgBox, the `button_*` functions
in `stocks_report.py` lost their `try/except` wrapper — check the source.

**`xlwings addin install` says "No module named xlwings.__main__"** —
xlwings doesn't ship a `__main__`. Use the installed console script:
`.\.venv\Scripts\xlwings.exe addin install` (not `python -m xlwings`).

## Day-to-day usage

**From Excel** (recommended): open `stocks.xlsm`, click a button. Watch the
Status cell on Main for live progress. Click your way through everything.

**From the command line:**

```powershell
# Daily quote refresh
.\.venv\Scripts\python.exe stocks_report.py get-quotes

# Weekly-ish full rebuild
.\.venv\Scripts\python.exe stocks_report.py rebuild-inventory

# Test mode: BEL 20 only + one quote (~40 sec end-to-end)
.\.venv\Scripts\python.exe stocks_report.py rebuild-inventory --test
.\.venv\Scripts\python.exe stocks_report.py get-quotes --test
```

When the CLI is invoked without `--test`, both commands read the workbook's
`Main!TestMode` cell — so you can toggle the checkbox in Excel and the CLI
will honor it.

## Tests

```powershell
.\.venv\Scripts\python.exe -m pytest test_stocks_report.py -v
```

78 unit tests, all offline (no network calls).

## Sources

- **yfinance** — prices, P/E, sector, description, currency.
- **Wikipedia** — index constituents for each of the 7 indexes.
- **Yahoo's predefined-screener API** (`/v1/finance/screener/predefined/saved`)
  for Recent 52-Week Highs and Berkshire's holdings, with the cookie+crumb
  auth dance.
- **dataroma.com** plain HTML for the 4 "super investor" 13F-derived lists.

All free, no accounts or API keys required.

## Caveats

- yfinance is unofficial; Yahoo can break it. `uv pip install --upgrade
  yfinance` if a previously-working run fails.
- Watchlists are US-only — non-US stocks (BEL/CAC/DAX/FTSE/Nikkei/ESTOXX50)
  always have an empty Watchlists column. Expected.
- 5-year lookback is blank for recently IPO'd companies. Expected.
- Historical prices use TODAY's FX rate, not the rate on that historical
  date — so non-EUR stock movements blend price movement with FX movement.
  Deliberate choice (see `CLAUDE.md`).
- xlwings buttons are Windows + Mac only. Linux / Pi users can still run
  the CLI commands; just no buttons.

## Project history

The original implementation was a single-shot script that wrote
`stocks_report_YYYY-MM-DD.xlsx` each run. That mode is still available as
`python stocks_report.py run --indexes ...` for backwards compatibility.
The "two jobs in one persistent workbook" architecture replaced it in
May 2026 to keep a portfolio sheet alongside the daily refresh.

See `MEMORY.md` for the decision log.
