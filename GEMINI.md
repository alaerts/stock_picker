# Stock Picker — Project Context for Claude

## What this is

A personal stock-tracking workbook driven by Python. One persistent
`stocks.xlsm` covering **~1000 tickers** — ~975 deduped stocks across 7
indexes (SP500, NIKKEI225, FTSE100, DAX, CAC40, BEL20, ESTOXX50) plus 38
curated ETFs (one per index, eleven sector SPDRs, twenty single-country
iShares MSCI funds). The user opens Excel, clicks a button, sees live
progress in a Status cell, and ends up with refreshed data — all in EUR
at today's FX rate, plus a manually-maintained portfolio sheet that drives
an Owned? column on the data sheet. Every Symbol cell is hyperlinked
to its Yahoo Finance quote page.

User is in Brussels (hence EUR base, BEL20 inclusion). Has a Raspberry Pi
running Home Assistant — possible target for headless CLI scheduling. Technical
user, Windows + Excel locally.

## Architecture at a glance

Two CLI jobs (also wired to Excel buttons via xlwings):

- **rebuild_inventory** (run weekly-ish): refreshes structural columns —
  Symbol/Name/Owned?/Indexes/Sector/Watchlists/Currency/Description. Reads
  Main!Portfolio for the Owned? lookup. Overwrites Market sheet.
- **get_quotes** (run daily): refreshes prices + P/E for every Market row.
  Sets per-row Last update / Last error. Surgical update — preserves
  structural columns.

One workbook (filename `stocks_picker_{SCHEMA_VERSION}.xlsm`), seven
sheets (Errors created on demand):

- **Main**: manual portfolio (Symbol + Notes — no Quantity), job controls
  (two buttons + a real Excel checkbox for test mode), Status cell, and
  the metadata block (last-run timestamps, FX rates, row count). Controls
  live at left=720 to clear column B; portfolio data starts at row 16.
  TestMode is the linked cell behind the checkbox (`Main!B5`).
- **Help**: append-only changelog. `init-workbook` populates one row per
  entry in `VERSION_HISTORY`. Row 1 is reserved for user-authored content
  (e.g. their own credit line); versioned entries get appended below.
- **Market**: 25-column inventory of the universe (see `MARKET_COLUMNS`).
  Every Symbol cell carries a Yahoo hyperlink. The user enables AutoFilter
  on this sheet; both rebuild paths re-apply the filter to the new data
  extent on every rebuild.
- **Monthly movers**: top 50 stocks ranked by 1M % descending, filtered
  to require 1D % and 1W % also positive (no short-term weakness on a
  monthly winner). Populated by `get_quotes` after the per-row loop.
- **Currencies**: one row per FX pair (USD / JPY / GBP / CHF) with rates
  at every LOOKBACKS offset (Today, 1D, 1W, 1M, 6M, 1Y, 5Y). Populated
  by `get_quotes` after `get_fx_history()` fetches 5y series per pair.
- **xlwings.conf** (very-hidden): interpreter path + PYTHONPATH for xlwings.
- **Errors** (created on first failure): one row per button-triggered
  exception with timestamp, job name, exception type, message, and the
  full traceback.

## Decisions already made — do not re-litigate without asking

These were explicitly chosen by the user in the planning conversations
(2026-05-06 and 2026-05-12). If you think a different choice is better, raise
it as a question rather than changing the implementation.

- **Indexes**: SP500, NIKKEI225, FTSE100, DAX, CAC40, BEL20, ESTOXX50.
- **ETFs**: curated list of 38 (one per index, eleven sector SPDRs,
  twenty single-country iShares MSCI funds). See `ETF_LIST` in
  `stocks_report.py`. Flow through `aggregate_constituents()` alongside
  index constituents and get the same .info / price / hyperlink
  treatment as stocks. ETFs are included in test-mode runs too.
- **Currencies tracked**: EUR/USD, EUR/JPY, EUR/GBP, EUR/CHF — used both
  to convert prices into EUR and to populate the Currencies sheet. CHF
  added 2026-05-13 at user request.
- **Schema versioning**: the workbook filename embeds `SCHEMA_VERSION`
  (e.g. `stocks_picker_v02.xlsm`). Bump it on major changes that add a
  sheet, a column, or change the data contract. `VERSION_HISTORY` is
  append-only and surfaces in the Help sheet.
- **Cross-index ticker normalization**: a stock listed in multiple
  indexes is one row in Market via `aggregate_constituents`. Any ticker
  already ending in a `KNOWN_EXCHANGE_SUFFIXES` value passes through
  `normalize_ticker` untouched — prevents double-suffixing when the same
  stock appears in CAC40 (uses bare "MT") and ESTOXX50 (uses "MT.AS").
  FTSE100 sub-class shares like BT.A become BT-A.L on Yahoo (hyphen,
  like SP500's BRK-B).
- **FX**: today's rate, applied to ALL historical prices. User chose this
  over historical rates. Trade-off accepted: mixes price movement with FX
  movement on non-EUR stocks.
- **Workbook**: single persistent `stocks.xlsm` (one row per Symbol with
  Indexes as comma-list, since stocks like SAP.DE belong to DAX + ESTOXX50).
  No daily dated `stocks_report_YYYY-MM-DD.xlsx` — that's the legacy `run`
  subcommand, kept for backwards compatibility only.
- **Watchlist column format**: ONE column, comma-separated list names. NOT
  N boolean columns.
- **Lookbacks**: Today, 1D, 1W, 1M, 6M, 1Y, 5Y ago.
- **Lookback semantics**: last available close *on or before* the target
  date — handles weekends/holidays by rolling back to the previous trading
  day.
- **Source**: yfinance for prices/PE/sector/description/currency; Wikipedia
  for index constituents; Yahoo's predefined-screener API
  (`/v1/finance/screener/predefined/saved`, with cookie+crumb auth) and
  dataroma.com (plain HTML "super investor" 13F aggregator) for watchlist
  memberships. All free, no accounts or API keys.
- **Test mode**: a real Excel form-control checkbox in `Main` with its
  LinkedCell set to `Main!B5` (the address tracked by `MAIN_CELLS["TestMode"]`).
  Excel writes TRUE/FALSE into that cell as the user toggles it.
  `read_test_mode()` accepts TRUE/FALSE/Yes/Y/1/bool. CLI `--test` flag is
  an explicit override; without it, both jobs read the cell. Test mode
  reduces a 10-min full rebuild to ~60s (BEL20 ~20 + ETFs 38 = ~58 rows).
- **Buttons**: Excel form controls (left=720, anchored upper-right to
  clear column A + B) assigned to VBA macros that call `RunPython` via
  the xlwings add-in. Buttons + checkbox are added by `setup-buttons`;
  VBA module is imported once by the user via Alt+F11 (programmatic VBA
  injection needs Excel's "Trust access to VBA project object model"
  setting, which we can't toggle).
- **Live progress feedback**: xlwings writes ticker counts to
  `Main!Status` (`Main!B6`) directly while the job runs.
- **Idempotent init-workbook**: re-running `init-workbook` on an existing
  file does NOT overwrite user cosmetic edits — label text, fonts, column
  widths the user has changed all survive. Only label cells with
  None-values get filled, and column widths are reset only on fresh
  creation. Same idea on `_layout_market_sheet`.
- **Market!Symbol hyperlinks**: every Symbol cell is hyperlinked to
  `https://finance.yahoo.com/quote/{symbol}`. Both rebuild paths attach
  the hyperlink; cell value stays plain text so sorting / filtering /
  `_market_col` lookups keep working.
- **Error reporting**: button-triggered failures land in THREE places
  (no truncated VBA MsgBox): `stocks_errors.log` next to the workbook
  (full traceback, append-only), an `Errors` sheet inside the workbook
  (one row per failure with timestamp, job, exception, message, tb),
  and a concise `Main!Status` cell summary with a hint pointing at both.
  The `_handle_button_exception()` helper does this; each write is
  wrapped in try/except so the handler itself never crashes.

## Watchlists tracked

Slim 6-list mapping adopted 2026-05-06 (Option B), replacing the original
8-list spec after Yahoo deprecated the legacy watchlist pages. See
`WATCHLISTS` in `stocks_report.py` for the canonical list.

| Label                                  | Source                | Reference |
|----------------------------------------|-----------------------|-----------|
| Recent 52-Week Highs                   | Yahoo screener API    | `recent_52_week_highs` |
| Berkshire Hathaway Portfolio           | Yahoo screener API    | `top_stocks_owned_by_warren_buffet` (undocumented but live) |
| Top Quarterly Buys (Super Investors)   | dataroma.com          | `/m/g/portfolio_b.php?q=q&o=c` |
| Top Quarterly Sells (Super Investors)  | dataroma.com          | `/m/g/portfolio_s.php?q=q&o=c` |
| Most-Held by Super Investors           | dataroma.com          | `/m/g/portfolio.php` (~100 entries) |
| Activist Hedge Fund Positions          | dataroma.com          | aggregate of `/m/holdings.php?m={code}` for codes `ic, psc, VA, TF, ENG, tp, tci` (Icahn, Ackman, ValueAct, Trian, Engaged, Third Point, TCI) |

Yahoo screener auth: visit `fc.yahoo.com` then any /research-hub/screener/ page
to seed the consent cookie, then GET `/v1/test/getcrumb` for a crumb token.
The screener API caps responses at 5 records per call, so paginate via `&start=N`.

Lists dropped from the original spec because Yahoo no longer serves them and
the closest dataroma equivalents would duplicate the lists above:
*Smart Money Stocks*, *Crowded Hedge Fund Positions* (subsumed by "Most-Held"),
*Most Bought by Hedge Funds* (subsumed by "Top Quarterly Buys"),
*Most Bought / Most Sold by Activist Hedge Funds* (subsumed by the activist
positions aggregate).

## Files

- `stocks_report.py` — single-file script. All Python logic.
- `vba/stocks_picker.bas` — VBA module the user imports once into the
  workbook. Each macro is a one-line `RunPython` into a `button_*`
  entry-point function.
- `requirements.txt` — yfinance, pandas, openpyxl, requests, lxml,
  beautifulsoup4, xlwings (win32 + darwin only).
- `test_stocks_report.py` — 78 offline unit tests.
- `README.md` — user-facing docs.
- `MEMORY.md` — decision log (read at the start of every session per
  the user's instruction).

## Code structure (in `stocks_report.py`)

1. Logging.
2. Config: `INDEX_WIKI`, `ETF_LIST`, `WATCHLISTS`,
   `DATAROMA_ACTIVIST_CODES`, `FX_PAIRS`, `LOOKBACKS`,
   `INDEX_DEFAULT_CCY`, `DEFAULT_WORKBOOK_PATH`, `YAHOO_QUOTE_URL`,
   `ERROR_LOG_FILENAME`, `ERRORS_SHEET_NAME`.
3. `normalize_ticker()` / `_parse_bel20_ticker()` /
   `_parse_nikkei225_components()` — Wikipedia → Yahoo per index,
   including BEL20's "Euronext Brussels:\xa0SYMBOL" parsing and the
   bullet-list scrape for Nikkei (which no longer publishes a table).
4. `get_index_constituents()` / `get_etfs()` / `aggregate_constituents()`
   — Wikipedia table fetcher + curated ETF DataFrame, then per-Symbol
   aggregation with comma-joined `Indexes`.
5. `fetch_yahoo_screener()`, `fetch_dataroma_tickers()`,
   `fetch_dataroma_activist_aggregate()`, `fetch_all_watchlists()` —
   Yahoo predefined-screener API (cookie+crumb) plus dataroma.com
   plain-HTML scraping.
6. `get_fx_rates()` — EURUSD=X, EURJPY=X, EURGBP=X.
7. `to_eur()` — incl. GBp pence handling.
8. `fetch_close_prices()` — chunked, throttled.
9. `fetch_ticker_info()` / `fetch_all_info()` — per-ticker, throttled,
   with optional progress callback (so the workbook Status cell updates
   live during rebuild_inventory / get_quotes).
10. `price_at_or_before()` — lookback date resolver.
11. `build_report()` + `write_excel()` — legacy single-shot mode
    (`stocks_report.py run`).
12. `MARKET_COLUMNS`, `MAIN_CELLS`, `PORTFOLIO_*` — workbook layout
    constants. `_layout_main_sheet(overwrite=True/False)`,
    `_layout_market_sheet(overwrite=True/False)`, `init_workbook()`.
13. `read_test_mode()`, `read_portfolio_symbols()`, `_owned_for()`,
    `_market_col()`, `_load_xl()`, `_resolve_workbook()`,
    `yahoo_quote_url()` — workbook helpers.
14. `_append_error_log()`, `_handle_button_exception()` — failure path
    that writes to log file + Errors sheet + Status cell.
15. `_xw_status()`, `_xw_read_portfolio()`, `_xw_read_test_mode()`,
    `button_rebuild_inventory()`, `button_get_quotes()` — xlwings entry
    points; use the live workbook for both reads and writes, attach
    Yahoo hyperlinks to Symbol cells via COM Hyperlinks.Add.
16. `rebuild_inventory()` (Job 1, openpyxl, CLI) and
    `get_quotes()` (Job 2, openpyxl, CLI).
17. CLI subcommands: `init-workbook`, `setup-buttons` (also adds the
    Test-mode checkbox), `rebuild-inventory`, `get-quotes`, `run`
    (legacy).

Market sheet columns (see `MARKET_COLUMNS`):
`Symbol | Name | Owned? | Indexes | Sector | Watchlists | Currency |
 Today (EUR) | 1D ago (EUR) | 1W ago (EUR) | 1M ago (EUR) | 6M ago (EUR) |
 1Y ago (EUR) | 5Y ago (EUR) | P/E (TTM) | Forward P/E | Description |
 Last update (UTC) | Last error`

Symbol cells carry a Yahoo Finance hyperlink (`finance.yahoo.com/quote/{sym}`).
ETF rows have `Indexes` values like `ETF — SP500` or `ETF — Sector:
Utilities`; otherwise their shape matches stock rows exactly.

## Non-obvious things to know

- **Ticker symbol mismatches**: Wikipedia uses BRK.B; Yahoo uses BRK-B. `normalize_ticker()` handles SP500 dot→hyphen, FTSE trailing-dot stripping, and adds suffixes (`.T`, `.L`, `.DE`, `.PA`, `.BR`). ESTOXX50 ticker cells already arrive with Yahoo-correct suffixes (`.DE/.PA/.AS/.MI/.MC/.HE/.IR`) and are passed through.
- **GBp pence**: many UK stocks report in pence on yfinance with `currency = "GBp"`. We divide by 100 before converting via the GBP rate. See `to_eur()`.
- **Watchlist sources are fragile**: Yahoo's screener API is undocumented and rate-gates aggressively without cookie+crumb; the Berkshire scrIds is not in the public catalog and could disappear in any Yahoo redesign. Dataroma is plain HTML but its URL structure could change. Failures log warnings and the report still completes with empty Watchlists for the affected entries.
- **Dataroma uses BRK.B / BF.B; Yahoo uses BRK-B / BF-B**: `_extract_dataroma_tickers()` emits both forms so SP500 tickers match regardless of which form the index constituents table contains.
- **Watchlists are US-only**: hedge fund / Berkshire lists contain US tickers. Non-US stocks (BEL/CAC/DAX/FTSE/Nikkei/ESTOXX50) will always have empty `Watchlists` column. This is expected, not a bug.
- **Watchlist matching uses both the full Yahoo symbol AND the suffix-stripped root** — defensive in case a watchlist surfaces a name differently.
- **One row per Symbol with multi-index membership**: `aggregate_constituents()` collapses Wikipedia per-index constituent DataFrames into a single row per Symbol whose `Indexes` column lists every index that owns it (e.g. SAP.DE → "DAX, ESTOXX50"). The CLI legacy `run` mode and the new Market sheet both use this shape.
- **Workbook lifecycle**: `init-workbook` creates a vanilla .xlsx via openpyxl (it can't author a valid macro-enabled .xlsm from scratch — Excel rejects the content type). `setup-buttons` opens that .xlsx via Excel COM, adds button shapes + a test-mode checkbox (LinkedCell=Main!B5) + an `xlwings.conf` very-hidden sheet, and SaveAs's to .xlsm. `_resolve_workbook()` auto-routes the daily commands from "stocks.xlsx" to "stocks.xlsm" when the latter exists. `_load_xl()` opens with `keep_vba=True` for .xlsm so user-imported VBA survives every job. On re-runs `init-workbook` skips writing label cells that already have a value and skips re-applying column widths — protecting user cosmetic edits.
- **Status cell live updates require xlwings**: openpyxl can't write to a workbook that Excel currently has open. The `button_*` entry points use the xlwings cell API for both reads (portfolio, TestMode) and writes (Status, Market data), so Excel reflects updates instantly. CLI `rebuild_inventory` / `get_quotes` use openpyxl and only log progress to stdout.
- **`.info` rate limiting**: ~1010 calls × 0.25s ≈ 4 min (975 stocks + 38 ETFs). If Yahoo throttles, the script doesn't crash — those stocks just get blank P/E. Consider caching before scaling up frequency. Both jobs make `.info` calls (rebuild for sector/description/currency/name, get_quotes for trailing+forward P/E).
- **`yf.download()` with multiple tickers** returns a multi-index DataFrame. `fetch_close_prices()` extracts just the `Close` level. Watch for breaking changes across yfinance major versions.
- **5-year blanks for recent IPOs are correct, not a bug.**
- **The script is tolerant of partial failures**: a missing watchlist, a failing ticker, or a timed-out chunk all log warnings (or set the per-row Last error column) and the run still completes with the rest.
- **VBA Trust setting on user's machine**: programmatic VBA injection via `wb.api.VBProject.VBComponents.Add(...)` requires Excel's "Trust access to the VBA project object model" setting (off by default). That's why `setup-buttons` doesn't inject VBA; the user imports `vba/stocks_picker.bas` once manually.

## Testing

Offline unit tests (no network):
```powershell
.\.venv\Scripts\python.exe -m pytest -v
```

84 unit tests cover the parsing/aggregation/lookup logic. All should pass
before any commit (per the user's commit-cadence rule in
`~/.claude/projects/.../memory/reference_github_repo.md`).

Live integration tests (hits Wikipedia / Yahoo / dataroma; ~30 sec):
```powershell
.\.venv\Scripts\python.exe -m pytest --integration -v
```

16 integration tests guard against the failure class where an upstream
source restructures its page and our parser silently breaks — the
"Could not locate constituent table for NIKKEI225" button error from
2026-05-12 is the canonical example. Skipped by default; run before any
release / scheduling change.

End-to-end smoke (test mode = BEL20 + 1 quote, ~40 sec):
```powershell
.\.venv\Scripts\python.exe stocks_report.py init-workbook
.\.venv\Scripts\python.exe stocks_report.py rebuild-inventory --test
.\.venv\Scripts\python.exe stocks_report.py get-quotes --test
```

Full smoke (BEL20 + SP500, ~7 minutes — proves watchlists populate):
```powershell
.\.venv\Scripts\python.exe stocks_report.py rebuild-inventory --indexes BEL20,SP500
.\.venv\Scripts\python.exe stocks_report.py get-quotes
```

Full run (~15 min):
```powershell
.\.venv\Scripts\python.exe stocks_report.py rebuild-inventory
.\.venv\Scripts\python.exe stocks_report.py get-quotes
```

For Excel-button-driven validation, run `setup-buttons` once and import the
VBA module per the README setup steps; clicking the button is then the
smoke test.

## Open items / next steps (priority order)

1. **First full run on the new architecture**: run `rebuild-inventory` then
   `get-quotes` against the live workbook. Watch for:
   - `.info` rate limiting at scale (~975 tickers × 0.25s ≈ 4 min each job).
   - Currency column blanks on the new ESTOXX50 multi-exchange tickers.
   - Watchlists populating for SP500 stocks (203/503 in the BEL20+SP500
     smoke; should be similar on the full run).
2. **First button-driven run**: import `vba/stocks_picker.bas` per README,
   click the buttons. Validate live ticker count in Main!Status.
3. **Scheduling**: still TBD. Could be Windows Task Scheduler for the local
   `.venv\Scripts\python.exe stocks_report.py get-quotes` daily, with the
   workbook closed during scheduled runs. Pi cron also possible (CLI works
   without xlwings/Excel; just lose buttons + live status).
4. **Optional retry/cache layer** for `.info` if rate-limiting becomes a
   problem. SQLite cache keyed by ticker, with TTL.
5. **Optional**: per-stock %-change columns, sparklines, historical FX
   flag, diff-vs-last-week. All discussed but deferred.

## Extension ideas — discuss before implementing

- Per-stock daily/weekly/monthly % change columns (computable in Excel
  formula or in pandas).
- Historical FX option (we deliberately chose today's rate; could be a
  `--historical-fx` flag).
- Additional indexes (FTSE 250, SMI, IBEX 35, AEX, EuroStoxx 600).
- Diff against last run (which tickers entered/exited an index or a
  watchlist).
- An "industry" column (yfinance has Sector + Industry; we only surface
  Sector).
- A "% change" sparkline column using openpyxl's image embedding.

## What NOT to change without asking

- FX semantics (today's rate). The user explicitly chose this.
- One row per Symbol with `Indexes` as a comma-joined list (NOT one row
  per Index-Symbol pair).
- Comma-separated single Watchlists column (NOT N boolean columns).
- Free / no-API-key requirement.
- Throttling and chunk delays — Yahoo will rate-limit aggressively if
  removed.
- The two-jobs architecture (`rebuild_inventory` + `get_quotes`) and the
  single persistent `stocks_picker_v{NN}.xlsm`. Substantial deviations
  require the user's say-so.
- Cosmetic preservation: `init-workbook` re-runs must NOT touch any cell
  on the Main sheet — neither rewrite default labels (even if cleared)
  nor restyle. Same rule applies to Help (append-only). Market and
  Currencies/Monthly-movers header text IS re-asserted (column-name
  contract) but bold + fill + widths are gated on fresh creation.
- The Portfolio is **Symbol + Notes** only (no Quantity column).
- The Test-mode control is a real Excel checkbox (added by
  `setup-buttons`) with LinkedCell=Main!B5.
- Error reporting is log file + Errors sheet + Status cell. Do NOT
  re-introduce the truncated VBA MsgBox by re-raising from `button_*`.
- Market AutoFilter must be re-applied to the new data extent on every
  rebuild — the user's sort/filter survives that way.
- New sheets (Help / Currencies / Monthly movers) are populated by the
  scripts but their styling/widths are set only on FIRST creation
  (cosmetic preservation rule applies).

## Conventions

- Logging via `logging` module, INFO level by default.
- Failures log warnings and continue rather than crash, unless they'd produce a blank report.
- All user-facing strings in English.



## Memory file
Maintain a file called MEMORY.md. After any significant decision, about direction, format, content, approach, or strategy, add an entry:

   ## [Date], [Decision]
   **What was decided:** [the choice made]
   **Why:** [the reasoning]
   **What was rejected:** [alternatives considered and why they were ruled out]

Read MEMORY.md at the start of every session before doing anything. Never contradict a logged decision without flagging it first.

## session end
When I say "session end", "wrapping up", or "let's stop here", write a session summary to MEMORY.md:

   ## Session Summary, [Date]
   **Worked on:** [what we focused on]
   **Completed:** [what's finished]
   **In progress:** [what's started but not done]
   **Decisions made:** [key choices from this session]
   **Next session:** [what to pick up first and any important context to carry forward]

## error file
Maintain a file called ERRORS.md. When an approach takes more than 2 attempts to work, log it:

   ## [Task type or description]
   **What didn't work:** [approaches that failed and why]
   **What worked:** [the approach that finally succeeded]
   **Note for next time:** [anything worth remembering for similar tasks]

Check ERRORS.md before suggesting approaches to tasks similar to logged ones. If a task matches a logged failure, say so and skip to what worked.

## Test-driven approach
Use a test-driven approach. Every features should have a test harness. 
All the tests should succeed upon completing a change

## Tools and permissions
Use WebFetch tool to scrape the web
Do not ask permission to make changes within the project folder
Save everything to https://github.com/alaerts/stock_picker
Commit and push after each major changes, and after every set of 5 small changes, with a summary of those changes. Make sure the tests run clean before comitting

