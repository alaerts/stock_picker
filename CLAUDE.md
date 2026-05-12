# Stock Picker — Project Context for Claude

## What this is

A personal stock-tracking workbook driven by Python. One persistent
`stocks.xlsm` covering ~975 deduped stocks across 7 indexes (SP500, NIKKEI225,
FTSE100, DAX, CAC40, BEL20, ESTOXX50). The user opens Excel, clicks a button,
sees live progress in a Status cell, and ends up with refreshed data — all in
EUR at today's FX rate, plus a manually-maintained portfolio sheet that drives
an Owned? column on the data sheet.

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

One workbook, three sheets:

- **Main**: manual portfolio + job controls + Status cell + metadata block
  (last-run timestamps, FX rates, row count). TestMode lives at `Main!B5`.
- **Market**: 19-column inventory of the universe (see `MARKET_COLUMNS`).
- **xlwings.conf** (very-hidden): interpreter path + PYTHONPATH for xlwings.

## Decisions already made — do not re-litigate without asking

These were explicitly chosen by the user in the planning conversations
(2026-05-06 and 2026-05-12). If you think a different choice is better, raise
it as a question rather than changing the implementation.

- **Indexes**: SP500, NIKKEI225, FTSE100, DAX, CAC40, BEL20, ESTOXX50.
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
- **Test mode**: a single cell (`Main!B5`, named `TestMode` in
  `MAIN_CELLS`) gates whether each job processes the full universe or just
  BEL20 + 1 quote. CLI `--test` flag is an explicit override; without it,
  both jobs read the cell. Test mode reduces a 10-min full rebuild to ~40s.
- **Buttons**: Excel form controls assigned to VBA macros that call
  `RunPython` via the xlwings add-in. Buttons are added by
  `setup-buttons`; VBA module is imported once by the user via Alt+F11
  (programmatic VBA injection needs Excel's "Trust access to VBA project
  object model" setting, which we can't toggle).
- **Live progress feedback**: xlwings writes ticker counts to
  `Main!Status` (`Main!B6`) directly while the job runs.

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
2. Config: `INDEX_WIKI`, `WATCHLISTS`, `DATAROMA_ACTIVIST_CODES`,
   `FX_PAIRS`, `LOOKBACKS`, `INDEX_DEFAULT_CCY`,
   `DEFAULT_WORKBOOK_PATH`.
3. `normalize_ticker()` / `_parse_bel20_ticker()` — Wikipedia → Yahoo
   per index, including BEL20's "Euronext Brussels:\xa0SYMBOL" parsing.
4. `get_index_constituents()` / `aggregate_constituents()` — Wikipedia
   table fetcher, then per-Symbol aggregation with comma-joined
   `Indexes`.
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
    constants. `_layout_main_sheet()`, `_layout_market_sheet()`,
    `init_workbook()`.
13. `read_test_mode()`, `read_portfolio_symbols()`, `_owned_for()`,
    `_market_col()`, `_load_xl()`, `_resolve_workbook()` — workbook
    helpers.
14. `button_rebuild_inventory()` / `button_get_quotes()` — xlwings
    entry points; use the live workbook for both reads and writes.
15. `rebuild_inventory()` (Job 1, openpyxl, CLI) and
    `get_quotes()` (Job 2, openpyxl, CLI).
16. CLI subcommands: `init-workbook`, `setup-buttons`, `rebuild-inventory`,
    `get-quotes`, `run` (legacy).

Market sheet columns (see `MARKET_COLUMNS`):
`Symbol | Name | Owned? | Indexes | Sector | Watchlists | Currency |
 Today (EUR) | 1D ago (EUR) | 1W ago (EUR) | 1M ago (EUR) | 6M ago (EUR) |
 1Y ago (EUR) | 5Y ago (EUR) | P/E (TTM) | Forward P/E | Description |
 Last update (UTC) | Last error`

## Non-obvious things to know

- **Ticker symbol mismatches**: Wikipedia uses BRK.B; Yahoo uses BRK-B. `normalize_ticker()` handles SP500 dot→hyphen, FTSE trailing-dot stripping, and adds suffixes (`.T`, `.L`, `.DE`, `.PA`, `.BR`). ESTOXX50 ticker cells already arrive with Yahoo-correct suffixes (`.DE/.PA/.AS/.MI/.MC/.HE/.IR`) and are passed through.
- **GBp pence**: many UK stocks report in pence on yfinance with `currency = "GBp"`. We divide by 100 before converting via the GBP rate. See `to_eur()`.
- **Watchlist sources are fragile**: Yahoo's screener API is undocumented and rate-gates aggressively without cookie+crumb; the Berkshire scrIds is not in the public catalog and could disappear in any Yahoo redesign. Dataroma is plain HTML but its URL structure could change. Failures log warnings and the report still completes with empty Watchlists for the affected entries.
- **Dataroma uses BRK.B / BF.B; Yahoo uses BRK-B / BF-B**: `_extract_dataroma_tickers()` emits both forms so SP500 tickers match regardless of which form the index constituents table contains.
- **Watchlists are US-only**: hedge fund / Berkshire lists contain US tickers. Non-US stocks (BEL/CAC/DAX/FTSE/Nikkei/ESTOXX50) will always have empty `Watchlists` column. This is expected, not a bug.
- **Watchlist matching uses both the full Yahoo symbol AND the suffix-stripped root** — defensive in case a watchlist surfaces a name differently.
- **One row per Symbol with multi-index membership**: `aggregate_constituents()` collapses Wikipedia per-index constituent DataFrames into a single row per Symbol whose `Indexes` column lists every index that owns it (e.g. SAP.DE → "DAX, ESTOXX50"). The CLI legacy `run` mode and the new Market sheet both use this shape.
- **Workbook lifecycle**: `init-workbook` creates a vanilla .xlsx via openpyxl (it can't author a valid macro-enabled .xlsm from scratch — Excel rejects the content type). `setup-buttons` opens that .xlsx via Excel COM, adds button shapes + an `xlwings.conf` very-hidden sheet, and SaveAs's to .xlsm. `_resolve_workbook()` auto-routes the daily commands from "stocks.xlsx" to "stocks.xlsm" when the latter exists. `_load_xl()` opens with `keep_vba=True` for .xlsm so user-imported VBA survives every job.
- **Status cell live updates require xlwings**: openpyxl can't write to a workbook that Excel currently has open. The `button_*` entry points use the xlwings cell API for both reads (portfolio, TestMode) and writes (Status, Market data), so Excel reflects updates instantly. CLI `rebuild_inventory` / `get_quotes` use openpyxl and only log progress to stdout.
- **`.info` rate limiting**: ~975 calls × 0.25s ≈ 4 min. If Yahoo throttles, the script doesn't crash — those stocks just get blank P/E. Consider caching before scaling up frequency. Both jobs make `.info` calls (rebuild for sector/description/currency/name, get_quotes for trailing+forward P/E).
- **`yf.download()` with multiple tickers** returns a multi-index DataFrame. `fetch_close_prices()` extracts just the `Close` level. Watch for breaking changes across yfinance major versions.
- **5-year blanks for recent IPOs are correct, not a bug.**
- **The script is tolerant of partial failures**: a missing watchlist, a failing ticker, or a timed-out chunk all log warnings (or set the per-row Last error column) and the run still completes with the rest.
- **VBA Trust setting on user's machine**: programmatic VBA injection via `wb.api.VBProject.VBComponents.Add(...)` requires Excel's "Trust access to the VBA project object model" setting (off by default). That's why `setup-buttons` doesn't inject VBA; the user imports `vba/stocks_picker.bas` once manually.

## Testing

Offline unit tests (no network):
```powershell
.\.venv\Scripts\python.exe -m pytest test_stocks_report.py -v
```

78 unit tests cover the parsing/aggregation/lookup logic. All should pass
before any commit (per the user's commit-cadence rule in
`~/.claude/projects/.../memory/reference_github_repo.md`).

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
  single persistent `stocks.xlsm`. Substantial deviations require the
  user's say-so.

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

