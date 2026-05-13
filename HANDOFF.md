# Stock Picker — Session Handoff (2026-05-06 to 2026-05-13)

This document summarizes what we built together in this multi-day session, the decisions taken, the current state of the repo, and outstanding work. Read it first if you (or future-me) are picking the project back up.

## Repo

- **GitHub:** https://github.com/alaerts/stock_picker (public, branch `main`)
- **Local:** `c:\Dev\my\stock_picker\`
- **Active workbook:** `stocks_picker_v02.xlsm` (~761 KB, 983 Market rows, your portfolio)
- **Schema version (code):** `v03` (workbook still on v02 filename — see "Outstanding")

## Initial objectives

When the session opened, the project was a single-file Python script (`stocks_report.py`) that wrote a fresh dated Excel report each run. The original BEL 20 smoke test wouldn't even start. Goals over the session:

1. Get the existing single-shot script working end-to-end against current data sources.
2. Verify watchlist data is real (early discovery: it wasn't — see "Decisions").
3. Move the project to GitHub with a sensible commit cadence.
4. Re-architect into a "two-jobs + persistent workbook" model the user could drive from Excel buttons.
5. Add several new Excel-side features (Help, Currencies, Monthly winners, Monthly losers, ETFs, hyperlinks, AutoFilter, schema versioning).
6. Build a real test pyramid (offline / network-integration / Excel-COM) that catches the bug classes we were hitting in production.

## Decisions taken (durable)

| # | Decision | Notes |
|---|----------|-------|
| 1 | **6-list watchlist mapping (Option B)** | Yahoo's legacy `/u/yahoo-finance/watchlists/{slug}/` pages were dead. We kept Berkshire via the undocumented `top_stocks_owned_by_warren_buffet` scrIds + `recent_52_week_highs` from Yahoo screener, and switched the rest to dataroma.com (4 super-investor lists). |
| 2 | **Two-jobs architecture** | `rebuild_inventory` (slow, ~10 min, structural data) + `get_quotes` (daily, ~10 min, prices and P/E). |
| 3 | **One persistent workbook** | `stocks_picker_v{NN}.xlsm`. Replaces the prior "fresh dated xlsx per run". The user manually maintains the Portfolio sheet inside; data sheets are refreshed by the jobs. |
| 4 | **xlwings + VBA buttons** | Two Excel buttons (`RebuildInventory`, `GetQuotes`) + a real form-control checkbox for Test mode. VBA module ships in `vba/stocks_picker.bas` and is auto-imported by `setup-buttons` (requires "Trust access to VBA project object model"). |
| 5 | **One row per Symbol with `Indexes` list** | A stock in multiple indexes (e.g. SAP.DE in DAX + ESTOXX50) is ONE row with `Indexes = "DAX, ESTOXX50"`. |
| 6 | **`KNOWN_EXCHANGE_SUFFIXES` cross-index normalization** | A ticker arriving fully-qualified (e.g. `AIR.PA` in DAX's table) passes through `normalize_ticker` unchanged. Prevents double-suffix bugs like `AIR.PA.DE`. |
| 7 | **Cosmetic preservation rule** | `init-workbook` re-runs on an existing file MUST NOT touch the Main sheet (no rewriting labels, no restyling, no width resets). User edits — including deliberate clearances — survive. Per-cell rule in `_layout_main_sheet`. Same idea for Currencies/Movers headers (text re-asserted because column-name lookups depend on it; styling gated on fresh creation). |
| 8 | **Error reporting: log file + Errors sheet + Status cell** | Replaces the truncated VBA MsgBox. `button_*` functions catch all exceptions and surface them in three places without re-raising. |
| 9 | **Schema versioning** | `SCHEMA_VERSION` constant in code. `VERSION_HISTORY` list. `init-workbook` auto-populates a row per version into the Help sheet (append-only). Bumped on major changes. Filename mirrors. |
| 10 | **Test mode is BEL 20 first 5 + no ETFs + no watchlists** | Brings `rebuild-inventory --test` from ~65 sec to ~4 sec. `get-quotes --test` is 1 ticker (~3 sec). Both still exercise the .info loop and the Wikipedia parser. |

## What was built and shipped (commits, in order)

The repo has ~35 commits on `main`. The major milestones:

### Bug fixes (early session)
- pandas 3.0 compat: `pd.read_html(StringIO(...))` wrap.
- BEL 20 ticker extraction: parses "Euronext Brussels:\xa0SYMBOL"; handles APAM on Amsterdam → `.AS`.
- Excel auto-size loop tripping on NaN under pandas 3.0.
- `datetime.utcnow()` → `datetime.now(dt.UTC)`.
- NIKKEI 225 bullet-list scrape (Wikipedia removed the table).
- Cross-index double-suffix bug (Airbus AIR.PA in DAX, ArcelorMittal MT.AS in CAC40, BT-A in FTSE100).
- AutoFilter re-apply failure when Market not active sheet.
- `sheets.add(after=...)` failure when target is a hidden sheet.

### Architecture
- One-row-per-Symbol refactor (`aggregate_constituents`).
- EuroStoxx 50 added (7th index, multi-exchange).
- Sector + description fields added from yfinance `.info`.
- Watchlist rewrite: Yahoo screener (cookie+crumb) + dataroma 4-page scrape + 7-manager activist aggregate.

### Two-jobs + persistent workbook
- `init-workbook` subcommand (idempotent, cosmetic-preserving).
- `rebuild-inventory` (Job 1).
- `get-quotes` (Job 2) with per-row `Last update` / `Last error`.
- `setup-buttons` subcommand: COM-driven Excel buttons + checkbox + xlwings.conf sheet + VBA auto-import.

### Workbook polish
- 38 curated ETFs (per-index + sector-SPDR + country-iShares).
- Symbol cells hyperlinked to Yahoo Finance.
- % change columns (`1D %`, `1W %`, `1M %`, `6M %`, `1Y %`, `5Y %`) interleaved with prices.
- Comma Style on quote/P-E columns, Percent style on `% change` columns.
- Description column with `wrap_text=False` forced.
- Main!B5 (TestMode) hidden via `;;;` number format.
- Help sheet (auto-populated from VERSION_HISTORY).
- Currencies sheet (EUR/USD/JPY/GBP/CHF history at every lookback).
- Monthly winners sheet (top 1M gainers with no 1D/1W weakness).
- Monthly losers sheet (sustained 1M decline; pre-filtered to Owned?=Yes).
- AutoFilter on Market + Monthly winners + Monthly losers.
- Cross-index ticker normalization (fixes the AIR.PA.DE / MT.AS.PA / BT.A.L bugs).
- Schema versioning v01 → v02 → v03.

### Test pyramid (final state)
- **132 offline** tests (`pytest`): parsing, computation, openpyxl writes, layout, version logic. ~1.5 sec.
- **18 integration** tests (`pytest --integration`): live Wikipedia + Yahoo screener + dataroma + FX network. ~30 sec.
- **11 Excel-COM** tests in `test_xlwings.py` (`pytest --integration test_xlwings.py`): workbook manipulation paths, hidden sheets, AutoFilter resilience, ranking sheet creation, legacy-name migration. Skip automatically when user has Excel open (otherwise `app.quit()` would close it).

### Documentation
- `README.md` rewritten end-to-end for the two-jobs flow + Troubleshooting section.
- `CLAUDE.md` rewritten in three batches matching the architectural shifts.
- `MEMORY.md` with three decision entries (Yahoo deprecation, pandas 3.0 fix, BEL 20 parsing).
- `vba/stocks_picker.bas` — the VBA module that `setup-buttons` auto-imports.

### Hardening / process
- Commit cadence rule: after each major change OR every 5 small changes. Tests must pass.
- Integration tests must also pass for "major" commits.
- Memory rules saved (Claude project memory):
  - Don't ask before fetching from the web.
  - Use WebFetch for one-shot reads when it suffices.
  - Don't kill Excel without explicit approval.
  - Respect user cosmetic edits on stocks.xlsm (rule reinforced twice).
  - xlwings/COM tests must skip when Excel is running.
  - Kill stray tail.exe processes after Monitor usage (caught a "PC won't sleep" issue).

## Current state at handoff

| Aspect | State |
|--------|-------|
| Last commit | `5fabbbf` — v03 Monthly winners + Losers + Owned? + AutoFilters |
| Latest pushed to GitHub | ✓ |
| All tests passing | ✓ — 132 offline, 18 integration, 11 Excel-COM (skip when Excel open) |
| User workbook | `stocks_picker_v02.xlsm` (still on v02 filename despite code at v03) |
| Workbook contents | Main (portfolio incl. BRK-B), Help (credit line), Market (983 rows + AutoFilter), xlwings.conf. Monthly winners/losers + Currencies will appear after the next `get-quotes` run. |
| Excel button macros | VBA imported, buttons wired, Test-mode checkbox at B5 |
| Manual one-time setup completed | ✓ — `xlwings.exe addin install`, VBA Tools → References → xlwings, "Enable Content" on workbook open |

## Outstanding (proposed next steps, not yet done)

In rough priority order:

1. **Rename workbook to `stocks_picker_v03.xlsm`** to match the bumped schema version. One command (close Excel first):
   ```powershell
   mv stocks_picker_v02.xlsm stocks_picker_v03.xlsm
   ```
   No data loss; `_resolve_workbook` will pick it up automatically.

2. **Run `rebuild-inventory` once** so the cross-index normalization fix (commit `498d3ce`) purges the three stale rows from the May 13 run (`BT.A.L`, `AIR.PA.DE`, `MT.AS.PA`).

3. **Run `get-quotes`** to:
   - Populate Currencies + Monthly winners + Monthly losers sheets.
   - Pick up the v03 Help-sheet entry.
   - Verify AutoFilters re-apply cleanly post-rebuild.

4. **Performance: parallelize `.info` calls.** Currently sequential, 0.25s sleep each → ~5 min for 983 tickers. Discussed but explicitly NOT built per "do nothing for now" earlier. Approach:
   - `ThreadPoolExecutor`, 12 workers around `fetch_ticker_info`, exponential backoff on 429s. Expected ~10× speedup.
   - Optional: `--portfolio-only` flag on `get-quotes` for "just refresh what I own" runs.

5. **Job interruptibility.** Discussed but not built. The right approach (after the user corrected my "Excel UI is frozen" claim): a cell-based STOP flag on Main, polled by the `.info` loop every 25 tickers. Cleanly breaks and saves partial state.

6. **Update `CLAUDE.md` for v03.** The doc still describes v02. Should be a single small commit; not blocking anything.

7. **Optional polish** (no urgency):
   - Test mode runs in 4 seconds for rebuild but Monthly winners/losers now compute from the full 983-row Market — already covered by the v03 batch, so test mode produces meaningful rankings.
   - Per-stock daily/weekly/monthly % change columns are present; could add a sparkline using openpyxl's image embedding.
   - Historical FX option (currently uses today's rate for all historical conversions, a deliberate choice).
   - Additional indexes: FTSE 250, SMI, IBEX 35, AEX, EuroStoxx 600.
   - Diff against last run (which tickers entered/exited an index or watchlist).

## How to pick this up next time

1. Read `CLAUDE.md` (project rules + non-obvious knowledge).
2. Read `MEMORY.md` (decision log).
3. Read this `HANDOFF.md` (what we just did).
4. `git log --oneline -20` to see recent commits.
5. Run `pytest` (should be green in ~1.5 sec).
6. Optionally `pytest --integration` if you want network checks.
7. Optionally `pytest --integration test_xlwings.py` (close Excel first).
8. Open `stocks_picker_v02.xlsm` to see current state.

## Known quirks

- **xlwings VBA import** needs Excel's "Trust access to the VBA project object model" setting enabled (one-time per machine). Otherwise `setup-buttons` falls back to logged instructions for a manual Alt+F11 import.
- **`tail -f` survives `TaskStop`/timeout** on Windows. After using the Monitor tool, run `tasklist /FI "IMAGENAME eq tail.exe"` and kill stragglers — they keep the PC from sleeping.
- **`xw.App(visible=False)` can interact badly with an open user Excel session** on some COM version combos. The `test_xlwings.py` fixture refuses to run if Excel is open.
- **AutoFilter via COM `Range.AutoFilter()` is finicky** about active sheet + headless mode. The helper activates Market briefly inside `ScreenUpdating=False` and catches any failure; AutoFilter is best-effort.
- **openpyxl can't author a valid `.xlsm` from scratch** (content-type mismatch). `init-workbook` writes `.xlsx`; `setup-buttons` does the `.xlsm` SaveAs via Excel COM.
