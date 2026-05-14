# Stock Picker — Session Handoff (through 2026-05-14)

Two-session pair-programming run with Claude Opus 4.7. Session 01 (2026-05-06 → 2026-05-13) bootstrapped the project from a single-file script into a two-jobs persistent-workbook architecture with the v03 schema. Session 02 (2026-05-13 → 2026-05-14) added performance, polish, a tagged release, and the v04 schema. Read this first if you're picking it back up.

- Session 01 HTML digest: [claude_session_01.html](claude_session_01.html)
- Session 02 HTML digest: [claude_session_02.html](claude_session_02.html)

## Repo

- **GitHub:** https://github.com/alaerts/stock_picker (public, branch `main`)
- **First major release:** [v1.0.0](https://github.com/alaerts/stock_picker/releases/tag/v1.0.0) at `69d159e` — annotated tag pushed on 2026-05-14.
- **Local:** `c:\Dev\my\stock_picker\`
- **Active workbook:** `stocks_picker_v03.xlsm` (renamed from v02 in session 02 to preserve cosmetic edits; `_resolve_workbook` always picks the highest-versioned file regardless of code-side SCHEMA_VERSION)
- **Schema version in code:** `v04` (Industry column added; workbook auto-upgrades on next `rebuild-inventory`)

## Decisions taken (durable)

The 10 from session 01 still hold. Session 02 added:

| # | Decision | Notes |
|---|----------|-------|
| 11 | **Portfolio auto-adoption** | Unresolved `Main!Portfolio` entries get a Yahoo lookup. Found → appended as synthetic constituent with `Indexes="Portfolio"`. Not found → error written to `Main` column C of the offending row. Replaces the prior hard-RuntimeError check. Stale errors clear on the next successful run. |
| 12 | **Parallel .info / prices / watchlists** | `ThreadPoolExecutor` at every layer that's HTTP-bound. .info: 8 workers. Price chunks: 6 workers. Watchlists: 6 workers. Plus a 1.0s inter-chunk sleep removed from `fetch_close_prices` — concurrency cap is the new rate-limiter. |
| 13 | **No session-passing to yfinance** | yfinance 0.2.x raises `YFDataException` on plain `requests.Session` (it requires curl_cffi for TLS fingerprinting). Pinned by 6 regression tests. NEVER add `session=` to `yf.Ticker` / `yf.download` calls again. |
| 14 | **Industry column** | Added between Sector and Watchlists in MARKET_COLUMNS. Sourced from yfinance `.info["industry"]`. |
| 15 | **AutoFilter self-healing on every code touch** | Excel strips openpyxl-set autofilter on manual save. Both `rebuild_inventory` AND `get_quotes` now re-apply Market AutoFilter. Ranking + Currencies sheets re-apply their own. |
| 16 | **Header re-styling on schema bumps** | When a header cell's VALUE is rewritten by `_layout_market_sheet` / `_ensure_ranking_sheet_openpyxl` / `_ensure_currencies_sheet_openpyxl`, the cell's bold + fill are also re-asserted. Preserves user cosmetics on unchanged cells; ensures new schema-bump cells get styled. |
| 17 | **InfoCache SQLite (7-day TTL)** | Slow-changing .info fields cached in `stocks_info_cache.sqlite` next to the workbook (gitignored). Schema migrates transparently via `ALTER TABLE ADD COLUMN` when new fields are added. Used by rebuild_inventory only — get_quotes still hits fresh for P/E. |
| 18 | **Freshness skip in get_quotes** | Rows whose `Last update (UTC)` is within `QUOTE_FRESHNESS_HOURS` (4h) skip the .info call. Prices still refresh every run. Re-runs within the window become near-instant. |
| 19 | **Cell-based STOP control** | Real Excel button is technically infeasible with synchronous `RunPython`. Form-control checkbox writes TRUE to `Main!B13`; Python polls every 25 tickers and breaks cleanly with partial state. |
| 20 | **Busy-flag guard against double-clicks** | VBA wrapper sets `Main!B14` TRUE at job start, FALSE on cleanup. `GuardJobStart()` refuses to launch a second job — prevents two parallel rebuilds from clicking too fast. |
| 21 | **Accounting-comma format on FX rates** | Currencies sheet uses `'_-* #,##0.0000_-;...'` matching Market's EUR price columns, at 4 decimals (2dp would hide intra-week moves). |

## What was built in session 02 (11 commits)

`5fabbbf` (session 01 end) → `5c90809` (current HEAD), with `v1.0.0` tag at `69d159e`:

| Commit | Summary |
|--------|---------|
| `bb87937` | Test-mode rebuild preserves Market (Monthly winners/losers empty round 1) |
| `70d2d4f` | Test-mode get_quotes refreshes top 20 rows (round 2) |
| `9fcbdfd` | Adaptive Owned?=Yes filter on Monthly losers (round 3 — true fix) |
| `fd0203c` | Cell-based STOP + busy-flag + portfolio validation |
| `42856b5` | Speedups A+C+F (parallel .info + freshness skip + SQLite cache) |
| `69d159e` | Auto-adopt unresolved portfolio symbols + max_workers 12→8 |
| `8a530b2` | Batch 1 attempt: shared Yahoo session + parallel prices + parallel watchlists |
| `14118e8` | v04: Industry column + AutoFilter on all sheets + **revert** Batch 1 session experiment |
| `89382a4` | Re-assert Market headers on rebuild + yfinance-session regression tests |
| `5c90809` | Restyle headers + AutoFilter on get_quotes + comma-style FX format |
| (tag) | **v1.0.0** annotated tag on `69d159e` |

## Performance (982 tickers, user's actual workbook)

| Job | Session 01 end | Session 02 cold cache | Session 02 warm cache |
|-----|---------------|------------------------|------------------------|
| `rebuild-inventory` | ~4–10 min | ~57s | ~10s (cache + watchlist hits) |
| `get-quotes` | ~5 min | ~50s | ~5s (freshness skip) |

## Test pyramid (final state)

- **181 offline** tests (`pytest`): parsing, computation, openpyxl writes, layout, version logic, yfinance-session regression guards, header re-styling, AutoFilter survival.
- **16 integration** tests (`pytest --integration`): live Wikipedia + Yahoo screener + dataroma + FX network.
- **14 Excel-COM** tests (`pytest --integration test_xlwings.py`): workbook manipulation, AutoFilter resilience, header re-assertion, ranking-sheet creation, legacy-name migration. Auto-skips when user has Excel open.
- **1 other**.
- **Total: 212 passed.**

## Current state at handoff

| Aspect | State |
|--------|-------|
| Last commit | `5c90809` |
| First tag | `v1.0.0` at `69d159e` |
| Latest pushed to GitHub | ✓ |
| All tests passing | ✓ — 212 across three layers |
| User workbook | `stocks_picker_v03.xlsm` (was v02 — renamed in session 02 to preserve cosmetics) |
| Workbook contents | Main (9 portfolio entries incl. 4 ETFs), Help (credit line), Market (982 rows, v04 layout — Industry at col F), Monthly winners (50), Monthly losers (50), Currencies (4 pairs × 8 lookbacks), xlwings.conf |
| SQLite cache | `stocks_info_cache.sqlite` next to workbook (gitignored) |
| Excel buttons + macros | Wired since session 01; cell-based STOP + busy-flag added session 02 |

## Outstanding (proposed next steps, not yet done)

In rough priority order:

1. **Run `get-quotes` once** after pulling the latest. That single run will:
   - Re-apply Market AutoFilter (Excel likely stripped the prior one)
   - Re-style any unstyled header cells (e.g. I1 will become bold)
   - Re-format Currencies cells to accounting-comma style
   - Refresh prices + P/E

2. **Optional `--full` CLI flag.** Right now there's no command-line way to force full mode when the workbook has TestMode=TRUE — you have to either pass nothing (reads B5) or `--test` (forces True). Session 02 worked around this by calling `get_quotes(test_mode=False)` from a one-off Python invocation when needed. Adding a `--full` flag would be a 4-line change.

3. **Optional Batch 2 (future speedup, more code).** Discussed in session 02 but not built:
   - Price-history SQLite cache (similar shape to `InfoCache`) — yesterday's close never changes; only the last few days move. Would cut warm-rerun `get_quotes` from ~50s to ~5s.
   - Watchlist results cache (24h TTL) — Berkshire's positions don't change between morning and afternoon runs. Same SQLite store.

4. **Optional reports** (from the brainstorm in session 02, ranked):
   - **Sector dashboard** — one row per sector with aggregates. Cheapest insight-per-LoC.
   - **Owned positions deep-dive** — pre-curated `Owned?=Yes` sheet with totals.
   - **Drawdown opportunities** — stocks furthest below their 5Y high.
   - **Weekly index/watchlist diff** — needs snapshot persistence.
   - **Daily email digest** — Pi-cron friendly, sends to Gmail.
   - Discuss before implementing.

5. **README + CLAUDE.md** still describe v03. Should mention v04 features (Industry column, accounting-comma FX format, auto-adoption, freshness skip).

## How to pick this up next time

1. Read `CLAUDE.md` (project rules + non-obvious knowledge).
2. Read `MEMORY.md` (decision log).
3. Read this `HANDOFF.md` (what we built across the two sessions).
4. `git log --oneline -25` + `git tag --list` to see recent commits + the v1.0.0 tag.
5. Run `pytest` (should be green in ~2 sec, 181 tests).
6. Optionally `pytest --integration` if you want network checks (16 more).
7. Optionally `pytest --integration test_xlwings.py` (close Excel first; 14 more).
8. Open `stocks_picker_v03.xlsm` to see current state.

## Known quirks (carried over + new)

From session 01:
- **xlwings VBA import** needs Excel's "Trust access to the VBA project object model" setting enabled (one-time per machine).
- **`tail -f` survives `TaskStop`/timeout** on Windows — kill stragglers via `tasklist /FI "IMAGENAME eq tail.exe"` after Monitor use.
- **`xw.App(visible=False)` can interact badly with an open user Excel session** — `test_xlwings.py` fixture refuses to run if Excel is open.
- **AutoFilter via COM `Range.AutoFilter()`** is finicky about active sheet; the helper activates Market briefly inside `ScreenUpdating=False`.
- **openpyxl can't author a valid `.xlsm` from scratch** — `init-workbook` writes `.xlsx`; `setup-buttons` does `.xlsm` SaveAs via COM.

New in session 02:
- **yfinance 0.2.x requires curl_cffi session** — passing a plain `requests.Session` raises `YFDataException` and silently empties `.info`. The 6 regression tests in `test_stocks_report.py` pin this contract.
- **Excel strips openpyxl-set `<autoFilter>` on manual save** — every code path that writes data must re-apply Market AutoFilter to self-heal. Ranking + Currencies helpers already do this for their own sheets.
- **Header cells added on a schema bump need explicit styling** — when MARKET_COLUMNS / RANKING_HEADERS / CURRENCY_HEADERS grows, the new cell's value is written but the styling logic was previously gated on `is_new=True`. Fixed to re-style on every value rewrite.
- **The SQLite info cache `ALTER TABLE`-migrates** on open. Adding a new field to InfoCache (e.g. v04's `industry` column) is transparent — pre-v04 cache files gain the column on next open with NULL values for old rows.
