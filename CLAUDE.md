# Daily Multi-Index Stock Report — Project Context for Claude

## What this is

A Python script that produces a daily Excel report covering ~925 stocks across S&P 500, Nikkei 225, FTSE 100, DAX, CAC 40, and BEL 20. For each stock: current price + prices at 1d/1w/1mo/6mo/1y/5y ago, all converted to EUR at today's FX rate, plus trailing P/E, forward P/E, native currency, and any Yahoo Finance watchlist memberships.

User is in Brussels (hence EUR base, BEL 20 inclusion). Has a Raspberry Pi running Home Assistant — likely target for scheduling. Technical user.

## Decisions already made — do not re-litigate without asking

These were explicitly chosen by the user in the planning conversation. If you think a different choice is better, raise it as a question rather than changing the implementation.

- **Indexes**: SP500, NIKKEI225, FTSE100, DAX, CAC40, BEL20.
- **FX**: today's rate, applied to ALL historical prices. User chose this over historical rates. Trade-off accepted: mixes price movement with FX movement on non-EUR stocks.
- **Output layout**: SINGLE combined sheet (not one per index). Plus a small Metadata sheet with run timestamp and FX rates.
- **Watchlist column format**: ONE column, comma-separated list names. NOT 8 boolean columns.
- **Lookbacks**: Today, 1D, 1W, 1M, 6M, 1Y, 5Y ago.
- **Lookback semantics**: last available close *on or before* the target date — handles weekends/holidays by rolling back to the previous trading day.
- **Source**: yfinance for prices/PE/currency; Wikipedia for index constituents; Yahoo's predefined-screener API (`/v1/finance/screener/predefined/saved`, with cookie+crumb auth) and dataroma.com (plain HTML "super investor" 13F aggregator) for watchlist memberships. All free, no accounts or API keys. — *Source for watchlists changed 2026-05-06 after Yahoo deprecated the legacy `/u/yahoo-finance/watchlists/{slug}/` pages; see MEMORY.md.*
- **Scheduling**: deferred. User will decide where to run it.

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

- `stocks_report.py` — single-file script, all logic.
- `requirements.txt` — yfinance, pandas, openpyxl, requests, lxml, beautifulsoup4.
- `tests/test_stocks_report.py` — offline unit tests for ticker normalization, FX conversion, lookback resolution.
- `README.md` — user-facing docs.

## Code structure (in `stocks_report.py`)

1. Logging
2. Config: `INDEX_WIKI`, `WATCHLISTS`, `FX_PAIRS`, `LOOKBACKS`, `INDEX_DEFAULT_CCY`
3. `normalize_ticker()` — Wikipedia → Yahoo symbol mapping per index
4. `get_index_constituents()` — Wikipedia table fetcher with auto-detection
5. `fetch_yahoo_screener()`, `fetch_dataroma_tickers()`, `fetch_dataroma_activist_aggregate()`, `fetch_all_watchlists()` — Yahoo predefined-screener API (cookie+crumb) plus dataroma.com plain-HTML scraping
6. `get_fx_rates()` — EURUSD=X, EURJPY=X, EURGBP=X
7. `to_eur()` — incl. GBp pence handling
8. `fetch_close_prices()` — chunked, throttled
9. `fetch_ticker_info()` / `fetch_all_info()` — per-ticker, throttled
10. `price_at_or_before()` — lookback date resolver
11. `build_report()` — orchestrator
12. `write_excel()` — openpyxl with frozen header, auto-sized cols, number formats
13. CLI

Output columns: `Index | Symbol | Name | Short description of the company | Sector | Currency | Today (EUR) | 1D ago (EUR) | 1W ago (EUR) | 1M ago (EUR) | 6M ago (EUR) | 1Y ago (EUR) | 5Y ago (EUR) | P/E (TTM) | Forward P/E | Watchlists`

## Non-obvious things to know

- **Ticker symbol mismatches**: Wikipedia uses BRK.B; Yahoo uses BRK-B. `normalize_ticker()` handles SP500 dot→hyphen, FTSE trailing-dot stripping, and adds suffixes (`.T`, `.L`, `.DE`, `.PA`, `.BR`).
- **GBp pence**: many UK stocks report in pence on yfinance with `currency = "GBp"`. We divide by 100 before converting via the GBP rate. See `to_eur()`.
- **Watchlist sources are fragile**: Yahoo's screener API is undocumented and rate-gates aggressively without cookie+crumb; the Berkshire scrIds is not in the public catalog and could disappear in any Yahoo redesign. Dataroma is plain HTML but its URL structure could change. Failures log warnings and the report still completes with empty Watchlists for the affected entries.
- **Dataroma uses BRK.B / BF.B; Yahoo uses BRK-B / BF-B**: `_extract_dataroma_tickers()` emits both forms so SP500 tickers match regardless of which form the index constituents table contains.
- **Watchlists are US-only**: hedge fund / Berkshire lists contain US tickers. Non-US stocks (BEL/CAC/DAX/FTSE/Nikkei) will always have empty `Watchlists` column. This is expected, not a bug.
- **Watchlist matching uses both the full Yahoo symbol AND the suffix-stripped root** — defensive in case a watchlist surfaces a name differently.
- **`.info` rate limiting**: ~925 calls × 0.25s ≈ 4 min. If Yahoo throttles, the script doesn't crash — those stocks just get blank P/E. Consider caching before scaling up frequency.
- **`yf.download()` with multiple tickers** returns a multi-index DataFrame. `fetch_close_prices()` extracts just the `Close` level. Watch for breaking changes across yfinance major versions.
- **5-year blanks for recent IPOs are correct, not a bug.**
- **The script is tolerant of partial failures**: a missing watchlist, a failing ticker, or a timed-out chunk all log warnings and the run still completes with the rest.

## Testing

Offline unit tests (no network):
```bash
python -m pytest tests/ -v
```

End-to-end smoke test (needs internet, ~1 min, ~20 tickers):
```bash
python stocks_report.py --indexes BEL20
```

After BEL 20 succeeds, validate watchlist column populates by adding S&P 500:
```bash
python stocks_report.py --indexes BEL20,SP500
```

Full run (~10–20 min):
```bash
python stocks_report.py
```

## Open items / next steps (priority order)

1. **First real run**: BEL 20 smoke test. Capture log output. Look for:
   - Any watchlist fetching 0 tickers (likely "Most Bought by Hedge Funds" — unverified slug).
   - Currency column blanks (would indicate yfinance .info issues).
   - 5Y blanks on non-recent IPOs (would be unexpected and indicate a price-fetch problem).
2. **Validate all 8 watchlist slugs** return non-empty sets. If any fail, find the correct slug at `https://finance.yahoo.com/u/yahoo-finance/watchlists/` (or its sibling pages).
3. **Scheduling**: user TBD. Most likely Pi cron given Home Assistant context. If so, add a wrapper shell script: activate venv → run script → rotate outputs.
4. **Optional retry/cache layer** for `.info` if rate-limiting becomes a problem at full 925-ticker scale. SQLite cache keyed by ticker, with TTL.
5. **Optional**: email or cloud upload of the result.

## Extension ideas — discuss before implementing

- Per-stock daily/weekly/monthly % change columns (computable in Excel formula or in pandas).
- Historical FX option (we deliberately chose today's rate; could be a `--historical-fx` flag).
- Additional indexes (FTSE 250, Euro Stoxx 50, SMI, IBEX 35, AEX).
- Diff against yesterday's report (which tickers entered/exited a watchlist).
- Sector/industry columns from `info`.
- A "% change" sparkline column using openpyxl's image embedding.

## What NOT to change without asking

- FX semantics (today's rate). The user explicitly chose this.
- Single combined sheet layout.
- Comma-separated single Watchlists column (not 8 booleans).
- Free / no-API-key requirement.
- Throttling and chunk delays — Yahoo will rate-limit aggressively if removed.

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

