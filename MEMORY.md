# MEMORY.md — Daily Multi-Index Stock Report

## 2026-05-06, Watchlists rewritten — Yahoo screener API + dataroma.com (RESOLVED)
**What was decided:** Implemented the slim 6-list mapping (Option B). The original 8-list spec is gone:
  1. Recent 52-Week Highs — Yahoo `recent_52_week_highs`
  2. Berkshire Hathaway Portfolio — Yahoo `top_stocks_owned_by_warren_buffet`
  3. Top Quarterly Buys (Super Investors) — dataroma `/m/g/portfolio_b.php?q=q&o=c`
  4. Top Quarterly Sells (Super Investors) — dataroma `/m/g/portfolio_s.php?q=q&o=c`
  5. Most-Held by Super Investors — dataroma `/m/g/portfolio.php`
  6. Activist Hedge Fund Positions — aggregate of dataroma `/m/holdings.php?m={code}` for codes ic, psc, VA, TF, ENG, tp, tci (Icahn / Ackman / ValueAct / Trian / Engaged / Third Point / TCI)

Yahoo screener calls require cookie+crumb (visit fc.yahoo.com → /research-hub/screener/ → /v1/test/getcrumb), then `&useRecordsResponse=true&betaFeatureFlag=true` and pagination via `&start=N` because the API caps at 5 records per call.

CLAUDE.md updated: "Source" line and "Watchlists tracked" table both rewritten.

**Why:** Yahoo deprecated `/u/yahoo-finance/watchlists/{slug}/` pages — 7 of 8 originals returned "Oops, something went wrong" with no scrIds and no embedded payload. β probe of 30 guessed scrIds for the hedge-fund lists all returned "Not Found" — those screeners are gone, not just relocated. Yahoo screener API alone can only deliver 2 of the 6; dataroma.com fills the rest with plain HTML 13F-aggregator data.

**What was rejected:**
  - Headless browser (3b): the data isn't JS-deferred — it's gone from Yahoo entirely.
  - HedgeFollow.com: pages render via JS, no usable static HTML.
  - WhaleWisdom: undisclosed paywall structure on the public landing page.
  - 8-list mapping (Option A): adds the activist-buys / activist-sells split — ~22 HTTP calls/run, much overlap with "Activist Hedge Fund Positions".

**Validation:** BEL20+SP500 smoke test on 2026-05-06: 523 rows, 203 SP500 stocks tagged, 76 in multiple lists, 8 of 9 spot-checked Berkshire holdings present. End-to-end run time ~6m20s.

## 2026-05-06, Yahoo public watchlist scraping was broken — 7 of 8 lists deprecated (HISTORICAL)
**What was decided:** Path forward TBD by user. Investigation complete (3a time-boxed probe). Findings:
  - The static-HTML scrape returns 22 tickers per list, but they're page-chrome / sidebar widgets — not the actual watchlist holdings.
  - **Berkshire Hathaway list IS recoverable** via Yahoo's screener API: `https://query1.finance.yahoo.com/v1/finance/screener/predefined/saved?count=N&start=N&useRecordsResponse=true&betaFeatureFlag=true&scrIds=top_stocks_owned_by_warren_buffet&crumb=...` — needs cookie+crumb auth (visit fc.yahoo.com first, then /v1/test/getcrumb). Returns 5 records per page; total = 111 holdings. Records contain `ticker`, `companyName`, `currentShares`, `formType: "13F"`. AAPL/BAC/KO/AXP/OXY etc. all confirmed present.
  - **The other 7 watchlists are dead.** Their `/u/yahoo-finance/watchlists/{slug}/` pages render "Oops, something went wrong" with no scrIds and no content. Confirmed for `fiftytwo-wk-gain`, `crowded-hedge-fund-positions`, `most-bought-by-hedge-funds`. Same root cause assumed for the rest.
  - **Yahoo's public screener catalog** at `/research-hub/screener/` has ~18 screeners (most_actives, day_gainers, recent_52_week_highs, undervalued_large_caps, etc.) but NONE cover hedge funds / activists / smart money / Berkshire-equivalent strategies. The Berkshire scrIds itself is NOT publicly catalogued — only discoverable from the legacy page.
  - 3b (headless browser) won't help: the data isn't lazy-loaded, it's gone. JavaScript can't conjure what Yahoo no longer serves.
**Why:** Discovered during BEL 20 smoke test on 2026-05-06.
**What was rejected so far:** Nothing chosen. Paths on the table:
  - α: Salvage Berkshire via working scrIds; substitute "Largest 52-Week Gains" with `recent_52_week_highs`; drop the 6 hedge-fund / activist / smart-money lists. Update CLAUDE.md.
  - β: Guess undocumented scrIds for the 6 hedge-fund lists (Berkshire's was `top_stocks_owned_by_warren_buffet`, not the slug — patterns may exist). 5-min probe, low odds, free intel.
  - γ: Switch to a different free source for hedge-fund lists (dataroma.com, WhaleWisdom, Insider Monkey). Rewrites scraper, amends CLAUDE.md "Sources" + "Watchlists tracked".
  - δ: Drop the Watchlists column entirely. Cleanest, smallest scope.
  - User-recommended path forward (proposed): β first (cheap), then α + γ combined.

## 2026-05-06, pandas 3.0 compatibility fix
**What was decided:** `pd.read_html(resp.text)` → wrap in `io.StringIO`. Added `import io`. uv resolved pandas as 3.0.2; pandas 3.x removed support for passing literal HTML strings to read_html (it tries to interpret them as filenames).
**Why:** Smoke test crashed at the first Wikipedia fetch with `OSError: Error reading file '<!DOCTYPE html>...'`.
**What was rejected:** Pinning `pandas<3.0` in requirements.txt — would lock the project on an old major and only delay the same fix.

## 2026-05-06, BEL 20 Wikipedia ticker extraction needs special-casing
**What was decided:** Wikipedia's BEL 20 page formats the ticker column as `"Euronext Brussels:\xa0SYMBOL"` or `"Euronext Amsterdam:\xa0SYMBOL"` (one constituent: APAM is on Amsterdam). The generic length filter `between(1, 12)` discarded all 20 rows. Adding a BEL20-specific extractor in `get_index_constituents` to parse the prefix and assign the right Yahoo suffix (`.BR` for Brussels, `.AS` for Amsterdam). `normalize_ticker` updated to leave `.AS` suffix alone for BEL20.
**Why:** Without this, BEL 20 returns 0 constituents and the whole index is silently skipped.
**What was rejected:** Stripping the prefix generically (`s.split(":")[-1]`) would lose the exchange info and force every BEL 20 stock onto `.BR`, which would break APAM (Aperam, listed on Euronext Amsterdam → must be `APAM.AS` on Yahoo).
