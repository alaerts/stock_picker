# Daily Multi-Index Stock Report

A Python script that produces a daily Excel report covering S&P 500, Nikkei 225, FTSE 100, DAX, CAC 40, and BEL 20 stocks (~925 total).

## What's in the report

For each stock, in EUR using today's exchange rate:

- Current price
- Prices 1 day, 1 week, 1 month, 6 months, 1 year, 5 years ago
- Trailing P/E, Forward P/E
- Native currency
- Any Yahoo Finance watchlist memberships (Largest 52-Week Gains, Crowded Hedge Fund Positions, Berkshire Hathaway Portfolio, Smart Money Stocks, Most Sold/Bought by Activist Hedge Funds, Most Bought by Hedge Funds, Activist Hedge Fund Positions)

Output is a single `.xlsx` with a combined sheet plus a metadata tab.

## Setup

```bash
pip install -r requirements.txt
```

## Usage

Smoke test (BEL 20 only, ~1 minute):

```bash
python stocks_report.py --indexes BEL20
```

Full run (all six indexes, 10–20 minutes):

```bash
python stocks_report.py
```

Custom output path:

```bash
python stocks_report.py --output ~/reports/stocks.xlsx
```

Output filename defaults to `stocks_report_YYYY-MM-DD.xlsx` in the current directory.

## Tests

```bash
python -m pytest tests/ -v
```

## Sources

- **yfinance** — Yahoo Finance prices, P/E, currency
- **Wikipedia** — index constituents
- **Yahoo Finance public watchlist pages** — hedge fund / Berkshire memberships

All free, no accounts or API keys required.

## Caveats

- yfinance is unofficial; Yahoo can break it. Run `pip install --upgrade yfinance` if a previously-working run fails.
- Yahoo watchlists are US-only — non-US stocks will always have an empty Watchlists column. Expected.
- 5-year lookback may be blank for recently IPO'd companies. Expected.
- Historical prices are converted at TODAY's FX rate (not the rate on that historical date), so non-EUR stock movements blend price movement with FX movement. This was a deliberate choice.
