# Portfolio Dashboard

A personal investment-portfolio pipeline that ingests broker trade exports, fetches
daily/live prices, computes positions and P&L, and pushes the results into an Excel
dashboard.

## How it works

Broker export files dropped into `imports/` are parsed into a standardized trades
table in a local SQLite database (`data/portfolio.db`). From there, prices are
fetched, positions and portfolio-level analytics (cost basis, P&L, returns) are
built, and the latest snapshot is written into `Portafogliov4.xlsm` for viewing.

```
imports/*.xls|*.xlsx  →  trades_parser.py  →  portfolio.db (trades)
                                                     │
                              fetch_prices.py  →  portfolio.db (prices, live_prices)
                                                     │
                            build_analytics.py  →  portfolio.db (instruments, positions, portfolio_summary)
                                                     │
                           export_to_excel.py  →  Portafogliov4.xlsm (data, charts sheets)
```

## Structure

- `scripts/trades_parser.py` — parses broker exports (Fineco `.xls`, CA `.xlsx`) from
  `imports/` into `portfolio.db`, and writes a standardized trades CSV to
  `data/trades csv/`. Source files are deleted after a successful parse.
- `scripts/fetch_prices.py` — fetches daily closing prices and (with `--live`)
  intraday prices from Yahoo Finance, Deutsche Börse, and Euronext MOT.
- `scripts/build_analytics.py` — builds the `instruments`, `positions`, and
  `portfolio_summary` tables (cost basis, market value, P&L, period returns).
- `scripts/accrued.py` — bond accrued-interest calculations.
- `scripts/migrate_bonds.py` — one-off migration adding the `bonds` metadata table.
- `scripts/export_to_excel.py` — writes the latest snapshot and a returns chart into
  `Portafogliov4.xlsm`.
- `scripts/run_pipeline.py` — orchestrates the full pipeline (see flags below).
- `middlemanXLSDB.py` / `Portafogliov4.py` — entry point called from inside Excel via
  [xlwings](https://www.xlwings.org/) to push live prices into the open workbook.
- `open_dashboard.bat` — Windows launcher: activates the conda environment, runs a
  live pipeline refresh, and opens the workbook.

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. This repo does **not** include `data/`, `imports/`, or `Portafogliov4.xlsm` —
   they hold personal financial data and are excluded via `.gitignore`. To run the
   pipeline you'll need:
   - An `imports/` folder for dropping broker export files.
   - A `Portafogliov4.xlsm` workbook with `data` and `charts` sheets (created
     automatically by `export_to_excel.py` on first export if missing sheets, but
     the workbook file itself must already exist).
   - `data/portfolio.db` is created automatically on first run.

## Usage

Full pipeline (parse new trades, fetch prices, rebuild analytics, export to Excel):
```bash
python scripts/run_pipeline.py
```

Other modes:
```bash
python scripts/run_pipeline.py --skip-trades   # skip parsing imports/
python scripts/run_pipeline.py --skip-prices   # skip fetching prices
python scripts/run_pipeline.py --export-only   # only re-export existing data
python scripts/run_pipeline.py --live          # intraday price refresh + export
```

On Windows, `open_dashboard.bat` runs a live refresh and opens the workbook in one step.

## Notes

- Account labels (`ACCOUNT_1`, `ACCOUNT_2` in `scripts/trades_parser.py`) identify
  the two tracked accounts and can be renamed there without touching parser logic.
- Ticker resolution (ISIN → Yahoo ticker) is hardcoded per-instrument in
  `scripts/fetch_prices.py`, `scripts/build_analytics.py`, and `middlemanXLSDB.py`
  — update the relevant map when adding a new instrument.
