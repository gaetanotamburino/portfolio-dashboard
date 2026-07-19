"""
run_pipeline.py
────────────────
Full portfolio pipeline orchestrator.

Steps
  1  trades_parser   — ingest broker exports from imports/
  2  fetch_prices    — fetch closing prices for all instruments
  3  build_analytics — rebuild positions + portfolio_summary
  4  export_to_excel — write dashboard sheet in Portafogliov4.xlsm

Usage
  python scripts/run_pipeline.py                  full run
  python scripts/run_pipeline.py --skip-trades    skip step 1 (no new imports)
  python scripts/run_pipeline.py --skip-prices    skip step 2 (prices already fresh)
  python scripts/run_pipeline.py --export-only    only step 4 (re-export existing data)
  python scripts/run_pipeline.py --live           intraday fetch (step 2) + Excel export (step 4)
"""

import argparse
import logging
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("pipeline")


def _run(label: str, script: str, extra_args: list | None = None) -> None:
    path = SCRIPTS / script
    log.info("──── %s", label)
    t0 = time.perf_counter()
    result = subprocess.run(
        [sys.executable, str(path)] + (extra_args or []),
        check=False,
    )
    elapsed = time.perf_counter() - t0
    if result.returncode != 0:
        log.error("     FAILED (exit %d) after %.1fs", result.returncode, elapsed)
        sys.exit(result.returncode)
    log.info("     done in %.1fs", elapsed)


def main() -> None:
    parser = argparse.ArgumentParser(description="Portfolio pipeline orchestrator")
    parser.add_argument("--skip-trades",  action="store_true", help="skip step 1")
    parser.add_argument("--skip-prices",  action="store_true", help="skip step 2")
    parser.add_argument("--export-only",  action="store_true", help="only step 4")
    parser.add_argument("--live",         action="store_true", help="live refresh only: intraday fetch + Excel export")
    args = parser.parse_args()

    log.info("Pipeline start — %s", datetime.now().strftime("%Y-%m-%d %H:%M"))
    t_start = time.perf_counter()

    if args.live:
        _run("Step 2 · fetch_prices (live)", "fetch_prices.py", extra_args=["--live"])
        _run("Step 3 · build_analytics (live)", "build_analytics.py", extra_args=["--live"])
        _run("Step 4 · export_to_excel",     "export_to_excel.py")
        log.info("Live refresh complete — %.1fs total", time.perf_counter() - t_start)
        return

    if not args.export_only:
        if not args.skip_trades:
            _run("Step 1 · trades_parser",   "trades_parser.py")
        else:
            log.info("Step 1 · trades_parser   [skipped]")

        if not args.skip_prices:
            _run("Step 2 · fetch_prices",    "fetch_prices.py")
        else:
            log.info("Step 2 · fetch_prices    [skipped]")

        _run("Step 3 · build_analytics", "build_analytics.py")

    _run("Step 4 · export_to_excel",  "export_to_excel.py")

    log.info("Pipeline complete — %.1fs total", time.perf_counter() - t_start)


if __name__ == "__main__":
    main()
