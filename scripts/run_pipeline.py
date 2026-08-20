"""
run_pipeline.py

Runs the whole 01 -> 06 pipeline for every ticker in config/tickers.csv
(or a list you pass on the command line), so you don't have to manually
invoke 8 scripts and retype the ticker each time.

07 (split detection) and 08 (split adjustment) are NOT run automatically
here, and that's intentional -- 08 depends on a human having reviewed
07's candidates and confirmed them in config/confirmed_splits.csv first.
Running 08 blindly in a loop would apply adjustments nobody checked.

Usage:
    python scripts\\run_pipeline.py                  # uses config/tickers.csv
    python scripts\\run_pipeline.py AAPL PG NVDA      # explicit list, ignores the csv
    python scripts\\run_pipeline.py --skip-fetch      # reuse existing data/raw/*.json,
                                                       # only re-run 02-06 (fast when you're
                                                       # just re-testing pipeline logic)

After this finishes, run separately:
    python scripts\\07_detect_stock_splits.py
    (review outputs\\06_quality_check\\07_split_candidates.csv, edit config\\confirmed_splits.csv)
    python scripts\\08_apply_split_adjustments.py
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = BASE_DIR / "scripts"
TICKERS_PATH = BASE_DIR / "config" / "tickers.csv"

PIPELINE_STEPS = [
    "01_fetch_sec_companyfacts.py",
    "02_extract_annual_metrics.py",
    "03_add_valuation.py",
    "04_build_summary.py",
    "05_build_charts.py",
]

QUALITY_CHECK_STEP = "06_fundamentals_quality_check.py"


def load_tickers_from_csv() -> list[str]:
    if not TICKERS_PATH.exists():
        return []
    df = pd.read_csv(TICKERS_PATH)
    # Be tolerant of the column being named "ticker" or "tickers".
    col = "ticker" if "ticker" in df.columns else df.columns[0]
    return [str(t).strip().upper() for t in df[col].dropna().tolist()]


LOG_DIR = BASE_DIR / "outputs" / "logs"


def run_step(script_name: str, ticker: str, skip_fetch: bool) -> bool:
    if skip_fetch and script_name == "01_fetch_sec_companyfacts.py":
        print(f"  [skip] {script_name} (--skip-fetch, reusing data/raw/{ticker}_companyfacts.json)")
        return True

    script_path = SCRIPTS_DIR / script_name
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_path = LOG_DIR / f"{ticker}_{script_name.replace('.py', '.log')}"

    result = subprocess.run(
        [sys.executable, str(script_path), ticker],
        cwd=str(BASE_DIR),
        capture_output=True,
        text=True,
    )

    log_path.write_text(
        f"$ python scripts\\{script_name} {ticker}\n\n"
        f"--- stdout ---\n{result.stdout}\n"
        f"--- stderr ---\n{result.stderr}\n"
        f"--- exit code: {result.returncode} ---\n",
        encoding="utf-8",
    )

    if result.returncode != 0:
        print(f"  [FAILED] {script_name} for {ticker} (exit code {result.returncode})")
        # Show the last few lines right here so most failures don't even
        # require opening the log file, but the full detail is always saved.
        tail_lines = (result.stderr or result.stdout).strip().splitlines()[-8:]
        for line in tail_lines:
            print(f"    | {line}")
        print(f"    full log: {log_path}")
        return False

    print(f"  [ok] {script_name}")
    return True


def main():
    args = sys.argv[1:]
    skip_fetch = "--skip-fetch" in args
    args = [a for a in args if a != "--skip-fetch"]

    tickers = [a.upper() for a in args] if args else load_tickers_from_csv()

    if not tickers:
        print(f"No tickers given and none found in {TICKERS_PATH}.")
        print("Either pass tickers on the command line, or fill in config/tickers.csv.")
        return

    print(f"Running pipeline for: {', '.join(tickers)}\n")

    failed = []
    for ticker in tickers:
        print(f"=== {ticker} ===")
        ok = True
        for step in PIPELINE_STEPS:
            ok = run_step(step, ticker, skip_fetch)
            if not ok:
                failed.append((ticker, step))
                break
        print()

    # Quality check runs once across all tickers together, not per-ticker.
    print("=== quality check (all tickers) ===")
    subprocess.run(
        [sys.executable, str(SCRIPTS_DIR / QUALITY_CHECK_STEP), "--input-dir", str(BASE_DIR / "outputs" / "excel")],
        cwd=str(BASE_DIR),
    )

    print()
    if failed:
        print("Some tickers did not complete:")
        for ticker, step in failed:
            log_path = LOG_DIR / f"{ticker}_{step.replace('.py', '.log')}"
            print(f"  - {ticker} stopped at {step}  (log: {log_path})")
    else:
        print("All tickers completed 01-06 successfully.")

    print("\nNext (manual, on purpose):")
    print("  python scripts\\07_detect_stock_splits.py")
    print("  (review outputs\\06_quality_check\\07_split_candidates.csv,")
    print("   confirm real splits in config\\confirmed_splits.csv)")
    print("  python scripts\\08_apply_split_adjustments.py")


if __name__ == "__main__":
    main()
