"""
07_detect_stock_splits.py

Scans each ticker's `per_share` sheet for year-over-year jumps in
shares_diluted that look like a stock split (not adjusted automatically).

This script ONLY detects and writes candidates to a CSV for human review.
It never modifies the source Excel files. Confirmed splits get added to
config/confirmed_splits.csv, then 08_apply_split_adjustments.py does the
actual math.

Why ratio-matching instead of just "> X% YoY change":
Large one-off share issuances (secondary offerings, big M&A stock deals)
can also cause big jumps, but they rarely land close to a clean ratio like
2:1, 3:1, 4:1, 7:1, 20:1. Real splits do, almost exactly. Matching against
known clean ratios cuts down false positives a lot.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
EXCEL_DIR = BASE_DIR / "outputs" / "excel"
OUT_DIR = BASE_DIR / "outputs" / "06_quality_check"

# Common clean split ratios seen in US equities. A candidate YoY shares
# multiple within TOLERANCE of one of these is flagged.
KNOWN_RATIOS = [1.5, 2, 3, 4, 5, 6, 7, 10, 20, 25, 50]
TOLERANCE = 0.06  # 6% wiggle room around a clean ratio


def closest_known_ratio(multiple: float):
    best = min(KNOWN_RATIOS, key=lambda r: abs(r - multiple))
    if abs(best - multiple) / best <= TOLERANCE:
        return best
    return None


def detect_for_ticker(ticker: str, excel_path: Path) -> list[dict]:
    try:
        per_share = pd.read_excel(excel_path, sheet_name="per_share")
    except Exception as exc:  # noqa: BLE001
        print(f"  [{ticker}] could not read per_share sheet: {exc}")
        return []

    if "shares_diluted" not in per_share.columns or "fy" not in per_share.columns:
        return []

    per_share = per_share.sort_values("fy").reset_index(drop=True)
    per_share["shares_multiple"] = per_share["shares_diluted"] / per_share["shares_diluted"].shift(1)

    candidates = []
    for _, row in per_share.iterrows():
        multiple = row["shares_multiple"]
        if pd.isna(multiple):
            continue
        matched_ratio = closest_known_ratio(multiple)
        if matched_ratio and matched_ratio > 1.2:
            candidates.append(
                {
                    "ticker": ticker,
                    "fy_effective": int(row["fy"]),
                    "observed_shares_multiple": round(multiple, 3),
                    "closest_known_ratio": matched_ratio,
                    "action": "REVIEW: confirm this is a real split (check 8-K / investor "
                    "relations), then add a row to config/confirmed_splits.csv",
                }
            )
    return candidates


def main():
    all_candidates = []
    for excel_path in sorted(EXCEL_DIR.glob("*_annual_fundamentals.xlsx")):
        ticker = excel_path.stem.replace("_annual_fundamentals", "")
        print(f"Scanning {ticker}...")
        all_candidates.extend(detect_for_ticker(ticker, excel_path))

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUT_DIR / "07_split_candidates.csv"

    if all_candidates:
        pd.DataFrame(all_candidates).to_csv(out_path, index=False, encoding="utf-8-sig")
        print(f"\nFound {len(all_candidates)} candidate split(s). Written to: {out_path}")
        print("Nothing has been changed yet. Review each candidate, then add confirmed")
        print("ones to config/confirmed_splits.csv before running 08_apply_split_adjustments.py.")
    else:
        pd.DataFrame(
            columns=["ticker", "fy_effective", "observed_shares_multiple", "closest_known_ratio", "action"]
        ).to_csv(out_path, index=False, encoding="utf-8-sig")
        print(f"\nNo split candidates found. Empty report written to: {out_path}")


if __name__ == "__main__":
    main()
