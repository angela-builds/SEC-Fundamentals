"""
07_detect_stock_splits.py

Scans each ticker's `per_share` sheet for year-over-year jumps in
shares_diluted_as_reported that look like a stock split.

This script ONLY detects and writes candidates to a CSV for human review.
It never modifies the source Excel files. Confirmed splits get added to
config/confirmed_splits.csv, then 08_apply_split_adjustments.py does the
actual math.

Why use shares_diluted_as_reported?
-----------------------------------
`shares_diluted` may already be split-adjusted by the upstream pipeline.
Using that field for split detection can hide real stock splits.

For example, NVDA's reported diluted shares show:

    FY2021 -> FY2022: ~628M -> ~2,535M = ~4.04x
    FY2024 -> FY2025: ~2,494M -> ~24,804M = ~9.95x

These correspond to real 4-for-1 and 10-for-1 stock splits.

Therefore, this detector prioritizes `shares_diluted_as_reported`,
which preserves the originally reported share counts.

Fallback behavior
-----------------
If an older workbook does not contain `shares_diluted_as_reported`,
the script falls back to `shares_diluted` so that the detector remains
compatible with older workbooks.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[1]
EXCEL_DIR = BASE_DIR / "outputs" / "excel"
OUT_DIR = BASE_DIR / "outputs" / "06_quality_check"


# Common clean split ratios seen in US equities.
# A candidate YoY shares multiple within TOLERANCE of one of these
# ratios is flagged for human review.
KNOWN_RATIOS = [1.5, 2, 3, 4, 5, 6, 7, 10, 20, 25, 50]

# 6% wiggle room around a clean ratio.
TOLERANCE = 0.06


def closest_known_ratio(multiple: float):
    """Return the closest known split ratio if within tolerance."""
    best = min(KNOWN_RATIOS, key=lambda r: abs(r - multiple))

    if abs(best - multiple) / best <= TOLERANCE:
        return best

    return None


def detect_for_ticker(ticker: str, excel_path: Path) -> list[dict]:
    """Detect possible stock splits for one ticker."""

    try:
        per_share = pd.read_excel(
            excel_path,
            sheet_name="per_share",
        )
    except Exception as exc:  # noqa: BLE001
        print(f"  [{ticker}] could not read per_share sheet: {exc}")
        return []

    # ------------------------------------------------------------
    # Validate required columns
    # ------------------------------------------------------------
    if "fy" not in per_share.columns:
        print(f"  [{ticker}] missing required column: fy")
        return []

    # ------------------------------------------------------------
    # IMPORTANT:
    # Prefer the original SEC-reported share count.
    #
    # `shares_diluted` may already have been split-adjusted by
    # the upstream pipeline, which can hide real stock splits.
    # ------------------------------------------------------------
    if "shares_diluted_as_reported" in per_share.columns:
        shares_column = "shares_diluted_as_reported"
        print(
            f"  [{ticker}] using {shares_column} for split detection"
        )

    elif "shares_diluted" in per_share.columns:
        # Backward compatibility for older workbooks.
        shares_column = "shares_diluted"
        print(
            f"  [{ticker}] WARNING: {shares_column} used as fallback "
            f"because shares_diluted_as_reported is unavailable"
        )

    else:
        print(
            f"  [{ticker}] missing both "
            f"shares_diluted_as_reported and shares_diluted"
        )
        return []

    # ------------------------------------------------------------
    # Prepare data
    # ------------------------------------------------------------
    per_share = per_share.sort_values("fy").reset_index(drop=True)

    per_share[shares_column] = pd.to_numeric(
        per_share[shares_column],
        errors="coerce",
    )

    # ------------------------------------------------------------
    # Calculate year-over-year shares multiple
    # ------------------------------------------------------------
    per_share["shares_multiple"] = (
        per_share[shares_column]
        / per_share[shares_column].shift(1)
    )

    candidates = []

    for _, row in per_share.iterrows():

        multiple = row["shares_multiple"]

        if pd.isna(multiple):
            continue

        # We only care about upward jumps.
        if multiple <= 1.2:
            continue

        matched_ratio = closest_known_ratio(multiple)

        if matched_ratio and matched_ratio > 1.2:

            candidates.append(
                {
                    "ticker": ticker,
                    "fy_effective": int(row["fy"]),
                    "observed_shares_multiple": round(
                        float(multiple),
                        3,
                    ),
                    "closest_known_ratio": matched_ratio,
                    "detection_source": shares_column,
                    "action": (
                        "REVIEW: confirm this is a real split "
                        "(check 8-K / investor relations), "
                        "then add a row to "
                        "config/confirmed_splits.csv"
                    ),
                }
            )

    return candidates


def main():
    all_candidates = []

    for excel_path in sorted(
        EXCEL_DIR.glob("*_annual_fundamentals.xlsx")
    ):
        ticker = excel_path.stem.replace(
            "_annual_fundamentals",
            "",
        )

        print(f"Scanning {ticker}...")

        all_candidates.extend(
            detect_for_ticker(
                ticker,
                excel_path,
            )
        )

    # ------------------------------------------------------------
    # Write output
    # ------------------------------------------------------------
    OUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    out_path = OUT_DIR / "07_split_candidates.csv"

    columns = [
        "ticker",
        "fy_effective",
        "observed_shares_multiple",
        "closest_known_ratio",
        "detection_source",
        "action",
    ]

    if all_candidates:

        pd.DataFrame(
            all_candidates,
            columns=columns,
        ).to_csv(
            out_path,
            index=False,
            encoding="utf-8-sig",
        )

        print(
            f"\nFound {len(all_candidates)} candidate split(s). "
            f"Written to: {out_path}"
        )

        print(
            "Nothing has been changed yet. "
            "Review each candidate, then add confirmed "
            "ones to config/confirmed_splits.csv before "
            "running 08_apply_split_adjustments.py."
        )

    else:

        pd.DataFrame(
            columns=columns
        ).to_csv(
            out_path,
            index=False,
            encoding="utf-8-sig",
        )

        print(
            f"\nNo split candidates found. "
            f"Empty report written to: {out_path}"
        )


if __name__ == "__main__":
    main()
