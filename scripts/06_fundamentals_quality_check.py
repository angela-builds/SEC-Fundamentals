from __future__ import annotations

import argparse
import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd


REQUIRED_SHEETS = {
    "per_share": [
        "fy",
        "eps_diluted",
        "shares_diluted",
        "dividends_paid",
        "dividend_per_share",
        "book_value_per_share",
    ],
    "cash_flow": [
        "fy",
        "operating_cash_flow",
        "capex",
        "fcf",
        "fcf_margin",
        "fcf_growth_yoy",
    ],
    "profitability": [
        "fy",
        "revenue",
        "revenue_growth_yoy",
        "net_income",
        "net_margin",
        "roe",
    ],
    "balance_sheet": [
        "fy",
        "equity",
        "cash",
        "total_debt",
        "invested_capital",
        "debt_to_equity",
    ],
    "valuation": [
        "fy",
        "fiscal_start",
        "fiscal_end",
        "eps_diluted",
        "fiscal_year_high_close",
        "fiscal_year_high_date",
        "fiscal_year_high_pe",
        "current_price",
        "current_price_date",
        "current_pe_using_latest_eps",
    ],
    "raw_annual_facts": [
        "fy",
        "start",
        "end",
        "period_days",
        "val",
        "form",
        "filed",
        "accn",
        "fact_type",
        "metric",
        "sec_tag",
        "unit",
    ],
    "data_notes": [
        "fy",
        "metric",
        "fact_type",
        "sec_tag",
        "unit",
        "period_days",
        "form",
        "filed",
        "accn",
    ],
}


METRIC_SOURCE_FIELDS = {
    "eps_diluted": ["eps_diluted"],
    "shares_diluted": ["shares_diluted"],
    "dividend_per_share": ["dividends_paid", "shares_diluted"],
    "book_value_per_share": ["equity", "shares_diluted"],
    "fcf": ["operating_cash_flow", "capex"],
    "fcf_margin": ["fcf", "revenue"],
    "net_margin": ["net_income", "revenue"],
    "roe": ["net_income", "equity", "prior_equity"],
    "invested_capital": ["equity", "total_debt", "cash"],
    "debt_to_equity": ["total_debt", "equity"],
    "fiscal_year_high_pe": ["fiscal_year_high_close", "eps_diluted"],
    "current_pe_using_latest_eps": ["current_price", "latest_eps_diluted"],
}


CHECK_DEFINITIONS = [
    {
        "check_id": "dividend_per_share_formula",
        "sheet": "per_share",
        "metric": "dividend_per_share",
        "formula": "dividends_paid / shares_diluted",
        "fields": ["dividends_paid", "shares_diluted", "dividend_per_share"],
    },
    {
        "check_id": "book_value_per_share_formula",
        "sheet": "combined",
        "metric": "book_value_per_share",
        "formula": "equity / shares_diluted",
        "fields": ["equity", "shares_diluted", "book_value_per_share"],
    },
    {
        "check_id": "fcf_formula",
        "sheet": "cash_flow",
        "metric": "fcf",
        "formula": "operating_cash_flow - capex",
        "fields": ["operating_cash_flow", "capex", "fcf"],
    },
    {
        "check_id": "fcf_margin_formula",
        "sheet": "combined",
        "metric": "fcf_margin",
        "formula": "fcf / revenue",
        "fields": ["fcf", "revenue", "fcf_margin"],
    },
    {
        "check_id": "net_margin_formula",
        "sheet": "profitability",
        "metric": "net_margin",
        "formula": "net_income / revenue",
        "fields": ["net_income", "revenue", "net_margin"],
    },
    {
        "check_id": "roe_formula",
        "sheet": "combined",
        "metric": "roe",
        "formula": "net_income / average equity",
        "fields": ["net_income", "equity", "prior_equity", "roe"],
    },
    {
        "check_id": "invested_capital_formula",
        "sheet": "balance_sheet",
        "metric": "invested_capital",
        "formula": "equity + total_debt - cash",
        "fields": ["equity", "total_debt", "cash", "invested_capital"],
    },
    {
        "check_id": "debt_to_equity_formula",
        "sheet": "balance_sheet",
        "metric": "debt_to_equity",
        "formula": "total_debt / equity",
        "fields": ["total_debt", "equity", "debt_to_equity"],
    },
    {
        "check_id": "fiscal_year_high_pe_formula",
        "sheet": "valuation",
        "metric": "fiscal_year_high_pe",
        "formula": "fiscal_year_high_close / eps_diluted",
        "fields": ["fiscal_year_high_close", "eps_diluted", "fiscal_year_high_pe"],
    },
]


@dataclass(frozen=True)
class InputFile:
    ticker: str
    version: int
    path: Path


def parse_input_file(path: Path) -> InputFile | None:
    pattern = r"(?:(?P<version>\d+)-)?(?P<ticker>[A-Za-z.]+)_annual_fundamentals\.xlsx$"
    match = re.match(pattern, path.name)

    if not match:
        return None

    version = int(match.group("version")) if match.group("version") else 0
    return InputFile(
        ticker=match.group("ticker").upper(),
        version=version,
        path=path,
    )


def latest_files(input_dir: Path) -> list[InputFile]:
    files: dict[str, InputFile] = {}

    for path in input_dir.glob("*_annual_fundamentals.xlsx"):
        parsed = parse_input_file(path)

        if parsed is None:
            continue

        current = files.get(parsed.ticker)

        if current is None or parsed.version > current.version:
            files[parsed.ticker] = parsed

    return sorted(files.values(), key=lambda item: item.ticker)


def safe_divide(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    return numerator / denominator.replace({0: pd.NA})


def relative_diff(actual: pd.Series, expected: pd.Series) -> pd.Series:
    denominator = expected.abs().where(expected.abs() > 1e-9, 1.0)
    return (actual - expected).abs() / denominator


def issue(
    ticker: str,
    input_file: str,
    severity: str,
    check_type: str,
    fy: int | str | None,
    metric: str,
    issue_summary: str,
    actual=None,
    expected=None,
    difference=None,
    source_sheet: str = "",
    backcheck_fields: Iterable[str] = (),
    action: str = "",
) -> dict:
    return {
        "ticker": ticker,
        "input_file": input_file,
        "severity": severity,
        "check_type": check_type,
        "fy": "" if fy is None else fy,
        "metric": metric,
        "issue_summary": issue_summary,
        "actual": actual,
        "expected": expected,
        "difference": difference,
        "source_sheet": source_sheet,
        "backcheck_fields": ", ".join(backcheck_fields),
        "action": action,
    }


def read_workbook(path: Path) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(path)
    return {
        sheet: pd.read_excel(path, sheet_name=sheet)
        for sheet in xls.sheet_names
    }


def combine_core_tables(sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    combined = sheets["per_share"].copy()

    for sheet_name in ["cash_flow", "profitability", "balance_sheet", "valuation"]:
        df = sheets[sheet_name].copy()

        overlap = [
            col for col in df.columns
            if col in combined.columns and col != "fy"
        ]

        df = df.drop(columns=overlap)
        combined = combined.merge(df, on="fy", how="outer")

    combined = combined.sort_values("fy").reset_index(drop=True)

    if "equity" in combined.columns:
        combined["prior_equity"] = combined["equity"].shift(1)

        raw = sheets.get("raw_annual_facts")

        if raw is not None and {"fy", "metric", "val"}.issubset(raw.columns):
            equity_by_fy = raw[raw["metric"] == "equity"].dropna(subset=["fy", "val"])
            equity_map = dict(zip(equity_by_fy["fy"], equity_by_fy["val"]))

            missing_prior = combined["prior_equity"].isna()
            combined.loc[missing_prior, "prior_equity"] = combined.loc[
                missing_prior, "fy"
            ].map(lambda fy: equity_map.get(fy - 1))

    return combined


def expected_for_check(df: pd.DataFrame, check_id: str) -> pd.Series:
    if check_id == "dividend_per_share_formula":
        return safe_divide(df["dividends_paid"], df["shares_diluted"])

    if check_id == "book_value_per_share_formula":
        return safe_divide(df["equity"], df["shares_diluted"])

    if check_id == "fcf_formula":
        return df["operating_cash_flow"] - df["capex"]

    if check_id == "fcf_margin_formula":
        return safe_divide(df["fcf"], df["revenue"])

    if check_id == "net_margin_formula":
        return safe_divide(df["net_income"], df["revenue"])

    if check_id == "roe_formula":
        average_equity = (df["equity"] + df["prior_equity"]) / 2
        return safe_divide(df["net_income"], average_equity)

    if check_id == "invested_capital_formula":
        return df["equity"] + df["total_debt"] - df["cash"]

    if check_id == "debt_to_equity_formula":
        return safe_divide(df["total_debt"], df["equity"])

    if check_id == "fiscal_year_high_pe_formula":
        return safe_divide(df["fiscal_year_high_close"], df["eps_diluted"])

    raise ValueError(f"Unknown check_id: {check_id}")


def add_schema_checks(
    ticker: str,
    input_file: str,
    sheets: dict[str, pd.DataFrame],
    issues: list[dict],
) -> None:
    for sheet_name, required_cols in REQUIRED_SHEETS.items():
        if sheet_name not in sheets:
            issues.append(
                issue(
                    ticker,
                    input_file,
                    "ERROR",
                    "schema",
                    None,
                    sheet_name,
                    f"Missing required sheet: {sheet_name}",
                    source_sheet=sheet_name,
                    action="確認前面資料整理工具是否有輸出此工作表。",
                )
            )
            continue

        missing_cols = [
            col for col in required_cols
            if col not in sheets[sheet_name].columns
        ]

        if missing_cols:
            issues.append(
                issue(
                    ticker,
                    input_file,
                    "ERROR",
                    "schema",
                    None,
                    sheet_name,
                    f"Missing required columns: {', '.join(missing_cols)}",
                    source_sheet=sheet_name,
                    action="回到資料整理腳本確認欄位命名或輸出邏輯。",
                )
            )


def add_missing_checks(
    ticker: str,
    input_file: str,
    combined: pd.DataFrame,
    issues: list[dict],
) -> None:
    key_metrics = [
        "eps_diluted",
        "dividend_per_share",
        "shares_diluted",
        "book_value_per_share",
        "fcf",
        "fcf_margin",
        "net_margin",
        "revenue",
        "net_income",
        "equity",
    ]

    for metric in key_metrics:
        if metric not in combined.columns:
            continue

        for _, row in combined[combined[metric].isna()].iterrows():
            issues.append(
                issue(
                    ticker,
                    input_file,
                    "ERROR",
                    "missing_value",
                    row.get("fy"),
                    metric,
                    "Metric is missing.",
                    source_sheet="combined",
                    backcheck_fields=METRIC_SOURCE_FIELDS.get(metric, [metric]),
                    action="到 raw_annual_facts / data_notes 檢查該年度是否缺少 SEC tag 或被篩選掉。",
                )
            )


def add_formula_checks(
    ticker: str,
    input_file: str,
    combined: pd.DataFrame,
    issues: list[dict],
    formula_rows: list[dict],
    tolerance: float,
) -> None:
    for definition in CHECK_DEFINITIONS:
        fields = definition["fields"]

        if any(field not in combined.columns for field in fields):
            continue

        metric = definition["metric"]
        actual = pd.to_numeric(combined[metric], errors="coerce")
        expected = pd.to_numeric(
            expected_for_check(combined, definition["check_id"]),
            errors="coerce",
        )
        diff = relative_diff(actual, expected)

        for idx, row in combined.iterrows():
            fy = row.get("fy")
            actual_value = actual.iloc[idx]
            expected_value = expected.iloc[idx]
            diff_value = diff.iloc[idx]

            status = "PASS"

            if pd.isna(actual_value) or pd.isna(expected_value):
                status = "SKIP"

            elif diff_value > tolerance:
                status = "FAIL"

                issues.append(
                    issue(
                        ticker,
                        input_file,
                        "ERROR",
                        "formula_consistency",
                        fy,
                        metric,
                        f"{definition['formula']} does not match reported {metric}.",
                        actual=float(actual_value),
                        expected=float(expected_value),
                        difference=float(diff_value),
                        source_sheet=definition["sheet"],
                        backcheck_fields=METRIC_SOURCE_FIELDS.get(metric, fields),
                        action="優先回查公式欄位，再看 raw_annual_facts 中該年度 SEC tag 是否取錯或方向處理錯。",
                    )
                )

            formula_rows.append(
                {
                    "ticker": ticker,
                    "input_file": input_file,
                    "fy": fy,
                    "check_id": definition["check_id"],
                    "metric": metric,
                    "formula": definition["formula"],
                    "actual": None if pd.isna(actual_value) else float(actual_value),
                    "expected": None if pd.isna(expected_value) else float(expected_value),
                    "relative_difference": None if pd.isna(diff_value) else float(diff_value),
                    "status": status,
                }
            )


def add_outlier_checks(
    ticker: str,
    input_file: str,
    combined: pd.DataFrame,
    issues: list[dict],
) -> None:
    checks = [
        (
            "revenue",
            "<= 0",
            lambda s: s <= 0,
            "Revenue is non-positive.",
        ),
        (
            "shares_diluted",
            "<= 0",
            lambda s: s <= 0,
            "Diluted shares are non-positive.",
        ),
        (
            "equity",
            "<= 0",
            lambda s: s <= 0,
            "Equity is non-positive.",
        ),
        (
            "fcf_margin",
            "< -50% or > 60%",
            lambda s: (s < -0.5) | (s > 0.6),
            "FCF margin is outside a broad sanity range.",
        ),
        (
            "net_margin",
            "< -50% or > 60%",
            lambda s: (s < -0.5) | (s > 0.6),
            "Net margin is outside a broad sanity range.",
        ),
        (
            "debt_to_equity",
            "< 0 or > 5",
            lambda s: (s < 0) | (s > 5),
            "Debt-to-equity is outside a broad sanity range.",
        ),
    ]

    for metric, rule, mask_func, summary in checks:
        if metric not in combined.columns:
            continue

        values = pd.to_numeric(combined[metric], errors="coerce")

        for _, row in combined[mask_func(values).fillna(False)].iterrows():
            issues.append(
                issue(
                    ticker,
                    input_file,
                    "WARNING",
                    "sanity_range",
                    row.get("fy"),
                    metric,
                    f"{summary} Rule: {rule}",
                    actual=row.get(metric),
                    source_sheet="combined",
                    backcheck_fields=METRIC_SOURCE_FIELDS.get(metric, [metric]),
                    action="這不一定是錯，但應回查原始欄位與公司當年事件。",
                )
            )

    yoy_checks = [
        ("revenue", 0.50, "Revenue YoY change is above 50%."),
        ("eps_diluted", 1.00, "EPS YoY change is above 100%."),
        ("fcf", 1.00, "FCF YoY change is above 100%."),
        ("shares_diluted", 0.20, "Diluted share count YoY change is above 20%."),
    ]

    for metric, threshold, summary in yoy_checks:
        if metric not in combined.columns:
            continue

        values = pd.to_numeric(combined[metric], errors="coerce")
        yoy = values.pct_change()

        for idx, diff in yoy[abs(yoy) > threshold].items():
            issues.append(
                issue(
                    ticker,
                    input_file,
                    "WARNING",
                    "large_yoy_change",
                    combined.loc[idx, "fy"],
                    metric,
                    summary,
                    actual=float(values.loc[idx]) if not pd.isna(values.loc[idx]) else None,
                    expected="Prior year comparison",
                    difference=float(diff) if not pd.isna(diff) and math.isfinite(diff) else None,
                    source_sheet="combined",
                    backcheck_fields=METRIC_SOURCE_FIELDS.get(metric, [metric]),
                    action="檢查是否為一次性事件、拆股、併購、會計分類或資料抓取問題。",
                )
            )


def add_raw_data_checks(
    ticker: str,
    input_file: str,
    sheets: dict[str, pd.DataFrame],
    issues: list[dict],
) -> None:
    if "raw_annual_facts" not in sheets:
        return

    raw = sheets["raw_annual_facts"].copy()

    if "period_days" in raw.columns:
        mask = (
            raw["period_days"].notna()
            & ((raw["period_days"] < 330) | (raw["period_days"] > 380))
        )

        for _, row in raw[mask].iterrows():
            issues.append(
                issue(
                    ticker,
                    input_file,
                    "WARNING",
                    "raw_period",
                    row.get("fy"),
                    row.get("metric"),
                    "Raw annual fact period_days is outside 330-380 days.",
                    actual=row.get("period_days"),
                    source_sheet="raw_annual_facts",
                    backcheck_fields=[row.get("metric")],
                    action="確認該 SEC fact 是否為完整年度資料，而不是季度、過渡期或重述資料。",
                )
            )


def build_source_map(
    ticker: str,
    input_file: str,
    sheets: dict[str, pd.DataFrame],
) -> list[dict]:
    if "data_notes" not in sheets:
        return []

    data_notes = sheets["data_notes"].copy()
    rows = []

    for _, row in data_notes.iterrows():
        rows.append(
            {
                "ticker": ticker,
                "input_file": input_file,
                "fy": row.get("fy"),
                "metric": row.get("metric"),
                "fact_type": row.get("fact_type"),
                "sec_tag": row.get("sec_tag"),
                "unit": row.get("unit"),
                "period_days": row.get("period_days"),
                "form": row.get("form"),
                "filed": str(row.get("filed"))[:10]
                if not pd.isna(row.get("filed"))
                else "",
                "accn": row.get("accn"),
            }
        )

    return rows


def build_summary(
    files: list[InputFile],
    issues: list[dict],
    formula_rows: list[dict],
) -> list[dict]:
    issue_df = pd.DataFrame(issues)
    formula_df = pd.DataFrame(formula_rows)
    rows = []

    for item in files:
        ticker_issues = (
            issue_df[issue_df["ticker"] == item.ticker]
            if not issue_df.empty
            else pd.DataFrame()
        )

        ticker_formula = (
            formula_df[formula_df["ticker"] == item.ticker]
            if not formula_df.empty
            else pd.DataFrame()
        )

        error_count = (
            int((ticker_issues["severity"] == "ERROR").sum())
            if not ticker_issues.empty
            else 0
        )

        warning_count = (
            int((ticker_issues["severity"] == "WARNING").sum())
            if not ticker_issues.empty
            else 0
        )

        info_count = (
            int((ticker_issues["severity"] == "INFO").sum())
            if not ticker_issues.empty
            else 0
        )

        formula_fail_count = (
            int((ticker_formula["status"] == "FAIL").sum())
            if not ticker_formula.empty
            else 0
        )

        formula_pass_count = (
            int((ticker_formula["status"] == "PASS").sum())
            if not ticker_formula.empty
            else 0
        )

        overall_status = (
            "Needs review"
            if error_count > 0 or warning_count > 0
            else "OK"
        )

        rows.append(
            {
                "ticker": item.ticker,
                "input_file": item.path.name,
                "error_count": error_count,
                "warning_count": warning_count,
                "info_count": info_count,
                "formula_fail_count": formula_fail_count,
                "formula_pass_count": formula_pass_count,
                "overall_status": overall_status,
            }
        )

    return rows


def normalize_records(records: list[dict]) -> list[dict]:
    normalized = []

    for record in records:
        clean = {}

        for key, value in record.items():
            if pd.isna(value):
                clean[key] = None

            elif isinstance(value, np.integer):
                clean[key] = int(value)

            elif isinstance(value, np.floating):
                clean[key] = float(value) if math.isfinite(float(value)) else None

            elif isinstance(value, float) and not math.isfinite(value):
                clean[key] = None

            else:
                clean[key] = value

        normalized.append(clean)

    return normalized


def write_csv(path: Path, rows: list[dict]) -> None:
    df = pd.DataFrame(rows)
    df.to_csv(path, index=False, encoding="utf-8-sig")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="06 quality checks for annual fundamentals workbooks."
    )

    parser.add_argument(
        "--input-dir",
        type=Path,
        default=Path("."),
        help="Folder containing *_annual_fundamentals.xlsx files.",
    )

    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("outputs/06_quality_check"),
        help="Folder for quality check outputs.",
    )

    parser.add_argument(
        "--tolerance",
        type=float,
        default=0.001,
        help="Relative formula tolerance. Default: 0.1%%.",
    )

    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)

    files = latest_files(args.input_dir)

    if not files:
        raise FileNotFoundError(
            f"No *_annual_fundamentals.xlsx files found in {args.input_dir}"
        )

    issues: list[dict] = []
    formula_rows: list[dict] = []
    source_rows: list[dict] = []

    for item in files:
        sheets = read_workbook(item.path)

        add_schema_checks(item.ticker, item.path.name, sheets, issues)

        missing_required = any(
            sheet not in sheets
            for sheet in [
                "per_share",
                "cash_flow",
                "profitability",
                "balance_sheet",
                "valuation",
            ]
        )

        if missing_required:
            continue

        combined = combine_core_tables(sheets)

        add_missing_checks(item.ticker, item.path.name, combined, issues)
        add_formula_checks(
            item.ticker,
            item.path.name,
            combined,
            issues,
            formula_rows,
            args.tolerance,
        )
        add_outlier_checks(item.ticker, item.path.name, combined, issues)
        add_raw_data_checks(item.ticker, item.path.name, sheets, issues)

        source_rows.extend(
            build_source_map(item.ticker, item.path.name, sheets)
        )

    summary_rows = build_summary(files, issues, formula_rows)

    outputs = {
        "summary": normalize_records(summary_rows),
        "quality_issues": normalize_records(issues),
        "formula_checks": normalize_records(formula_rows),
        "source_map": normalize_records(source_rows),
    }

    json_path = args.output_dir / "06_quality_check_report.json"
    json_path.write_text(
        json.dumps(outputs, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    write_csv(args.output_dir / "06_quality_summary.csv", outputs["summary"])
    write_csv(args.output_dir / "06_quality_issues.csv", outputs["quality_issues"])
    write_csv(args.output_dir / "06_formula_checks.csv", outputs["formula_checks"])
    write_csv(args.output_dir / "06_source_map.csv", outputs["source_map"])

    print(f"Checked {len(files)} ticker workbook(s): {', '.join(item.ticker for item in files)}")
    print(f"Quality issues: {len(outputs['quality_issues'])}")
    print(f"Formula checks: {len(outputs['formula_checks'])}")
    print(f"Wrote: {json_path}")


if __name__ == "__main__":
    main()