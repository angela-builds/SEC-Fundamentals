"""
01_fetch_sec_companyfacts.py

Fetch SEC CompanyFacts JSON for a ticker.

New behavior:
    Ticker
      ↓
    SEC company_tickers.json
      ↓
    Resolve CIK automatically
      ↓
    SEC CompanyFacts API
      ↓
    data/raw/{TICKER}_companyfacts.json

The ticker does NOT need to exist in config/tickers.csv.

config/tickers.csv can still be used as a local metadata/cache file,
but it is no longer the source of truth for ticker -> CIK lookup.

Usage:
    python scripts/01_fetch_sec_companyfacts.py AAPL
    python scripts/01_fetch_sec_companyfacts.py NVDA
    python scripts/01_fetch_sec_companyfacts.py AMZN

SEC API:
    https://www.sec.gov/files/company_tickers.json
    https://data.sec.gov/api/xbrl/companyfacts/CIK##########.json

Important:
    SEC requires a declared User-Agent for automated access.
    Set SEC_CONTACT_EMAIL before running the script.
"""

from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

import requests


# =============================================================================
# Paths
# =============================================================================

BASE_DIR = Path(__file__).resolve().parents[1]

RAW_DIR = BASE_DIR / "data" / "raw"
CONFIG_DIR = BASE_DIR / "config"
TICKERS_CSV = CONFIG_DIR / "tickers.csv"


# =============================================================================
# SEC endpoints
# =============================================================================

SEC_TICKER_URL = "https://www.sec.gov/files/company_tickers.json"

SEC_COMPANYFACTS_URL = (
    "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
)


# =============================================================================
# Settings
# =============================================================================

# SEC asks automated clients to identify themselves.
# You can set this in Colab:
#
#     import os
#     os.environ["SEC_CONTACT_EMAIL"] = "your@email.com"
#
# Do NOT hard-code your personal email into GitHub.
SEC_CONTACT_EMAIL = os.getenv(
    "SEC_CONTACT_EMAIL",
    "",
).strip()


REQUEST_TIMEOUT = 60

# Small pause between SEC requests.
# We normally make only 2 requests:
#   1. company_tickers.json
#   2. companyfacts
REQUEST_DELAY_SECONDS = 0.2


# =============================================================================
# HTTP helpers
# =============================================================================

def build_headers() -> dict[str, str]:
    """
    Build SEC-compliant request headers.
    """

    if not SEC_CONTACT_EMAIL:
        raise RuntimeError(
            "SEC_CONTACT_EMAIL is not set.\n\n"
            "Before running this script in Colab, execute:\n\n"
            "    import os\n"
            "    os.environ['SEC_CONTACT_EMAIL'] = 'your@email.com'\n\n"
            "Use your own contact email."
        )

    return {
        "User-Agent": (
            f"SEC Fundamental Data Tool "
            f"(contact: {SEC_CONTACT_EMAIL})"
        ),
        "Accept-Encoding": "gzip, deflate",
}


def get_json(
    url: str,
    headers: dict[str, str],
) -> dict:
    """
    GET JSON from SEC and return parsed JSON.
    """

    print(f"  Fetching: {url}")

    try:
        response = requests.get(
            url,
            headers=headers,
            timeout=REQUEST_TIMEOUT,
        )
    except requests.RequestException as exc:
        raise RuntimeError(
            f"SEC request failed:\n{exc}"
        ) from exc

    if response.status_code != 200:
        preview = response.text[:500].replace("\n", " ")

        raise RuntimeError(
            f"SEC returned HTTP {response.status_code}.\n"
            f"URL: {url}\n"
            f"Response: {preview}"
        )

    try:
        return response.json()
    except ValueError as exc:
        raise RuntimeError(
            f"SEC response was not valid JSON.\n"
            f"URL: {url}"
        ) from exc


# =============================================================================
# Ticker -> CIK
# =============================================================================

def normalize_ticker(ticker: str) -> str:
    """
    Normalize user input.
    """

    ticker = str(ticker).strip().upper()

    if not ticker:
        raise ValueError("Ticker cannot be empty.")

    return ticker


def load_sec_ticker_map(
    headers: dict[str, str],
) -> dict[str, dict]:
    """
    Download SEC's official ticker -> CIK mapping.

    SEC company_tickers.json has records such as:

        {
            "4": {
                "cik_str": 1018724,
                "ticker": "AMZN",
                "title": "AMAZON COM INC"
            }
        }
    """

    data = get_json(
        SEC_TICKER_URL,
        headers,
    )

    if not isinstance(data, dict):
        raise RuntimeError(
            "Unexpected format from SEC company_tickers.json."
        )

    ticker_map: dict[str, dict] = {}

    for item in data.values():

        if not isinstance(item, dict):
            continue

        ticker = item.get("ticker")
        cik = item.get("cik_str")
        title = item.get("title")

        if not ticker or cik is None:
            continue

        ticker = str(ticker).strip().upper()

        ticker_map[ticker] = {
            "cik": int(cik),
            "company_name": str(title or "").strip(),
        }

    return ticker_map


def resolve_ticker(
    ticker: str,
    ticker_map: dict[str, dict],
) -> dict:
    """
    Resolve ticker through SEC's official ticker mapping.
    """

    ticker = normalize_ticker(ticker)

    if ticker not in ticker_map:

        # Show a useful subset rather than dumping hundreds/thousands
        # of SEC tickers.
        suggestions = [
            t
            for t in ticker_map
            if ticker in t
            or t in ticker
        ]

        suggestion_text = ""

        if suggestions:
            suggestion_text = (
                "\nPossible matches: "
                + ", ".join(sorted(suggestions)[:10])
            )

        raise ValueError(
            f"Ticker not found in SEC company_tickers.json: {ticker}"
            f"{suggestion_text}\n\n"
            "Please check that the symbol is an SEC-listed company "
            "ticker and not an ETF/fund/share class symbol that SEC "
            "does not expose in this mapping."
        )

    return ticker_map[ticker]


# =============================================================================
# CompanyFacts
# =============================================================================

def fetch_companyfacts(
    ticker: str,
    cik: int,
    headers: dict[str, str],
) -> dict:
    """
    Fetch SEC CompanyFacts for a CIK.
    """

    cik_padded = f"{int(cik):010d}"

    url = SEC_COMPANYFACTS_URL.format(
        cik=cik_padded
    )

    time.sleep(REQUEST_DELAY_SECONDS)

    companyfacts = get_json(
        url,
        headers,
    )

    if not isinstance(companyfacts, dict):
        raise RuntimeError(
            "SEC CompanyFacts response has unexpected format."
        )

    returned_cik = companyfacts.get("cik")

    if returned_cik is not None:
        if int(returned_cik) != int(cik):
            raise RuntimeError(
                "CIK mismatch between ticker lookup and "
                "CompanyFacts response.\n"
                f"Expected: {cik}\n"
                f"Returned: {returned_cik}"
            )

    returned_name = companyfacts.get(
        "entityName",
        "",
    )

    facts = companyfacts.get(
        "facts",
        {},
    )

    us_gaap = facts.get(
        "us-gaap",
        {},
    )

    print(
        f"  Company: {returned_name or '(unknown)'}"
    )

    print(
        f"  CIK: {cik_padded}"
    )

    print(
        f"  us-gaap tags: {len(us_gaap):,}"
    )

    return companyfacts


# =============================================================================
# Validation
# =============================================================================

def validate_companyfacts(
    ticker: str,
    companyfacts: dict,
) -> None:
    """
    Basic validation only.

    This function does NOT decide whether the company is fully
    compatible with the downstream fundamental-analysis pipeline.

    That decision belongs to 02_extract_annual_metrics.py.

    Here we only make sure CompanyFacts contains the expected
    top-level structure.
    """

    if "facts" not in companyfacts:
        raise RuntimeError(
            f"{ticker}: CompanyFacts does not contain 'facts'."
        )

    facts = companyfacts["facts"]

    if not isinstance(facts, dict):
        raise RuntimeError(
            f"{ticker}: CompanyFacts 'facts' has unexpected format."
        )

    us_gaap = facts.get("us-gaap")

    if not us_gaap:
        print(
            f"  [WARNING] {ticker} has no us-gaap namespace."
        )

        print(
            "            This ticker may be an ETF, foreign issuer, "
            "fund, or otherwise incompatible with the current "
            "US-GAAP fundamental extraction pipeline."
        )

        print(
            "            The JSON will still be saved so the reason "
            "can be inspected."
        )


# =============================================================================
# Save
# =============================================================================

def save_companyfacts(
    ticker: str,
    companyfacts: dict,
) -> Path:
    """
    Save CompanyFacts to data/raw/{ticker}_companyfacts.json
    """

    RAW_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        RAW_DIR
        / f"{ticker}_companyfacts.json"
    )

    with output_path.open(
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            companyfacts,
            f,
            ensure_ascii=False,
            indent=2,
        )

    return output_path


# =============================================================================
# Optional local metadata update
# =============================================================================

def show_local_ticker_status(
    ticker: str,
    cik: int,
    company_name: str,
) -> None:
    """
    Show whether ticker already exists in config/tickers.csv.

    IMPORTANT:
    This function does NOT modify tickers.csv.

    The CSV is no longer required for ticker lookup.
    """

    if not TICKERS_CSV.exists():
        print(
            "\n  [info] config/tickers.csv not found."
        )
        return

    try:
        import csv

        with TICKERS_CSV.open(
            "r",
            encoding="utf-8-sig",
            newline="",
        ) as f:

            reader = csv.DictReader(f)

            existing = set()

            for row in reader:
                value = (
                    row.get("ticker")
                    or ""
                ).strip().upper()

                if value:
                    existing.add(value)

        if ticker in existing:
            print(
                f"  Local config: {ticker} already exists "
                "in config/tickers.csv"
            )

        else:
            print(
                f"  Local config: {ticker} is new."
            )

            print(
                "  Note: no manual CSV update is required "
                "for fetching CompanyFacts."
            )

    except Exception as exc:
        print(
            f"  [warning] Could not inspect "
            f"config/tickers.csv: {exc}"
        )


# =============================================================================
# Main
# =============================================================================

def main():
    if len(sys.argv) < 2:
        print(
            "Usage:\n"
            "  python scripts/01_fetch_sec_companyfacts.py TICKER\n\n"
            "Examples:\n"
            "  python scripts/01_fetch_sec_companyfacts.py AAPL\n"
            "  python scripts/01_fetch_sec_companyfacts.py NVDA\n"
            "  python scripts/01_fetch_sec_companyfacts.py AMZN"
        )

        raise SystemExit(1)

    ticker = normalize_ticker(
        sys.argv[1]
    )

    print("=" * 80)
    print(
        f"SEC CompanyFacts fetch: {ticker}"
    )
    print("=" * 80)

    # -------------------------------------------------------------------------
    # 1. SEC headers
    # -------------------------------------------------------------------------

    headers = build_headers()

    # -------------------------------------------------------------------------
    # 2. Resolve ticker -> CIK
    # -------------------------------------------------------------------------

    print(
        "\n[1/3] Resolving ticker through SEC..."
    )

    ticker_map = load_sec_ticker_map(
        headers
    )

    company_info = resolve_ticker(
        ticker,
        ticker_map,
    )

    cik = int(
        company_info["cik"]
    )

    company_name = company_info[
        "company_name"
    ]

    print(
        f"  [ok] Ticker: {ticker}"
    )

    print(
        f"  [ok] CIK: {cik:010d}"
    )

    print(
        f"  [ok] Company: {company_name}"
    )

    show_local_ticker_status(
        ticker,
        cik,
        company_name,
    )

    # -------------------------------------------------------------------------
    # 3. Fetch CompanyFacts
    # -------------------------------------------------------------------------

    print(
        "\n[2/3] Fetching SEC CompanyFacts..."
    )

    companyfacts = fetch_companyfacts(
        ticker,
        cik,
        headers,
    )

    # -------------------------------------------------------------------------
    # 4. Validate
    # -------------------------------------------------------------------------

    print(
        "\n[3/3] Validating response..."
    )

    validate_companyfacts(
        ticker,
        companyfacts,
    )

    # -------------------------------------------------------------------------
    # 5. Save
    # -------------------------------------------------------------------------

    output_path = save_companyfacts(
        ticker,
        companyfacts,
    )

    print(
        "\n" + "=" * 80
    )

    print(
        "[ok] CompanyFacts saved:"
    )

    print(
        f"    {output_path}"
    )

    print(
        "=" * 80
    )


if __name__ == "__main__":
    main()
