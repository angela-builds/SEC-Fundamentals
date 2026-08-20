import csv
import json
import os
import time
from pathlib import Path
from urllib.request import Request, urlopen


BASE_DIR = Path(__file__).resolve().parents[1]
TICKERS_PATH = BASE_DIR / "config" / "tickers.csv"
RAW_DIR = BASE_DIR / "data" / "raw"

SEC_COMPANYFACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
SEC_COMPANY_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"

# SEC requires a descriptive User-Agent with a real contact email.
# Read from an environment variable, not hardcoded -- this file is meant to
# be committed to git/GitHub, and a hardcoded personal email would get
# published along with the code.
#
# Set it before running, e.g. (Windows PowerShell):
#   $env:SEC_CONTACT_EMAIL = "your_email@example.com"
# or put it in a local .env file (see .env.example) and load it however
# your shell/notebook does that. Falls back to a placeholder if unset, so
# the script still runs, but SEC may rate-limit/block a non-identifying UA.
SEC_CONTACT_EMAIL = os.environ.get("SEC_CONTACT_EMAIL", "REPLACE_ME@example.com")
if SEC_CONTACT_EMAIL == "REPLACE_ME@example.com":
    print(
        "[warning] SEC_CONTACT_EMAIL environment variable is not set. "
        "Using a placeholder User-Agent -- set SEC_CONTACT_EMAIL to your "
        "real email before relying on this for real use."
    )

HEADERS = {
    "User-Agent": f"Angela sec-fundamental-tool {SEC_CONTACT_EMAIL}",
    "Host": "data.sec.gov",
}

SEC_FILES_HEADERS = {
    "User-Agent": f"Angela sec-fundamental-tool {SEC_CONTACT_EMAIL}",
    "Host": "www.sec.gov",
}


def normalize_cik(cik):
    cik = str(cik).strip()

    if not cik:
        return ""

    return cik.zfill(10)


def load_ticker_map():
    ticker_map = {}

    with open(TICKERS_PATH, mode="r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)

        for row in reader:
            ticker = row["ticker"].strip().upper()
            cik = normalize_cik(row.get("cik", ""))
            company_name = row.get("company_name", "").strip()

            if not ticker:
                continue

            ticker_map[ticker] = {
                "cik": cik,
                "company_name": company_name,
            }

    return ticker_map


def load_sec_ticker_lookup():
    request = Request(SEC_COMPANY_TICKERS_URL, headers=SEC_FILES_HEADERS)

    print("Loading SEC ticker to CIK lookup table...")
    print(SEC_COMPANY_TICKERS_URL)

    with urlopen(request, timeout=30) as response:
        raw_text = response.read().decode("utf-8")
        data = json.loads(raw_text)

    lookup = {}

    for item in data.values():
        ticker = item["ticker"].strip().upper()
        cik = normalize_cik(item["cik_str"])
        title = item["title"].strip()

        lookup[ticker] = {
            "cik": cik,
            "company_name": title,
        }

    return lookup


def resolve_company(ticker, ticker_map):
    if ticker not in ticker_map:
        available = ", ".join(ticker_map.keys())
        raise ValueError(f"Ticker not found in config/tickers.csv. Available: {available}")

    company = ticker_map[ticker]

    if company["cik"]:
        return company

    sec_lookup = load_sec_ticker_lookup()

    if ticker not in sec_lookup:
        raise ValueError(f"CIK not found from SEC company_tickers.json for ticker: {ticker}")

    resolved = sec_lookup[ticker]

    if not company["company_name"]:
        company["company_name"] = resolved["company_name"]

    company["cik"] = resolved["cik"]

    print(f"Resolved {ticker}: CIK {company['cik']} ({company['company_name']})")

    return company


def download_companyfacts(ticker, cik):
    url = SEC_COMPANYFACTS_URL.format(cik=cik)
    request = Request(url, headers=HEADERS)

    print(f"Downloading {ticker} companyfacts from SEC...")
    print(url)

    with urlopen(request, timeout=30) as response:
        raw_text = response.read().decode("utf-8")
        data = json.loads(raw_text)

    RAW_DIR.mkdir(parents=True, exist_ok=True)

    output_path = RAW_DIR / f"{ticker}_companyfacts.json"

    with open(output_path, mode="w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)

    print(f"Saved to: {output_path}")


def main():
    ticker_map = load_ticker_map()

    import sys as _sys
    ticker = _sys.argv[1].strip().upper() if len(_sys.argv) > 1 else input("Enter ticker, for example AAPL or PG: ").strip().upper()

    try:
        company = resolve_company(ticker, ticker_map)
    except ValueError as error:
        print(error)
        return

    download_companyfacts(ticker, company["cik"])

    # Be polite to SEC servers.
    time.sleep(0.2)


if __name__ == "__main__":
    main()
