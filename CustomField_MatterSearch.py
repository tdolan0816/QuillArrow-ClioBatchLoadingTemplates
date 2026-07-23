#!/usr/bin/env python3
"""
find_matters_by_custom_field.py

Paginates through ALL Clio matters, inspects each matter's custom_field_values,
and collects every matter that contains a specific custom field definition ID.

Because this is a large pull (23,000+ matters x 300+ custom fields), the script:
  - Pages 200 matters at a time (Clio's max)
  - Prints live progress so you can see it working
  - Retries automatically on rate-limit (429) responses
  - Writes matched matters to a CSV when done
"""

import csv
import json
import time
from pathlib import Path

import requests

# =============================================================================
# CONFIGURATION — edit these before running
# =============================================================================

TOKEN_PATH = Path(
    r"C:\Users\Tim\OneDrive - quillarrowlaw.com\Documents\ClioTemplates_CustomFields_MassUpdate\clio_tokens.json"
)

# The custom field DEFINITION ID you are searching for
TARGET_CF_ID = 13826763

# Where to save the results
OUTPUT_CSV = Path(
    r"C:\Users\Tim\OneDrive - quillarrowlaw.com\Documents\ClioTemplates_CustomFields_MassUpdate"
    r"\matters_with_cf_7134166.csv"
)

# =============================================================================
# CONSTANTS
# =============================================================================

BASE_URL   = "https://app.clio.com/api/v4"
PAGE_LIMIT = 500          # Clio's maximum per page
RETRY_WAIT = 15           # seconds to wait after a 429 rate-limit response
MAX_RETRIES = 5           # max retries per page before giving up

FIELDS = (
    "id,"
    "display_number,"
    "description,"
    "status,"
    "custom_field_values{id,value,custom_field}"
)

# =============================================================================
# LOAD AUTH TOKEN
# =============================================================================

def load_headers(token_path: Path) -> dict:
    with token_path.open("r") as f:
        tokens = json.load(f)
    return {
        "Authorization": f"Bearer {tokens['access_token']}",
        "Content-Type":  "application/json",
    }

# =============================================================================
# FETCH ONE PAGE WITH RETRY LOGIC
# =============================================================================

def fetch_page(url: str, headers: dict) -> dict:
    for attempt in range(1, MAX_RETRIES + 1):
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            return response.json()

        if response.status_code == 429:
            print(
                f"    ⚠ Rate limited (429). Waiting {RETRY_WAIT}s "
                f"(attempt {attempt}/{MAX_RETRIES})..."
            )
            time.sleep(RETRY_WAIT)
            continue

        if response.status_code == 401:
            raise RuntimeError(
                "401 Unauthorized — your access token has expired. "
                "Refresh it and update your token JSON file."
            )

        raise RuntimeError(
            f"Unexpected {response.status_code} from Clio: {response.text}"
        )

    raise RuntimeError(
        f"Failed to fetch page after {MAX_RETRIES} retries (rate limit). "
        "Try increasing RETRY_WAIT."
    )

# =============================================================================
# MAIN
# =============================================================================

def main() -> None:
    headers = load_headers(TOKEN_PATH)

    # Build the first page URL
    first_url = (
        f"{BASE_URL}/matters"
        f"?fields={FIELDS}"
        f"&limit={PAGE_LIMIT}"
        f"&order=id(asc)"   # stable ordering so pages don't shift mid-run

    )

    matched_matters: list[dict] = []
    page_num       = 0
    total_scanned  = 0

    print("=" * 65)
    print(f"Searching for Custom Field ID: {TARGET_CF_ID}")
    print("=" * 65)

    current_url = first_url

    while current_url:
        page_num += 1
        data = fetch_page(current_url, headers)

        matters    = data.get("data", [])
        paging     = data.get("paging", {})
        meta       = data.get("meta", {})

        # Total record count is in meta on the first page
        total_records = meta.get("records", "?")

        page_matches = 0

        for matter in matters:
            total_scanned += 1
            cfvs = matter.get("custom_field_values") or []

            for cfv in cfvs:
                cf_ref = cfv.get("custom_field") or {}
                if cf_ref.get("id") == TARGET_CF_ID:
                    matched_matters.append({
                        "matter_id":      matter.get("id"),
                        "display_number": matter.get("display_number"),
                        "description":    matter.get("description"),
                        "status":         matter.get("status"),
                        "cf_value_id":    cfv.get("id"),
                        "cf_value":       cfv.get("value"),
                    })
                    page_matches += 1
                    break   # only need to match once per matter

        # Progress line
        print(
            f"  Page {page_num:>4}  |  "
            f"Scanned: {total_scanned:>6}/{total_records}  |  "
            f"Page matches: {page_matches:>3}  |  "
            f"Total found so far: {len(matched_matters):>4}"
        )

        # Advance to next page (Clio gives us the full next URL)
        next_url = paging.get("next")
        current_url = next_url if next_url else None

    # -------------------------------------------------------------------------
    print()
    print("=" * 65)
    print(f"Scan complete.")
    print(f"  Total matters scanned : {total_scanned}")
    print(f"  Matters with CF {TARGET_CF_ID}: {len(matched_matters)}")
    print("=" * 65)

    if not matched_matters:
        print("No matters found containing this custom field.")
        return

    # -------------------------------------------------------------------------
    # Write results to CSV
    # -------------------------------------------------------------------------
    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT_CSV.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=[
                "matter_id", "display_number", "description",
                "status", "cf_value_id", "cf_value",
            ],
        )
        writer.writeheader()
        writer.writerows(matched_matters)

    print(f"\nResults saved to: {OUTPUT_CSV}")
    print()

    # Print a quick preview of the first 10 matches
    preview = matched_matters[:10]
    print(f"Preview (first {len(preview)} matches):")
    print(f"  {'Matter ID':<14} {'Display #':<22} {'Status':<10}  CF Value")
    print(f"  {'-'*12} {'-'*20} {'-'*8}  {'-'*20}")
    for m in preview:
        print(
            f"  {str(m['matter_id']):<14} "
            f"{str(m['display_number']):<22} "
            f"{str(m['status']):<10}  "
            f"{m['cf_value']}"
        )

    if len(matched_matters) > 10:
        print(f"  ... and {len(matched_matters) - 10} more (see CSV for full list)")


if __name__ == "__main__":
    main()