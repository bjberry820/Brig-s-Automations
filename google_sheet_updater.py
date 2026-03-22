"""
Google Sheet Updater
====================
Fetches data from an API or web scrape and writes it to a Google Sheet.

Setup:
1. Copy .env.example to .env and fill in your values.
2. Place your Google service account JSON key at the path set in GOOGLE_SERVICE_ACCOUNT_FILE.
3. Share your Google Sheet with the service account email (found in the JSON key).
4. Install dependencies: pip install -r requirements.txt
5. Run: python google_sheet_updater.py

To get Google credentials:
- Go to https://console.cloud.google.com/
- Create a project > Enable "Google Sheets API" and "Google Drive API"
- Go to IAM & Admin > Service Accounts > Create service account
- Download the JSON key and save it (e.g. as credentials.json)
- Share your Google Sheet with the service account email
"""

import os
import json
import sys
from typing import Optional

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ---------------------------------------------------------------------------
# Google Sheets helpers
# ---------------------------------------------------------------------------

def connect_to_sheet(service_account_file: str, sheet_id: str, worksheet_name: str):
    """Authenticate and return (spreadsheet, worksheet)."""
    creds = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)
    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Worksheet '{worksheet_name}' not found — creating it.")
        worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=26)
    return spreadsheet, worksheet


def read_sheet(worksheet) -> list[list]:
    """Return all values from the worksheet as a 2D list."""
    return worksheet.get_all_values()


def append_rows(worksheet, rows: list[list]):
    """Append a list of rows to the end of the worksheet."""
    if not rows:
        print("No rows to append.")
        return
    worksheet.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"Appended {len(rows)} row(s).")


def write_rows(worksheet, rows: list[list], start_cell: str = "A1"):
    """Overwrite data starting at start_cell."""
    if not rows:
        print("No rows to write.")
        return
    worksheet.update(start_cell, rows, value_input_option="USER_ENTERED")
    print(f"Wrote {len(rows)} row(s) starting at {start_cell}.")


# ---------------------------------------------------------------------------
# Data fetching helpers
# ---------------------------------------------------------------------------

def fetch_from_api(url: str, api_key: Optional[str] = None) -> list[list]:
    """
    Fetch JSON data from a REST API and return it as rows.

    Handles two common JSON shapes:
      - A list of dicts  -> rows are [headers] + [[v1, v2, ...], ...]
      - A dict with a list value -> same as above, using the first list found
    """
    headers = {}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    data = response.json()

    # Unwrap common wrapper shapes: {"data": [...]} or {"results": [...]}
    if isinstance(data, dict):
        for key in ("data", "results", "items", "records"):
            if key in data and isinstance(data[key], list):
                data = data[key]
                break
        else:
            # Fall back: first list value in the dict
            for value in data.values():
                if isinstance(value, list):
                    data = value
                    break
            else:
                print("Warning: API returned a dict with no list — wrapping as single row.")
                data = [data]

    if not data:
        print("API returned an empty list.")
        return []

    if isinstance(data[0], dict):
        headers_row = list(data[0].keys())
        rows = [headers_row] + [[str(item.get(h, "")) for h in headers_row] for item in data]
    else:
        rows = [[str(cell) for cell in row] for row in data]

    return rows


def fetch_from_scrape(url: str, table_selector: str = "table") -> list[list]:
    """
    Scrape an HTML table from a URL and return rows as a 2D list.
    Uses the first <table> matched by table_selector.
    """
    response = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "lxml")
    table = soup.select_one(table_selector)
    if not table:
        raise ValueError(f"No element found matching selector: '{table_selector}'")

    rows = []
    for tr in table.find_all("tr"):
        cells = [cell.get_text(strip=True) for cell in tr.find_all(["th", "td"])]
        if cells:
            rows.append(cells)

    print(f"Scraped {len(rows)} row(s) from {url}")
    return rows


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    service_account_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "credentials.json")
    sheet_id = os.getenv("GOOGLE_SHEET_ID", "")
    worksheet_name = os.getenv("WORKSHEET_NAME", "Sheet1")
    mode = os.getenv("MODE", "api").lower()
    api_url = os.getenv("API_URL", "")
    api_key = os.getenv("API_KEY", "") or None
    scrape_url = os.getenv("SCRAPE_URL", "")
    scrape_selector = os.getenv("SCRAPE_TABLE_SELECTOR", "table")

    # --- Validate config ---
    if not sheet_id:
        print("Error: GOOGLE_SHEET_ID is not set in .env")
        sys.exit(1)
    if not os.path.exists(service_account_file):
        print(f"Error: Service account file not found: {service_account_file}")
        print("See the setup instructions at the top of this file.")
        sys.exit(1)

    # --- Connect ---
    print(f"Connecting to Google Sheet: {sheet_id} / {worksheet_name}")
    _, worksheet = connect_to_sheet(service_account_file, sheet_id, worksheet_name)

    # --- Read existing data (optional - useful for deduplication) ---
    existing = read_sheet(worksheet)
    print(f"Sheet currently has {len(existing)} row(s).")

    # --- Fetch new data ---
    if mode == "api":
        if not api_url:
            print("Error: API_URL is not set in .env")
            sys.exit(1)
        print(f"Fetching data from API: {api_url}")
        new_rows = fetch_from_api(api_url, api_key)
    elif mode == "scrape":
        if not scrape_url:
            print("Error: SCRAPE_URL is not set in .env")
            sys.exit(1)
        print(f"Scraping data from: {scrape_url}")
        new_rows = fetch_from_scrape(scrape_url, scrape_selector)
    else:
        print(f"Error: MODE must be 'api' or 'scrape', got '{mode}'")
        sys.exit(1)

    if not new_rows:
        print("No data fetched. Nothing to write.")
        return

    print(f"Fetched {len(new_rows)} row(s).")

    # --- Write to sheet ---
    # If the sheet is empty, write everything including headers.
    # If it already has data, append only the data rows (skip the header row).
    if not existing:
        write_rows(worksheet, new_rows)
    else:
        data_rows = new_rows[1:] if len(new_rows) > 1 else new_rows
        append_rows(worksheet, data_rows)

    print("Done.")


if __name__ == "__main__":
    main()
