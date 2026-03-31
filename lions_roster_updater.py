"""
Detroit Lions Roster — Google Sheet Updater
============================================
Writes the current Lions roster to a Google Sheet.

Columns written (A–N):
  A  Full Name
  B  Jersey #
  C  Position (abbreviation, e.g. QB)
  D  Height (e.g. 6'4")
  E  Weight (e.g. 222 lbs)
  F  Age
  G  Date of Birth (YYYY-MM-DD)
  H  College
  I  Years of Experience
  J  Headshot URL
  K  Draft Year
  L  Draft Round
  M  Draft Pick (overall)
  N  Drafted By (team)

Data sources:
  • ESPN public API  — columns A–J  (reliable, no auth needed)
  • Pro-Football-Reference — columns K–N  (best-effort; skipped if rate-limited)

Setup:
  1. Copy .env.example → .env and fill in your values.
  2. Share the Google Sheet with the service-account email.
  3. pip install -r requirements.txt
  4. python lions_roster_updater.py
"""

import os
import re
import sys
import time
from typing import Optional

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ESPN_ROSTER_URL = (
    "https://site.api.espn.com/apis/site/v2/sports/football/nfl/teams/8/roster"
)

PFR_SEARCH_URL = "https://www.pro-football-reference.com/search/search.fcgi"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

SHEET_HEADERS = [
    "Name", "Jersey #", "Position", "Height", "Weight",
    "Age", "Date of Birth", "College", "Experience (yrs)",
    "Headshot URL", "Draft Year", "Draft Round", "Draft Pick", "Drafted By",
]

# ---------------------------------------------------------------------------
# Step 1 — ESPN roster fetcher
# ---------------------------------------------------------------------------

def fetch_espn_roster() -> list[dict]:
    """
    Call the ESPN public API and return a flat list of player dicts.

    Each dict has keys:
      id, name, jersey, position, height, weight,
      age, dob, college, experience, headshot_url
    """
    print("Fetching roster from ESPN API…")
    resp = requests.get(ESPN_ROSTER_URL, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    data = resp.json()

    players = []

    # ESPN groups athletes by position group; flatten all groups.
    for group in data.get("athletes", []):
        for athlete in group.get("items", []):
            player = _parse_espn_athlete(athlete)
            if player:
                players.append(player)

    print(f"  → {len(players)} players found on ESPN roster.")
    return players


def _parse_espn_athlete(a: dict) -> Optional[dict]:
    """Extract relevant fields from a single ESPN athlete object."""
    try:
        espn_id = a.get("id", "")

        # Height: ESPN stores total inches as an integer
        height_in = a.get("height", 0)
        feet = int(height_in) // 12
        inches = int(height_in) % 12
        height_str = f"{feet}'{inches}\"" if height_in else ""

        # Weight
        weight = a.get("weight", "")
        weight_str = f"{int(weight)} lbs" if weight else ""

        # Date of birth — strip time portion if present
        dob_raw = a.get("dateOfBirth", "")
        dob = dob_raw[:10] if dob_raw else ""  # "YYYY-MM-DD"

        # College name
        college_obj = a.get("college") or {}
        college = college_obj.get("name", "") or college_obj.get("shortDisplayName", "")

        # Position abbreviation
        pos_obj = a.get("position") or {}
        position = pos_obj.get("abbreviation", "")

        # Experience
        exp_obj = a.get("experience") or {}
        experience = exp_obj.get("years", "")

        # Headshot
        headshot_obj = a.get("headshot") or {}
        headshot_url = headshot_obj.get("href", "")
        if not headshot_url and espn_id:
            headshot_url = (
                f"https://a.espncdn.com/i/headshots/nfl/players/full/{espn_id}.png"
            )

        return {
            "id": espn_id,
            "name": a.get("fullName", a.get("displayName", "")),
            "jersey": a.get("jersey", ""),
            "position": position,
            "height": height_str,
            "weight": weight_str,
            "age": a.get("age", ""),
            "dob": dob,
            "college": college,
            "experience": experience,
            "headshot_url": headshot_url,
        }
    except Exception as exc:
        print(f"  Warning: could not parse athlete {a.get('fullName', '?')}: {exc}")
        return None


# ---------------------------------------------------------------------------
# Step 2 — PFR draft info fetcher (best-effort, rate-limited)
# ---------------------------------------------------------------------------

def fetch_pfr_draft_info(players: list[dict], delay: float = 1.5) -> dict[str, dict]:
    """
    Look up each player on Pro-Football-Reference and extract draft details.

    Returns a dict keyed by player name:
      { "Jared Goff": { "draft_year": "2016", "draft_round": "1",
                        "draft_pick": "1", "drafted_by": "Los Angeles Rams" } }

    Skips players gracefully if PFR returns an error or rate-limits us.
    """
    print(f"\nFetching draft info from PFR for {len(players)} players…")
    print("  (1 request per player — this takes a while)")

    results = {}
    for i, player in enumerate(players, 1):
        name = player["name"]
        print(f"  [{i}/{len(players)}] {name}", end="", flush=True)
        info = _pfr_search_player(name)
        results[name] = info
        status = "✓" if any(info.values()) else "–"
        print(f"  {status}")
        time.sleep(delay)

    found = sum(1 for v in results.values() if any(v.values()))
    print(f"  → Draft info found for {found}/{len(players)} players.")
    return results


def _pfr_search_player(name: str) -> dict:
    """Search PFR for a player by name and return draft fields."""
    empty = {"draft_year": "", "draft_round": "", "draft_pick": "", "drafted_by": ""}
    try:
        resp = requests.get(
            PFR_SEARCH_URL,
            params={"search": name},
            headers=HEADERS,
            timeout=15,
            allow_redirects=True,
        )
        if resp.status_code == 429:
            print("  [rate-limited — stopping PFR fetching]", end="")
            return empty
        if resp.status_code != 200:
            return empty

        # PFR redirects directly to the player page when there's an exact match.
        # Otherwise it shows a search results page — grab the first link.
        if "search" in resp.url:
            player_url = _pfr_first_result_url(resp.text)
            if not player_url:
                return empty
            time.sleep(0.5)
            resp = requests.get(player_url, headers=HEADERS, timeout=15)
            if resp.status_code != 200:
                return empty

        return _pfr_parse_draft(resp.text)

    except Exception:
        return empty


def _pfr_first_result_url(html: str) -> Optional[str]:
    """Return the href of the first player result on a PFR search page."""
    soup = BeautifulSoup(html, "lxml")
    link = soup.select_one("div.search-item-name a")
    if link and link.get("href"):
        href = link["href"]
        if href.startswith("/"):
            return "https://www.pro-football-reference.com" + href
        return href
    return None


def _pfr_parse_draft(html: str) -> dict:
    """Parse draft details from a PFR player page."""
    empty = {"draft_year": "", "draft_round": "", "draft_pick": "", "drafted_by": ""}
    soup = BeautifulSoup(html, "lxml")

    # Draft info lives in a <p> tag inside #meta that contains "Draft:"
    meta = soup.select_one("#meta")
    if not meta:
        return empty

    for p in meta.find_all("p"):
        text = p.get_text(" ", strip=True)
        if text.startswith("Draft:"):
            return _parse_draft_text(text)

    return empty


def _parse_draft_text(text: str) -> dict:
    """
    Parse a string like:
      "Draft: Los Angeles Rams in the 1st round (1st pick, 1st overall) of the 2016 NFL Draft."
    """
    result = {"draft_year": "", "draft_round": "", "draft_pick": "", "drafted_by": ""}

    # Year
    year_m = re.search(r"\b(19|20)\d{2}\b", text)
    if year_m:
        result["draft_year"] = year_m.group()

    # Round (1st / 2nd / 3rd … or ordinal)
    round_m = re.search(r"(\d+)(?:st|nd|rd|th)\s+round", text, re.IGNORECASE)
    if round_m:
        result["draft_round"] = round_m.group(1)

    # Overall pick — "123rd overall" or "123rd pick, 123rd overall"
    overall_m = re.search(r"(\d+)(?:st|nd|rd|th)\s+overall", text, re.IGNORECASE)
    if overall_m:
        result["draft_pick"] = overall_m.group(1)

    # Team — text between "Draft:" and "in the"
    team_m = re.search(r"Draft:\s+(.+?)\s+in the", text)
    if team_m:
        result["drafted_by"] = team_m.group(1).strip()

    return result


# ---------------------------------------------------------------------------
# Step 3 — Assemble rows
# ---------------------------------------------------------------------------

def build_rows(players: list[dict], draft_info: dict[str, dict]) -> list[list]:
    """Combine ESPN + PFR data into sheet rows."""
    rows = [SHEET_HEADERS]
    for p in players:
        di = draft_info.get(p["name"], {})
        row = [
            p["name"],
            p["jersey"],
            p["position"],
            p["height"],
            p["weight"],
            p["age"],
            p["dob"],
            p["college"],
            p["experience"],
            p["headshot_url"],
            di.get("draft_year", ""),
            di.get("draft_round", ""),
            di.get("draft_pick", ""),
            di.get("drafted_by", ""),
        ]
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Step 4 — Google Sheets helpers
# ---------------------------------------------------------------------------

def connect_sheet(service_account_file: str, sheet_id: str, worksheet_name: str):
    """Return an authenticated gspread Worksheet."""
    creds = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)
    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Worksheet '{worksheet_name}' not found — creating it.")
        worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows=200, cols=20)
    return worksheet


def write_roster(worksheet, rows: list[list]):
    """Clear the sheet and write rows starting at A1."""
    print(f"\nWriting {len(rows)} row(s) to sheet (including header)…")
    worksheet.clear()
    worksheet.update("A1", rows, value_input_option="USER_ENTERED")
    print("Done.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    service_account_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "credentials.json")
    sheet_id = os.getenv("GOOGLE_SHEET_ID", "")
    worksheet_name = os.getenv("WORKSHEET_NAME", "Sheet1")
    skip_pfr = os.getenv("SKIP_PFR", "false").lower() in ("1", "true", "yes")

    if not sheet_id:
        print("Error: GOOGLE_SHEET_ID is not set in .env")
        sys.exit(1)
    if not os.path.exists(service_account_file):
        print(f"Error: Service account file not found: {service_account_file}")
        sys.exit(1)

    # --- Fetch data ---
    players = fetch_espn_roster()
    if not players:
        print("No players returned from ESPN. Aborting.")
        sys.exit(1)

    if skip_pfr:
        print("\nSKIP_PFR=true — skipping Pro-Football-Reference lookup.")
        draft_info = {}
    else:
        draft_info = fetch_pfr_draft_info(players)

    rows = build_rows(players, draft_info)

    # --- Write to sheet ---
    print(f"\nConnecting to Google Sheet: {sheet_id} / {worksheet_name}")
    worksheet = connect_sheet(service_account_file, sheet_id, worksheet_name)
    write_roster(worksheet, rows)


if __name__ == "__main__":
    main()
