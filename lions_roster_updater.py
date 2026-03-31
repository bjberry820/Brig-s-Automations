"""
Detroit Lions Roster — Info Tab Updater
========================================
Populates columns A–V on the 'Info' worksheet.

Column map:
  A  Player name (first + last)      — detroitlions.com roster page
  B  Last name                        — derived from A
  C  Jersey number                    — detroitlions.com
  D  PFR profile URL                  — pro-football-reference.com (fallback: CFR, ESPN, college)
  E  Headshot image URL               — ESPN (fallback: PFR, college)
  F  Position abbreviation            — Lions page / PFR (specific ruleset — see NOTE below)
  G  Height — feet (integer floor)    — Lions page
  H  Height — inches remainder        — Lions page
  I  Weight                           — Lions page
  J  College / University             — Colleges primary-key tab (comma-sep, final school first)
  K  Birthdate                        — PFR
  L  Year entered NFL                 — PFR
  M  Drafted by (team)                — PFR
  N  Draft round                      — PFR
  O  Draft pick (overall)             — PFR
  P  Draft year                       — PFR
  Q  Year joined Lions                — PFR
  R  Teams played for                 — PFR
  S  Acquired                         — how Lions got player (team abbr in parens)
  T  Roster status                    — formula vs Depth Chart tab
  U  Money Sheet 1                    — formula vs New Money Sheet tab
  V  Money Sheet 2                    — formula vs 2023 Money tab

Position hierarchy (column F):
  Offense — QB · RB · FB · WR · TE
  OL group  OL > IOL > OG > {RG, LG} · OL > OT > {RT, LT} · C
  DL group  DL > {DT, DE}
  LB group  LB > {MLB, OLB}
  DB group  DB > {CB, S}
  Special   P · K · LS
  Rule: when in doubt, use the broader position.

Usage:
  # Fill one row (reads name from A60, writes D60:V60):
  python lions_roster_updater.py --row 60

  # Fill all rows from the Lions roster (writes A2:V{last}):
  python lions_roster_updater.py

  # Skip PFR lookups (faster, columns D/K-S will be blank):
  python lions_roster_updater.py --skip-pfr

Setup:
  1. pip install -r requirements.txt
  2. Copy .env.example to .env and fill in GOOGLE_SERVICE_ACCOUNT_FILE
  3. Share the sheet with the service-account email in your credentials JSON
"""

import argparse
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
# Config
# ---------------------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "15lIeWdwRmqtu7hqIolmRyO42X__d6iPwTZu_GKf2Khc")
INFO_TAB = os.getenv("WORKSHEET_NAME", "Info")
DATA_START_ROW = 2  # Row 1 = headers

LIONS_ROSTER_URL = "https://www.detroitlions.com/team/players-roster/"
ESPN_ROSTER_URL = "https://site.api.espn.com/apis/site/v2/sports/football/nfl/teams/8/roster"
PFR_BASE = "https://www.pro-football-reference.com"
PFR_SEARCH_URL = f"{PFR_BASE}/search/search.fcgi"

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
}

# ---------------------------------------------------------------------------
# Column formulas  (row number substituted at build time)
# ---------------------------------------------------------------------------

def _t_formula(row: int) -> str:
    r = row
    return (
        f"=IF("
        f"COUNTIF('Depth Chart'!$B$8:$AD$9,A{r})"
        f"+COUNTIF('Depth Chart'!$D$6:$D$7,A{r}),"
        f"\"Practice Squad\","
        f"IF(COUNTIF('Depth Chart'!$B$2:$AD$7,A{r}),\"Active\","
        f"IF(COUNTIF('Depth Chart'!$B$10:$AD$12,A{r}),\"Reserve\","
        f"\"Missing\")))"
    )


def _u_formula(row: int) -> str:
    return (
        f"=IF(ISERROR(MATCH(A{row},'New Money Sheet'!$B$2:$B$233,0)),"
        f"\"Missing\",\"Present\")"
    )


def _v_formula(row: int) -> str:
    return (
        f"=IF(ISERROR(MATCH(A{row},'2023 Money'!$A$2:$A$1153,0)),"
        f"\"Missing\",\"Present\")"
    )


# ---------------------------------------------------------------------------
# Position resolver
# ---------------------------------------------------------------------------
#
# Hierarchy (broader beats narrower; when in doubt go broader):
#
#   OL  ← IOL ← OG ← { RG, LG }
#   OL  ← OT  ← { RT, LT }
#   C
#   DL  ← { DT, DE }
#   LB  ← { MLB, OLB }
#   DB  ← { CB, S }
#
# Raw inputs from Lions page / PFR / ESPN are normalised to these codes.
# ---------------------------------------------------------------------------

# Exact matches (case-insensitive) → our code
_POS_EXACT: dict[str, str] = {
    # Skill
    "qb": "QB", "quarterback": "QB",
    "rb": "RB", "hb": "RB", "halfback": "RB", "runningback": "RB", "running back": "RB",
    "fb": "FB", "fullback": "FB",
    "wr": "WR", "wide receiver": "WR", "wideout": "WR",
    "te": "TE", "tight end": "TE",
    # OL — specific
    "rg": "RG", "right guard": "RG",
    "lg": "LG", "left guard": "LG",
    "rt": "RT", "right tackle": "RT",
    "lt": "LT", "left tackle": "LT",
    "c":  "C",  "center": "C",
    # OL — broad
    "og": "OG", "g": "OG", "guard": "OG",
    "ot": "OT", "t": "OT", "tackle": "OT",
    "iol": "IOL",
    "ol": "OL", "o-line": "OL", "offensive line": "OL", "offensive lineman": "OL",
    # DL
    "de": "DE", "defensive end": "DE",
    "dt": "DT", "defensive tackle": "DT",
    "nt": "DT", "nose tackle": "DT", "ng": "DT", "nose guard": "DT",
    "dl": "DL", "defensive line": "DL", "defensive lineman": "DL",
    # LB
    "mlb": "MLB", "ilb": "MLB", "inside linebacker": "MLB", "middle linebacker": "MLB",
    "olb": "OLB", "outside linebacker": "OLB",
    "lolb": "OLB", "rolb": "OLB",
    "lb": "LB", "linebacker": "LB",
    # DB
    "cb": "CB", "cornerback": "CB", "corner": "CB",
    "s":  "S",  "safety": "S",
    "ss": "S",  "strong safety": "S",
    "fs": "S",  "free safety": "S",
    "db": "DB", "defensive back": "DB",
    # Special teams
    "p": "P", "punter": "P",
    "k": "K", "pk": "K", "kicker": "K", "placekicker": "K",
    "ls": "LS", "long snapper": "LS",
}

# When a raw string contains MULTIPLE position codes, resolve via these
# combination → broader mapping rules.
_OL_CODES  = {"RG", "LG", "OG", "RT", "LT", "OT", "C", "IOL", "OL"}
_DL_CODES  = {"DE", "DT"}
_LB_CODES  = {"MLB", "OLB", "LB"}
_DB_CODES  = {"CB", "S", "DB"}


def resolve_position(raw_lions: str, raw_pfr: str = "") -> str:
    """
    Return the standardised position code for column F.

    Steps:
      1. Try exact lookup on Lions raw position.
      2. Try exact lookup on PFR raw position.
      3. If either source contains a slash/hyphen combination, resolve via
         the broader-position rules.
      4. Default to the broader position when ambiguous.
    """
    lions_code = _pos_lookup(raw_lions)
    pfr_code   = _pos_lookup(raw_pfr)

    # If both agree (or one is empty), return the non-empty one
    if lions_code == pfr_code:
        return lions_code
    if not lions_code:
        return pfr_code
    if not pfr_code:
        return lions_code

    # They differ — apply the "when in doubt go broader" rule
    return _broader(lions_code, pfr_code)


def _pos_lookup(raw: str) -> str:
    """Normalise a single raw position string → our code."""
    if not raw:
        return ""

    key = raw.strip().lower()

    # Direct hit
    if key in _POS_EXACT:
        return _POS_EXACT[key]

    # Try stripping punctuation / spaces
    key_clean = re.sub(r"[\s\-/]", "", key)
    if key_clean in _POS_EXACT:
        return _POS_EXACT[key_clean]

    # Slash or hyphen combo, e.g. "G/C", "DE/DT", "ILB/OLB", "G/T"
    parts_raw = re.split(r"[/\-]", raw.strip())
    if len(parts_raw) > 1:
        codes = [_pos_lookup(p.strip()) for p in parts_raw if p.strip()]
        codes = [c for c in codes if c]
        if codes:
            result = codes[0]
            for c in codes[1:]:
                result = _broader(result, c)
            return result

    # Partial match — walk the exact table looking for a substring
    for token, code in _POS_EXACT.items():
        if token in key:
            return code

    return ""  # unknown — leave blank, human can fill


def _broader(a: str, b: str) -> str:
    """Return the broader of two position codes per the hierarchy rules."""
    pair = frozenset({a, b})

    # ---- OL group ----
    if pair <= _OL_CODES:
        # If any tackle involved → OL beats IOL/OG
        if "OT" in pair or "RT" in pair or "LT" in pair:
            if "OG" in pair or "RG" in pair or "LG" in pair or "IOL" in pair or "C" in pair:
                return "OL"
            # Both are tackle-side
            if pair == {"RT", "LT"}:
                return "OT"
            return "OT" if "OT" in pair else "OL"
        # Interior only (G + C → IOL)
        if ("OG" in pair or "RG" in pair or "LG" in pair) and "C" in pair:
            return "IOL"
        if ("OG" in pair or "IOL" in pair) and b in {"RG", "LG"}:
            return "OG" if "OG" in pair else "IOL"
        if pair == {"RG", "LG"}:
            return "OG"
        # Default OL group: return the one already in the pair that's broadest
        for broad in ("OL", "IOL", "OT", "OG", "C", "RT", "LT", "RG", "LG"):
            if broad in pair:
                return broad

    # ---- DL group ----
    if pair <= (_DL_CODES | {"DL"}):
        return "DL"

    # ---- LB group ----
    if pair <= _LB_CODES:
        if "LB" in pair:
            return "LB"
        if pair == {"MLB", "OLB"}:
            return "LB"
        return a  # shouldn't reach here

    # ---- DB group ----
    if pair <= _DB_CODES:
        if "DB" in pair:
            return "DB"
        if pair == {"CB", "S"}:
            return "DB"
        return a

    # Cross-group: genuinely ambiguous — return first input (Lions is primary)
    return a


# ---------------------------------------------------------------------------
# Step 1 — Lions roster scraper (primary: detroitlions.com, fallback: ESPN)
# ---------------------------------------------------------------------------

def fetch_roster() -> list[dict]:
    """
    Return player list with keys:
      name, last_name, number, position, ht_ft, ht_in, weight, espn_id, espn_headshot
    """
    players = _scrape_lions_page()
    if players:
        print(f"  Lions page: {len(players)} players")
        # Supplement with ESPN IDs / headshots
        _merge_espn_ids(players)
        return players

    print("  Lions page unavailable — falling back to ESPN API")
    players = _fetch_espn_roster()
    print(f"  ESPN API: {len(players)} players")
    return players


def _scrape_lions_page() -> list[dict]:
    """Scrape player rows from detroitlions.com."""
    try:
        resp = requests.get(LIONS_ROSTER_URL, headers=BROWSER_HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
        players = []

        # Try standard NFL table layout first
        rows = soup.select("table tbody tr")
        if not rows:
            # Try card/grid layout used by some NFL sites
            rows = soup.select("[class*='roster'] [class*='player-row'], .d3-o-table__body tr")

        for row in rows:
            p = _parse_lions_row(row)
            if p:
                players.append(p)

        return players
    except Exception as e:
        print(f"  Lions page error: {e}")
        return []


def _parse_lions_row(row) -> Optional[dict]:
    cells = row.find_all(["td", "th"])
    try:
        if len(cells) >= 4:
            # Standard table: # | Name | Pos | Ht | Wt | ...
            number = cells[0].get_text(strip=True).lstrip("#")
            name   = cells[1].get_text(strip=True)
            pos    = cells[2].get_text(strip=True) if len(cells) > 2 else ""
            ht_str = cells[3].get_text(strip=True) if len(cells) > 3 else ""
            wt_str = cells[4].get_text(strip=True) if len(cells) > 4 else ""
        else:
            # Card layout — look by class hints
            name_el = row.select_one("[class*='name']")
            num_el  = row.select_one("[class*='number'], [class*='jersey']")
            pos_el  = row.select_one("[class*='position']")
            ht_el   = row.select_one("[class*='height']")
            wt_el   = row.select_one("[class*='weight']")
            if not name_el:
                return None
            name   = name_el.get_text(strip=True)
            number = num_el.get_text(strip=True).lstrip("#") if num_el else ""
            pos    = pos_el.get_text(strip=True) if pos_el else ""
            ht_str = ht_el.get_text(strip=True) if ht_el else ""
            wt_str = wt_el.get_text(strip=True) if wt_el else ""

        if not name or not re.search(r"[A-Za-z]", name):
            return None

        ht_ft, ht_in = _parse_height(ht_str)
        return {
            "name":         name,
            "last_name":    name.split()[-1],
            "number":       number,
            "position":     pos,
            "ht_ft":        ht_ft,
            "ht_in":        ht_in,
            "weight":       _parse_weight(wt_str),
            "espn_id":      "",
            "espn_headshot": "",
        }
    except Exception:
        return None


def _fetch_espn_roster() -> list[dict]:
    try:
        resp = requests.get(ESPN_ROSTER_URL, headers=BROWSER_HEADERS, timeout=20)
        resp.raise_for_status()
        players = []
        for group in resp.json().get("athletes", []):
            for athlete in group.get("items", []):
                p = _parse_espn_athlete(athlete)
                if p:
                    players.append(p)
        return players
    except Exception as e:
        print(f"  ESPN API error: {e}")
        return []


def _parse_espn_athlete(a: dict) -> Optional[dict]:
    try:
        h_in = int(a.get("height") or 0)
        w_raw = a.get("weight") or ""
        name  = a.get("fullName") or a.get("displayName") or ""
        pos   = (a.get("position") or {}).get("abbreviation", "")
        hs    = (a.get("headshot") or {}).get("href", "")
        eid   = a.get("id", "")
        return {
            "name":          name,
            "last_name":     name.split()[-1] if name else "",
            "number":        a.get("jersey", ""),
            "position":      pos,
            "ht_ft":         str(h_in // 12) if h_in else "",
            "ht_in":         str(h_in % 12)  if h_in else "",
            "weight":        str(int(w_raw)) if w_raw else "",
            "espn_id":       eid,
            "espn_headshot": hs or (f"https://a.espncdn.com/i/headshots/nfl/players/full/{eid}.png" if eid else ""),
        }
    except Exception:
        return None


def _merge_espn_ids(players: list[dict]):
    """Add ESPN IDs and headshots to players fetched from the Lions page."""
    try:
        espn = _fetch_espn_roster()
        by_name = {p["name"].lower(): p for p in espn}
        for p in players:
            match = by_name.get(p["name"].lower())
            if match:
                p["espn_id"]      = match["espn_id"]
                p["espn_headshot"] = match["espn_headshot"]
    except Exception:
        pass


def _parse_height(s: str) -> tuple[str, str]:
    """'6-2', '6\'2"', '74' (total inches) → ('6', '2')"""
    s = s.strip()
    m = re.search(r"(\d)\s*['\-]\s*(\d+)", s)
    if m:
        return m.group(1), m.group(2)
    m2 = re.match(r"^(\d{2})$", s)
    if m2:
        t = int(m2.group(1))
        return str(t // 12), str(t % 12)
    return "", ""


def _parse_weight(s: str) -> str:
    m = re.search(r"(\d{2,3})", s)
    return m.group(1) if m else ""


# ---------------------------------------------------------------------------
# Step 2 — PFR lookup  (rate-limited, cached)
# ---------------------------------------------------------------------------

_pfr_cache: dict[str, dict] = {}

_PFR_EMPTY = {
    "pfr_url": "", "headshot_pfr": "", "pfr_position": "",
    "dob": "", "nfl_year": "",
    "draft_team": "", "draft_round": "", "draft_pick": "", "draft_year": "",
    "lions_year": "", "teams": "", "acquired": "",
}


def fetch_pfr(name: str, delay: float = 1.5) -> dict:
    """Look up a player on PFR, return profile dict. Cached."""
    if name in _pfr_cache:
        return _pfr_cache[name]
    time.sleep(delay)
    result = _pfr_lookup(name)
    _pfr_cache[name] = result
    return result


def _pfr_lookup(name: str) -> dict:
    try:
        resp = requests.get(
            PFR_SEARCH_URL,
            params={"search": name},
            headers=BROWSER_HEADERS,
            timeout=15,
            allow_redirects=True,
        )
        if resp.status_code == 429:
            print("    [PFR rate-limited — stopping PFR requests]")
            return dict(_PFR_EMPTY)
        if resp.status_code != 200:
            return dict(_PFR_EMPTY)

        # PFR redirects directly to the player page on an exact match
        if "/players/" in resp.url and "search" not in resp.url:
            player_url = resp.url
            html = resp.text
        else:
            player_url = _pfr_first_result(resp.text, name)
            if not player_url:
                return dict(_PFR_EMPTY)
            time.sleep(0.5)
            r2 = requests.get(player_url, headers=BROWSER_HEADERS, timeout=15)
            if r2.status_code != 200:
                return dict(_PFR_EMPTY)
            html = r2.text

        return _pfr_parse_page(html, player_url)

    except Exception as e:
        print(f"    PFR error ({name}): {e}")
        return dict(_PFR_EMPTY)


def _pfr_first_result(html: str, name: str) -> Optional[str]:
    soup = BeautifulSoup(html, "lxml")
    parts = name.lower().split()
    # Prefer result where all name parts appear in the link text
    for item in soup.select("div.search-item-name a"):
        text = item.get_text(strip=True).lower()
        href = item.get("href", "")
        if all(p in text for p in parts):
            return (PFR_BASE + href) if href.startswith("/") else href
    # Fall back to first player result
    first = soup.select_one("div.search-item-name a")
    if first:
        href = first.get("href", "")
        return (PFR_BASE + href) if href.startswith("/") else href
    return None


def _pfr_parse_page(html: str, url: str) -> dict:
    result = dict(_PFR_EMPTY)
    result["pfr_url"] = url

    soup = BeautifulSoup(html, "lxml")
    meta = soup.select_one("#meta")
    if not meta:
        return result

    # Headshot
    img = meta.select_one("img")
    if img and img.get("src"):
        result["headshot_pfr"] = img["src"]

    for p in meta.find_all("p"):
        text = p.get_text(" ", strip=True)

        if text.startswith("Position:"):
            result["pfr_position"] = text.replace("Position:", "").strip().split()[0]

        if "Born:" in text:
            # Try structured element first
            birth_el = p.select_one("[data-birth]")
            if birth_el:
                result["dob"] = birth_el["data-birth"]
            else:
                m = re.search(
                    r"(January|February|March|April|May|June|July|August"
                    r"|September|October|November|December)\s+\d+,\s+\d{4}",
                    text,
                )
                if m:
                    result["dob"] = m.group()

        if text.startswith("Draft:"):
            result.update(_parse_draft_text(text))

    result["nfl_year"]  = _pfr_earliest_year(soup)
    result["teams"]     = _pfr_teams_text(meta)
    result.update(_pfr_lions_info(soup))

    return result


def _parse_draft_text(text: str) -> dict:
    out = {"draft_team": "", "draft_round": "", "draft_pick": "", "draft_year": ""}

    m = re.search(r"\b(19|20)\d{2}\b", text)
    if m:
        out["draft_year"] = m.group()

    m = re.search(r"(\d+)(?:st|nd|rd|th)\s+round", text, re.IGNORECASE)
    if m:
        out["draft_round"] = m.group(1)

    m = re.search(r"(\d+)(?:st|nd|rd|th)\s+overall", text, re.IGNORECASE)
    if m:
        out["draft_pick"] = m.group(1)

    m = re.search(r"Draft:\s+(.+?)\s+in\s+the\s+\d", text)
    if m:
        out["draft_team"] = m.group(1).strip()

    return out


def _pfr_earliest_year(soup: BeautifulSoup) -> str:
    years = []
    for el in soup.select("th[data-stat='year_id'], td[data-stat='year_id']"):
        t = el.get_text(strip=True)
        if re.match(r"^(19|20)\d{2}$", t):
            years.append(int(t))
    return str(min(years)) if years else ""


def _pfr_teams_text(meta) -> str:
    for p in meta.find_all("p"):
        text = p.get_text(" ", strip=True)
        if re.match(r"Teams?:", text):
            cleaned = re.sub(r"^Teams?:\s*", "", text)
            cleaned = re.sub(r"\(\d{4}(?:[–\-]\d{2,4})?\)", "", cleaned)
            return re.sub(r"\s+", " ", cleaned).strip().strip(",")
    return ""


def _pfr_lions_info(soup: BeautifulSoup) -> dict:
    """Extract year joined Lions and how acquired from transactions section."""
    out = {"lions_year": "", "acquired": ""}
    section = soup.select_one("#transactions")
    if not section:
        return out
    for li in section.find_all("li"):
        text = li.get_text(" ", strip=True)
        if "Detroit" not in text and "Lions" not in text:
            continue
        m = re.search(r"\b(20\d{2})\b", text)
        if m:
            out["lions_year"] = m.group()
        tl = text.lower()
        if "draft" in tl:
            out["acquired"] = "Drafted"
        elif "trade" in tl:
            out["acquired"] = "Trade"
        elif "sign" in tl:
            out["acquired"] = "Signed"
        elif "waiver" in tl:
            out["acquired"] = "Waivers"
        break
    return out


# ---------------------------------------------------------------------------
# Step 3 — College lookup from sheet
# ---------------------------------------------------------------------------

def load_college_map(spreadsheet) -> dict[str, str]:
    """
    Read the Colleges tab. Expects columns: Player Name | College(s)
    Returns { player_name_lower: college_string }
    """
    for tab_name in ("Colleges", "College", "colleges"):
        try:
            ws = spreadsheet.worksheet(tab_name)
            rows = ws.get_all_values()
            mapping = {}
            for row in rows[1:]:
                if len(row) >= 2 and row[0].strip():
                    mapping[row[0].strip().lower()] = row[1].strip()
            print(f"  College tab '{tab_name}': {len(mapping)} entries")
            return mapping
        except gspread.exceptions.WorksheetNotFound:
            continue
    print("  No Colleges tab found — column J will be empty")
    return {}


# ---------------------------------------------------------------------------
# Step 4 — Assemble one row
# ---------------------------------------------------------------------------

def build_row(player: dict, pfr: dict, row_num: int, college: str) -> list:
    """
    Returns all 22 values for columns A–V.
    Slice [3:] to get D–V only.
    """
    headshot = player.get("espn_headshot") or pfr.get("headshot_pfr") or ""

    return [
        player.get("name", ""),           # A
        player.get("last_name", ""),       # B
        player.get("number", ""),          # C
        pfr.get("pfr_url", ""),            # D
        headshot,                          # E
        resolve_position(                  # F
            player.get("position", ""),
            pfr.get("pfr_position", ""),
        ),
        player.get("ht_ft", ""),           # G
        player.get("ht_in", ""),           # H
        player.get("weight", ""),          # I
        college,                           # J
        pfr.get("dob", ""),                # K
        pfr.get("nfl_year", ""),           # L
        pfr.get("draft_team", ""),         # M
        pfr.get("draft_round", ""),        # N
        pfr.get("draft_pick", ""),         # O
        pfr.get("draft_year", ""),         # P
        pfr.get("lions_year", ""),         # Q
        pfr.get("teams", ""),              # R
        pfr.get("acquired", ""),           # S
        _t_formula(row_num),               # T
        _u_formula(row_num),               # U
        _v_formula(row_num),               # V
    ]


# ---------------------------------------------------------------------------
# Step 5 — Google Sheets connection + write helpers
# ---------------------------------------------------------------------------

def connect_sheet(service_account_file: str, sheet_id: str):
    creds = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(sheet_id)


def get_or_create_ws(spreadsheet, name: str):
    try:
        return spreadsheet.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Worksheet '{name}' not found — creating it.")
        return spreadsheet.add_worksheet(title=name, rows=300, cols=22)


def write_single_row_d_to_v(ws, row_values_a_to_v: list, row_num: int):
    """Write only columns D–V for one row (leaves A–C untouched)."""
    d_to_v = row_values_a_to_v[3:]   # indices 3–21
    ws.update(f"D{row_num}:V{row_num}", [d_to_v], value_input_option="USER_ENTERED")


def write_all_rows_a_to_v(ws, all_rows: list[list], start_row: int):
    """Bulk-write A–V for all player rows."""
    if not all_rows:
        return
    end_row = start_row + len(all_rows) - 1
    ws.update(f"A{start_row}:V{end_row}", all_rows, value_input_option="USER_ENTERED")
    print(f"  Wrote rows {start_row}–{end_row} (A:V)")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Update Detroit Lions Info tab")
    parser.add_argument(
        "--row", type=int, default=None,
        help="Update a single row (reads name from col A, writes D–V). E.g. --row 60",
    )
    parser.add_argument(
        "--skip-pfr", action="store_true",
        help="Skip PFR lookups (columns D, K–S will be blank)",
    )
    parser.add_argument(
        "--start-row", type=int, default=DATA_START_ROW,
        help=f"First data row for full-roster mode (default: {DATA_START_ROW})",
    )
    args = parser.parse_args()

    svc_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "credentials.json")
    if not os.path.exists(svc_file):
        print(f"Error: '{svc_file}' not found. Add your real Google service-account credentials.")
        sys.exit(1)

    print(f"Connecting to sheet {SHEET_ID}…")
    spreadsheet = connect_sheet(svc_file, SHEET_ID)
    ws = get_or_create_ws(spreadsheet, INFO_TAB)

    college_map = load_college_map(spreadsheet)

    # ---- Single-row mode ------------------------------------------------
    if args.row:
        row_num = args.row
        print(f"\nSingle-row mode: row {row_num}")

        # Read name from column A
        cell_val = ws.acell(f"A{row_num}").value or ""
        name = cell_val.strip()
        if not name:
            print(f"Error: A{row_num} is empty. Put the player name there first.")
            sys.exit(1)

        print(f"  Player: {name}")

        # Get physical stats from Lions/ESPN roster
        print("  Fetching roster for physical stats…")
        roster = fetch_roster()
        roster_by_name = {p["name"].lower(): p for p in roster}
        player = roster_by_name.get(name.lower())
        if not player:
            # Build minimal player dict from sheet columns if available
            row_data = ws.row_values(row_num)
            parts = name.split()
            player = {
                "name":          name,
                "last_name":     parts[-1] if parts else "",
                "number":        row_data[2].strip() if len(row_data) > 2 else "",
                "position":      row_data[5].strip() if len(row_data) > 5 else "",
                "ht_ft":         row_data[6].strip() if len(row_data) > 6 else "",
                "ht_in":         row_data[7].strip() if len(row_data) > 7 else "",
                "weight":        row_data[8].strip() if len(row_data) > 8 else "",
                "espn_id":       "",
                "espn_headshot": "",
            }
            print(f"  '{name}' not found in live roster — using existing sheet data for A–C/F–I")

        # PFR
        pfr = {}
        if not args.skip_pfr:
            print(f"  Fetching PFR data for {name}…")
            pfr = fetch_pfr(name)
            status = "✓" if pfr.get("pfr_url") else "–"
            print(f"  PFR {status}")

        college = college_map.get(name.lower(), "")
        row_values = build_row(player, pfr, row_num, college)

        print(f"  Writing D{row_num}:V{row_num}…")
        write_single_row_d_to_v(ws, row_values, row_num)
        print(f"Done — row {row_num} updated.")
        return

    # ---- Full roster mode -----------------------------------------------
    print("\nFull-roster mode")
    print("Fetching Lions roster…")
    players = fetch_roster()
    if not players:
        print("No roster data returned. Aborting.")
        sys.exit(1)

    all_rows = []
    for i, player in enumerate(players):
        row_num = args.start_row + i
        name = player["name"]
        print(f"  [{row_num}] {name}…", end=" ", flush=True)

        pfr = fetch_pfr(name) if not args.skip_pfr else {}
        college = college_map.get(name.lower(), "")
        row_values = build_row(player, pfr, row_num, college)
        all_rows.append(row_values)

        pfr_ok = "✓" if pfr.get("pfr_url") else "–"
        print(f"pfr:{pfr_ok}")

    print(f"\nWriting {len(all_rows)} rows to Info tab…")
    write_all_rows_a_to_v(ws, all_rows, args.start_row)
    print("Done.")


if __name__ == "__main__":
    main()
