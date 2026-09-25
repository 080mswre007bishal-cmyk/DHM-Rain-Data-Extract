"""
discover_stations.py
---------------------
Run this ONCE (locally, not in the scheduled job) to build stations.json.

DHM gives every rainfall station a page at:
    https://dhm.gov.np/hydrology/rainfallSingle/<url_id>

The station's name, official "Station Index" code, lat/long, elevation and
district ARE present in the plain HTML of that page (no JavaScript needed).
Only the live rainfall *reading* requires clicking "VIEW" (handled by
main.py with a real browser). So discovery is done here with plain
requests, which is fast and doesn't need Playwright.

We don't know exactly how many stations exist or which URL ids are used,
so we sweep a range of ids and keep whatever resolves to a real station
page. Adjust MAX_ID upward if you suspect stations are being missed
(check the end of stations.json — if it's still finding new ones near
MAX_ID, raise the number and re-run).

Usage:
    pip install requests beautifulsoup4
    python discover_stations.py
"""

import json
import re
import time

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://dhm.gov.np/hydrology/rainfallSingle/{}"
MAX_ID = 1200          # sweep url ids 1..MAX_ID — raise if needed, see note above
DELAY_SECONDS = 0.3    # be polite to a government server
OUTPUT_FILE = "stations.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; rainfall-data-research/1.0; contact: you@example.com)"
}

# Labels appear in this fixed order in the page text, one value between
# each pair of labels (a value can be blank, e.g. missing Elevation).
LABELS = ["Station Index", "Latitude", "Longitude", "Elevation", "District", "Description"]


def parse_station_page(html, url_id):
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n")
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # Name + id heading looks like "Suryapura (857)" and appears right
    # before the "Observations" line.
    name = None
    for i, line in enumerate(lines):
        if line == "Observations" and i > 0:
            m = re.match(r"^(.*)\s\(\d+\)$", lines[i - 1])
            name = m.group(1).strip() if m else lines[i - 1]
            break

    if not name:
        return None  # not a real station page (bad id, redirected, etc.)

    # Find each label's line index, value = everything between this label
    # and the next label (handles missing/blank fields gracefully).
    positions = {}
    for label in LABELS:
        for i, line in enumerate(lines):
            if line == label:
                positions[label] = i
                break

    if "Station Index" not in positions:
        return None

    ordered = sorted(positions.items(), key=lambda kv: kv[1])
    values = {}
    for idx, (label, pos) in enumerate(ordered):
        end = ordered[idx + 1][1] if idx + 1 < len(ordered) else pos + 2
        val_lines = lines[pos + 1:end]
        values[label] = " ".join(val_lines).strip()

    return {
        "url_id": url_id,
        "name": name,
        "station_index": values.get("Station Index", ""),
        "latitude": values.get("Latitude", ""),
        "longitude": values.get("Longitude", ""),
        "elevation": values.get("Elevation", ""),
        "district": values.get("District", ""),
        "description": values.get("Description", ""),
    }


def main():
    stations = []
    session = requests.Session()
    session.headers.update(HEADERS)

    for url_id in range(1, MAX_ID + 1):
        try:
            resp = session.get(BASE_URL.format(url_id), timeout=15)
            if resp.status_code == 200:
                station = parse_station_page(resp.text, url_id)
                if station:
                    stations.append(station)
                    print(f"[{url_id}] {station['name']} ({station['station_index']})")
        except requests.RequestException as e:
            print(f"[{url_id}] request failed: {e}")

        time.sleep(DELAY_SECONDS)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(stations, f, ensure_ascii=False, indent=2)

    print(f"\nDone. Found {len(stations)} stations. Saved to {OUTPUT_FILE}")
    print("Commit this file to your repo so the scheduled job doesn't redo this sweep.")


if __name__ == "__main__":
    main()
