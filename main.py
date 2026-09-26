"""
main.py
-------
For every station in stations.json:
  1. Go to the DHM rainfall-watch-map page.
  2. In the "SORT & FILTERS" panel, set Filter type to "Station No."
     and pick this station's own Station Index in the Station Number
     dropdown, then click "Apply Filter" -- this targets the exact
     right station instead of guessing among ~700+ map markers.
  3. Click the marker that appears/gets highlighted.
  4. In the popup panel that opens (CONFIRMED via a real screenshot
     you shared of the Thakmarpha popup): select "Daily" in the Period
     dropdown (a real <select> here, unlike the broken tab-style
     control on the standalone rainfallSingle/<id> page), click VIEW,
     and read the resulting Date | Daily | Total table.
  5. Skip today's row (still accumulating) and write the rest to
     output/<station>.xlsx. Dedup is currently OFF (per earlier
     request, while still debugging) -- every row gets written every
     run.

IMPORTANT -- what's confirmed vs. not:
CONFIRMED: the popup panel's Period control is a real <select>
(screenshot), and its resulting table has Date/Daily/Total columns
with the same date format as before ("Fri, Sep 25, 2026, 12:00 AM").
NOT CONFIRMED (I cannot load dhm.gov.np myself, so this is
best-effort): the exact selectors for the "Filter type" dropdown, the
"Station Number" dropdown, the "Apply Filter" button, and whether
clicking Apply Filter alone opens the popup or whether the marker
still needs a separate click afterward.

Run DEBUG_STATIONS=3 (or the Debug Rainfall Extraction workflow) and
check output/debug/ for each station:
  - <tag>_after_filter.png       -- what the map looked like right
                                     after clicking Apply Filter
  - <tag>_sidebar_html.txt       -- the real HTML of the filter
                                     sidebar, if filtering failed
  - <tag>_panel.png / _panel_html.txt -- the popup panel, if one
                                     was found
  - <tag>.txt                    -- the full log: what got clicked,
                                     what rows got read
If Apply Filter or the marker click isn't landing, the sidebar/panel
HTML dumps are what to paste back for a real fix, the same way that
worked for the Daily-tab bug on the old page.

Usage:
    pip install playwright openpyxl
    playwright install chromium
    python main.py
"""

import datetime
import json
import os
import re
import time

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright

STATIONS_FILE = "stations.json"
OUTPUT_DIR = "output"
DEBUG_DIR = os.path.join(OUTPUT_DIR, "debug")
MAP_URL = "https://dhm.gov.np/hydrology/rainfall-watch-map"

DEBUG_STATIONS = int(os.environ.get("DEBUG_STATIONS", "0"))

NEPAL_OFFSET = datetime.timedelta(hours=5, minutes=45)

DATE_FORMATS = [
    "%a, %b %d, %Y, %I:%M %p",
    "%Y-%m-%d",
    "%d %b, %Y", "%d %B, %Y", "%d %b %Y", "%d %B %Y",
    "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y",
]
DATE_PATTERNS = [
    r"[A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4}(?:,\s+\d{1,2}:\d{2}\s*[AP]M)?",
    r"\d{4}-\d{2}-\d{2}",
    r"\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{4}",
    r"\d{1,2}/\d{1,2}/\d{4}",
    r"\d{1,2}-\d{1,2}-\d{4}",
]


def safe_filename(name, station_index, url_id):
    base = station_index or url_id
    clean_name = re.sub(r"[^\w\-]+", "_", name).strip("_")
    return f"{base}_{clean_name}.xlsx"


def nepal_today_iso():
    now_npt = datetime.datetime.utcnow() + NEPAL_OFFSET
    return now_npt.date().strftime("%Y-%m-%d")


def parse_date_token(token):
    token = token.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(token, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def parse_number(text):
    matches = re.findall(r"\d+\.\d+|\d+", text)
    if not matches:
        return None
    val = float(matches[0])
    return val if 0 <= val < 1000 else None


def save_debug_file(tag, suffix, content):
    os.makedirs(DEBUG_DIR, exist_ok=True)
    safe = re.sub(r"[^\w.-]+", "_", tag)
    with open(os.path.join(DEBUG_DIR, f"{safe}{suffix}"), "w", encoding="utf-8") as f:
        f.write(content)


def dump_element_html(page_or_panel, text_to_find):
    """Find the DOM node whose exact text is text_to_find, walk up a few
    parent levels, and return that outerHTML. Returns a message string
    if not found."""
    try:
        return page_or_panel.evaluate(
            """(target) => {
                const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
                let node;
                while (node = walker.nextNode()) {
                    if (node.textContent.trim() === target) {
                        let el = node.parentElement;
                        for (let i = 0; i < 4 && el.parentElement; i++) {
                            el = el.parentElement;
                        }
                        return el.outerHTML;
                    }
                }
                return 'NO NODE WITH EXACT TEXT "' + target + '" FOUND';
            }""",
            text_to_find,
        )
    except Exception as e:
        return f"ERROR while extracting: {e}"


def open_station_via_filter(page, station_index, debug_tag=None):
    """Set Filter type to Station No., pick station_index in the
    Station Number dropdown, click Apply Filter. Returns True/False."""
    selects = page.locator("select")
    filter_type_ok = False
    station_number_ok = False

    try:
        for i in range(selects.count()):
            sel = selects.nth(i)
            opts = [o.strip() for o in sel.locator("option").all_inner_texts()]
            if any("station no" in o.lower() for o in opts):
                match = next(o for o in opts if "station no" in o.lower())
                sel.select_option(label=match, timeout=2000)
                filter_type_ok = True
                break
    except Exception:
        pass

    try:
        selects = page.locator("select")  # re-query; DOM may have changed
        for i in range(selects.count()):
            sel = selects.nth(i)
            opts = [o.strip() for o in sel.locator("option").all_inner_texts()]
            if station_index in opts:
                sel.select_option(label=station_index, timeout=2000)
                station_number_ok = True
                break
    except Exception:
        pass

    apply_ok = False
    for selector in ["button:has-text('Apply Filter')", "text='Apply Filter'", "input[value='Apply Filter']"]:
        try:
            page.locator(selector).first.click(timeout=3000)
            apply_ok = True
            break
        except Exception:
            continue

    if debug_tag:
        try:
            page.wait_for_timeout(1500)
            page.screenshot(path=os.path.join(DEBUG_DIR, f"{re.sub(r'[^\\w.-]+', '_', debug_tag)}_after_filter.png"))
        except Exception:
            pass
        if not (filter_type_ok and station_number_ok and apply_ok):
            save_debug_file(debug_tag, "_sidebar_html.txt", dump_element_html(page, "Apply Filter"))

    return filter_type_ok and station_number_ok and apply_ok


def get_marker_locator(page):
    for selector in ["#map .leaflet-marker-icon", ".leaflet-container .leaflet-marker-icon"]:
        loc = page.locator(selector)
        try:
            if loc.count() > 0:
                return loc
        except Exception:
            pass
    return None


def _candidate_panels(page):
    selectors = [
        "[role='dialog']:visible",
        ".modal.show:visible", ".modal.in:visible",
        ".offcanvas.show:visible",
        "div[class*='station']:visible", "section[class*='station']:visible",
        "div[class*='popup']:visible", "section[class*='popup']:visible",
    ]
    result = []
    for selector in selectors:
        try:
            loc = page.locator(selector)
            for i in range(loc.count()):
                el = loc.nth(i)
                try:
                    if not el.is_visible():
                        continue
                    inside_map = el.evaluate("el => !!el.closest('.leaflet-container')")
                    if inside_map:
                        continue
                    txt = el.inner_text(timeout=1000)
                    if "station index" in txt.lower() and "view" in txt.lower():
                        result.append(el)
                except Exception:
                    pass
        except Exception:
            pass
    return result


def find_station_panel(page):
    panels = _candidate_panels(page)
    if not panels:
        return None
    ranked = sorted(panels, key=lambda p: len(p.inner_text()))
    return ranked[0]


def set_period_to_daily_in_panel(panel):
    """Real <select>-based Period control, scoped to just this panel
    (confirmed via your Thakmarpha popup screenshot)."""
    selects = panel.locator("select")
    for i in range(selects.count()):
        sel = selects.nth(i)
        try:
            opts = [o.strip() for o in sel.locator("option").all_inner_texts()]
            if any(o.lower() == "daily" for o in opts):
                sel.select_option(label="Daily", timeout=2000)
                return True
        except Exception:
            continue
    return False


def click_view_in_panel(panel):
    for selector in ["text=VIEW", "button:has-text('VIEW')", "input[value='VIEW']"]:
        try:
            panel.locator(selector).first.click(timeout=3000)
            return True
        except Exception:
            continue
    return False


def table_rows_from_panel(panel):
    rows_data = []
    try:
        rows = panel.locator("table tr")
        count = rows.count()
    except Exception:
        return rows_data

    for i in range(count):
        try:
            cells = rows.nth(i).locator("td, th")
            texts = [cells.nth(j).inner_text().strip() for j in range(cells.count())]
        except Exception:
            continue

        if len(texts) < 2:
            continue

        date_str = None
        for pattern in DATE_PATTERNS:
            m = re.search(pattern, texts[0])
            if m:
                date_str = parse_date_token(m.group(0))
                if date_str:
                    break
        if not date_str:
            continue

        daily_val = parse_number(texts[1]) if len(texts) > 1 else None
        total_val = parse_number(texts[2]) if len(texts) > 2 else None
        if daily_val is not None:
            rows_data.append((date_str, daily_val, total_val))

    return rows_data


def close_any_panel(page):
    for selector in [
        "[aria-label='Close']:visible", "button[title='Close']:visible",
        ".btn-close:visible", ".close:visible", "button:has-text('Close'):visible",
    ]:
        try:
            loc = page.locator(selector).last
            if loc.is_visible():
                loc.click(force=True, timeout=1000)
                page.wait_for_timeout(250)
                return
        except Exception:
            pass
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(250)
    except Exception:
        pass


def process_station(page, station, debug_tag=None):
    """Returns list of (date_str, daily_value, total_value), excluding
    today's row. Raises on failure."""
    station_index = station.get("station_index")
    if not station_index:
        raise RuntimeError("This station has no station_index in stations.json -- can't filter by it.")

    page.goto(MAP_URL, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(2000)

    filtered_ok = open_station_via_filter(page, station_index, debug_tag=debug_tag)
    if not filtered_ok:
        raise RuntimeError(f"Could not apply the Station No. filter for index {station_index!r}.")

    panel = None
    for _ in range(20):
        panel = find_station_panel(page)
        if panel is not None:
            break
        # Popup may need an explicit marker click after filtering.
        markers = get_marker_locator(page)
        if markers and markers.count() > 0:
            try:
                markers.first.click(force=True, timeout=2000)
            except Exception:
                pass
        time.sleep(0.5)

    if panel is None:
        if debug_tag:
            page.screenshot(path=os.path.join(DEBUG_DIR, f"{re.sub(r'[^\\w.-]+', '_', debug_tag)}_no_panel.png"))
        raise RuntimeError("Filtered successfully but no station popup panel appeared.")

    if debug_tag:
        try:
            panel.screenshot(path=os.path.join(DEBUG_DIR, f"{re.sub(r'[^\\w.-]+', '_', debug_tag)}_panel.png"))
        except Exception:
            pass
        try:
            save_debug_file(debug_tag, "_panel_html.txt", panel.evaluate("el => el.outerHTML"))
        except Exception:
            pass

    period_ok = set_period_to_daily_in_panel(panel)
    view_ok = click_view_in_panel(panel)
    panel.page.wait_for_timeout(2000) if hasattr(panel, "page") else time.sleep(2)

    all_rows = table_rows_from_panel(panel)
    today = nepal_today_iso()
    kept_rows = [r for r in all_rows if r[0] != today]

    if debug_tag:
        save_debug_file(
            debug_tag, "_result.txt",
            f"filtered_ok: {filtered_ok}\n"
            f"period_ok (selected Daily): {period_ok}\n"
            f"view_ok (clicked VIEW): {view_ok}\n"
            f"all_rows: {all_rows}\n"
            f"kept_rows: {kept_rows}\n"
        )

    close_any_panel(page)
    return kept_rows


def append_all_rows(path, entries, run_timestamp):
    """Dedup disabled while debugging -- every row gets written every run."""
    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Date", "Daily Rainfall (mm)", "Running Total on page (mm)", "First recorded (UTC)"])

    for date_str, daily_val, total_val in sorted(entries):
        ws.append([date_str, daily_val, total_val, run_timestamp])

    wb.save(path)


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(STATIONS_FILE, encoding="utf-8") as f:
        stations = json.load(f)

    if DEBUG_STATIONS:
        stations = stations[:DEBUG_STATIONS]
        print(f"DEBUG MODE: running only {len(stations)} station(s)")
    else:
        print(f"LIVE RUN: processing all {len(stations)} station(s) in {STATIONS_FILE}.")

    run_timestamp = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        for station in stations:
            url_id = station["url_id"]
            name = station["name"]
            print(f"-> {name} (station_index={station.get('station_index')})")

            filename = safe_filename(name, station.get("station_index"), url_id)
            filepath = os.path.join(OUTPUT_DIR, filename)

            debug_tag = None
            if DEBUG_STATIONS:
                clean_name_for_tag = re.sub(r"[^\w-]+", "_", name)
                debug_tag = f"{url_id}_{clean_name_for_tag}"

            try:
                history = process_station(page, station, debug_tag=debug_tag)
            except Exception as e:
                print(f"   failed: {e}")
                history = []

            print(f"   read {len(history)} row(s): {history}")

            if history:
                append_all_rows(filepath, history, run_timestamp)
                print(f"   wrote {len(history)} row(s) (dedup disabled)")
            else:
                print("   nothing parsed, nothing written")

            time.sleep(0.5)

        browser.close()

    print("Done.")


if __name__ == "__main__":
    main()
