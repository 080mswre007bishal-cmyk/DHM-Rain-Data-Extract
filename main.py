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
  4. In the popup panel that opens, select "Daily" in the Period
     dropdown, click VIEW, and read the resulting Date | Daily | Total
     table.
  5. Skip today's row and write the rest to output/<station>.xlsx.

Dedup is currently OFF while debugging -- every row gets written every run.
"""

import datetime
import json
import os
import re
import time

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIGURATION
# ============================================================

STATIONS_FILE = "stations.json"
OUTPUT_DIR = "output"
DEBUG_DIR = os.path.join(OUTPUT_DIR, "debug")

MAP_URL = "https://dhm.gov.np/hydrology/rainfall-watch-map"

# Set DEBUG_STATIONS=3 to process only first 3 stations.
# Example:
#   DEBUG_STATIONS=3 python main.py
DEBUG_STATIONS = int(os.environ.get("DEBUG_STATIONS", "0"))

NEPAL_OFFSET = datetime.timedelta(hours=5, minutes=45)


# ============================================================
# DATE PARSING
# ============================================================

DATE_FORMATS = [
    "%a, %b %d, %Y, %I:%M %p",
    "%Y-%m-%d",
    "%d %b, %Y",
    "%d %B, %Y",
    "%d %b %Y",
    "%d %B %Y",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%d-%m-%Y",
]

DATE_PATTERNS = [
    r"[A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4}(?:,\s+\d{1,2}:\d{2}\s*[AP]M)?",
    r"\d{4}-\d{2}-\d{2}",
    r"\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{4}",
    r"\d{1,2}/\d{1,2}/\d{4}",
    r"\d{1,2}-\d{1,2}-\d{4}",
]


# ============================================================
# FILE / NAME HELPERS
# ============================================================

def safe_filename(name, station_index, url_id):
    base = station_index or url_id
    clean_name = re.sub(r"[^\w\-]+", "_", name).strip("_")
    return f"{base}_{clean_name}.xlsx"


def nepal_today_iso():
    now_npt = datetime.datetime.utcnow() + NEPAL_OFFSET
    return now_npt.date().strftime("%Y-%m-%d")


# ============================================================
# DATE / NUMBER PARSING
# ============================================================

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


# ============================================================
# DEBUG HELPERS
# ============================================================

def save_debug_file(tag, suffix, content):
    os.makedirs(DEBUG_DIR, exist_ok=True)

    safe = re.sub(r"[^\w.-]+", "_", tag)

    with open(
        os.path.join(DEBUG_DIR, f"{safe}{suffix}"),
        "w",
        encoding="utf-8",
    ) as f:
        f.write(content)


def dump_element_html(page_or_panel, text_to_find):
    """
    Find the DOM node whose exact text is text_to_find,
    walk up a few parent levels, and return its outerHTML.

    Returns a message string if not found.
    """

    try:
        return page_or_panel.evaluate(
            """(target) => {
                const walker = document.createTreeWalker(
                    document.body,
                    NodeFilter.SHOW_TEXT
                );

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


# ============================================================
# STATION FILTER
# ============================================================

def open_station_via_filter(page, station_index, debug_tag=None):
    """
    Set Filter type to Station No., pick station_index
    in the Station Number dropdown, click Apply Filter.

    Returns True/False.
    """

    selects = page.locator("select")

    filter_type_ok = False
    station_number_ok = False

    # --------------------------------------------------------
    # Find and select "Station No." filter type
    # --------------------------------------------------------

    try:
        for i in range(selects.count()):
            sel = selects.nth(i)

            opts = [
                o.strip()
                for o in sel.locator("option").all_inner_texts()
            ]

            if any("station no" in o.lower() for o in opts):

                match = next(
                    o for o in opts
                    if "station no" in o.lower()
                )

                sel.select_option(
                    label=match,
                    timeout=2000,
                )

                filter_type_ok = True
                break

    except Exception:
        pass

    # --------------------------------------------------------
    # Find and select station number
    # --------------------------------------------------------

    try:
        selects = page.locator("select")

        for i in range(selects.count()):
            sel = selects.nth(i)

            opts = [
                o.strip()
                for o in sel.locator("option").all_inner_texts()
            ]

            if station_index in opts:

                sel.select_option(
                    label=station_index,
                    timeout=2000,
                )

                station_number_ok = True
                break

    except Exception:
        pass

    # --------------------------------------------------------
    # Click Apply Filter
    # --------------------------------------------------------

    apply_ok = False

    for selector in [
        "button:has-text('Apply Filter')",
        "text='Apply Filter'",
        "input[value='Apply Filter']",
    ]:

        try:
            page.locator(selector).first.click(timeout=3000)

            apply_ok = True
            break

        except Exception:
            continue

    # --------------------------------------------------------
    # Debug screenshot / HTML
    # --------------------------------------------------------

    if debug_tag:

        try:
            page.wait_for_timeout(1500)

            # IMPORTANT:
            # Do not put re.sub(... r"\w" ...) directly
            # inside an f-string expression.
            safe_debug_tag = re.sub(
                r"[^\w.-]+",
                "_",
                debug_tag,
            )

            page.screenshot(
                path=os.path.join(
                    DEBUG_DIR,
                    f"{safe_debug_tag}_after_filter.png",
                )
            )

        except Exception:
            pass

        if not (
            filter_type_ok
            and station_number_ok
            and apply_ok
        ):
            save_debug_file(
                debug_tag,
                "_sidebar_html.txt",
                dump_element_html(
                    page,
                    "Apply Filter",
                ),
            )

    return (
        filter_type_ok
        and station_number_ok
        and apply_ok
    )


# ============================================================
# MAP MARKER
# ============================================================

def get_marker_locator(page):

    for selector in [
        "#map .leaflet-marker-icon",
        ".leaflet-container .leaflet-marker-icon",
    ]:

        loc = page.locator(selector)

        try:
            if loc.count() > 0:
                return loc

        except Exception:
            pass

    return None


# ============================================================
# STATION POPUP / PANEL
# ============================================================

def _candidate_panels(page):

    selectors = [
        "[role='dialog']:visible",
        ".modal.show:visible",
        ".modal.in:visible",
        ".offcanvas.show:visible",
        "div[class*='station']:visible",
        "section[class*='station']:visible",
        "div[class*='popup']:visible",
        "section[class*='popup']:visible",
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

                    inside_map = el.evaluate(
                        "el => !!el.closest('.leaflet-container')"
                    )

                    if inside_map:
                        continue

                    txt = el.inner_text(timeout=1000)

                    if (
                        "station index" in txt.lower()
                        and "view" in txt.lower()
                    ):
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

    ranked = sorted(
        panels,
        key=lambda p: len(p.inner_text()),
    )

    return ranked[0]


# ============================================================
# DAILY PERIOD
# ============================================================

def set_period_to_daily_in_panel(panel):
    """
    Real <select>-based Period control,
    scoped to just this panel.
    """

    selects = panel.locator("select")

    for i in range(selects.count()):

        sel = selects.nth(i)

        try:

            opts = [
                o.strip()
                for o in sel.locator("option").all_inner_texts()
            ]

            if any(
                o.lower() == "daily"
                for o in opts
            ):

                sel.select_option(
                    label="Daily",
                    timeout=2000,
                )

                return True

        except Exception:
            continue

    return False


# ============================================================
# VIEW BUTTON
# ============================================================

def click_view_in_panel(panel):

    for selector in [
        "text=VIEW",
        "button:has-text('VIEW')",
        "input[value='VIEW']",
    ]:

        try:

            panel.locator(selector).first.click(
                timeout=3000
            )

            return True

        except Exception:
            continue

    return False


# ============================================================
# TABLE PARSING
# ============================================================

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

            texts = [
                cells.nth(j).inner_text().strip()
                for j in range(cells.count())
            ]

        except Exception:
            continue

        if len(texts) < 2:
            continue

        # ----------------------------------------------------
        # Find date
        # ----------------------------------------------------

        date_str = None

        for pattern in DATE_PATTERNS:

            m = re.search(
                pattern,
                texts[0],
            )

            if m:

                date_str = parse_date_token(
                    m.group(0)
                )

                if date_str:
                    break

        if not date_str:
            continue

        # ----------------------------------------------------
        # Daily rainfall
        # ----------------------------------------------------

        daily_val = (
            parse_number(texts[1])
            if len(texts) > 1
            else None
        )

        # ----------------------------------------------------
        # Running total
        # ----------------------------------------------------

        total_val = (
            parse_number(texts[2])
            if len(texts) > 2
            else None
        )

        if daily_val is not None:

            rows_data.append(
                (
                    date_str,
                    daily_val,
                    total_val,
                )
            )

    return rows_data


# ============================================================
# CLOSE PANEL
# ============================================================

def close_any_panel(page):

    for selector in [
        "[aria-label='Close']:visible",
        "button[title='Close']:visible",
        ".btn-close:visible",
        ".close:visible",
        "button:has-text('Close'):visible",
    ]:

        try:

            loc = page.locator(selector).last

            if loc.is_visible():

                loc.click(
                    force=True,
                    timeout=1000,
                )

                page.wait_for_timeout(250)

                return

        except Exception:
            pass

    # Fallback
    try:

        page.keyboard.press("Escape")
        page.wait_for_timeout(250)

    except Exception:
        pass


# ============================================================
# PROCESS ONE STATION
# ============================================================

def process_station(page, station, debug_tag=None):
    """
    Returns list of:
        (date_str, daily_value, total_value)

    Today's row is excluded.
    Raises on failure.
    """

    station_index = station.get("station_index")

    if not station_index:

        raise RuntimeError(
            "This station has no station_index in stations.json "
            "-- can't filter by it."
        )

    # --------------------------------------------------------
    # Open map
    # --------------------------------------------------------

    page.goto(
        MAP_URL,
        wait_until="domcontentloaded",
        timeout=30000,
    )

    page.wait_for_timeout(2000)

    # --------------------------------------------------------
    # Apply station filter
    # --------------------------------------------------------

    filtered_ok = open_station_via_filter(
        page,
        station_index,
        debug_tag=debug_tag,
    )

    if not filtered_ok:

        raise RuntimeError(
            f"Could not apply the Station No. filter "
            f"for index {station_index!r}."
        )

    # --------------------------------------------------------
    # Find station panel
    # --------------------------------------------------------

    panel = None

    for _ in range(20):

        panel = find_station_panel(page)

        if panel is not None:
            break

        # Popup may need an explicit marker click
        # after filtering.

        markers = get_marker_locator(page)

        if markers and markers.count() > 0:

            try:

                markers.first.click(
                    force=True,
                    timeout=2000,
                )

            except Exception:
                pass

        time.sleep(0.5)

    # --------------------------------------------------------
    # No panel
    # --------------------------------------------------------

    if panel is None:

        if debug_tag:

            try:

                safe_debug_tag = re.sub(
                    r"[^\w.-]+",
                    "_",
                    debug_tag,
                )

                page.screenshot(
                    path=os.path.join(
                        DEBUG_DIR,
                        f"{safe_debug_tag}_no_panel.png",
                    )
                )

            except Exception:
                pass

        raise RuntimeError(
            "Filtered successfully but no station "
            "popup panel appeared."
        )

    # --------------------------------------------------------
    # Save panel debug files
    # --------------------------------------------------------

    if debug_tag:

        try:

            safe_debug_tag = re.sub(
                r"[^\w.-]+",
                "_",
                debug_tag,
            )

            panel.screenshot(
                path=os.path.join(
                    DEBUG_DIR,
                    f"{safe_debug_tag}_panel.png",
                )
            )

        except Exception:
            pass

        try:

            save_debug_file(
                debug_tag,
                "_panel_html.txt",
                panel.evaluate(
                    "el => el.outerHTML"
                ),
            )

        except Exception:
            pass

    # --------------------------------------------------------
    # Select Daily
    # --------------------------------------------------------

    period_ok = set_period_to_daily_in_panel(
        panel
    )

    # --------------------------------------------------------
    # Click VIEW
    # --------------------------------------------------------

    view_ok = click_view_in_panel(
        panel
    )

    # Wait for table
    try:

        page.wait_for_timeout(2000)

    except Exception:

        time.sleep(2)

    # --------------------------------------------------------
    # Read table
    # --------------------------------------------------------

    all_rows = table_rows_from_panel(
        panel
    )

    # --------------------------------------------------------
    # Remove today's row
    # --------------------------------------------------------

    today = nepal_today_iso()

    kept_rows = [
        r
        for r in all_rows
        if r[0] != today
    ]

    # --------------------------------------------------------
    # Debug result
    # --------------------------------------------------------

    if debug_tag:

        save_debug_file(
            debug_tag,
            "_result.txt",
            f"filtered_ok: {filtered_ok}\n"
            f"period_ok (selected Daily): {period_ok}\n"
            f"view_ok (clicked VIEW): {view_ok}\n"
            f"all_rows: {all_rows}\n"
            f"kept_rows: {kept_rows}\n",
        )

    # --------------------------------------------------------
    # Close popup
    # --------------------------------------------------------

    close_any_panel(page)

    return kept_rows


# ============================================================
# EXCEL OUTPUT
# ============================================================

def append_all_rows(path, entries, run_timestamp):
    """
    Dedup disabled while debugging.
    Every row is written every run.
    """

    if os.path.exists(path):

        wb = load_workbook(path)
        ws = wb.active

    else:

        wb = Workbook()
        ws = wb.active

        ws.append(
            [
                "Date",
                "Daily Rainfall (mm)",
                "Running Total on page (mm)",
                "First recorded (UTC)",
            ]
        )

    for date_str, daily_val, total_val in sorted(entries):

        ws.append(
            [
                date_str,
                daily_val,
                total_val,
                run_timestamp,
            ]
        )

    wb.save(path)


# ============================================================
# MAIN
# ============================================================

def main():

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True,
    )

    os.makedirs(
        DEBUG_DIR,
        exist_ok=True,
    )

    # --------------------------------------------------------
    # Load stations
    # --------------------------------------------------------

    with open(
        STATIONS_FILE,
        encoding="utf-8",
    ) as f:

        stations = json.load(f)

    # --------------------------------------------------------
    # Debug mode
    # --------------------------------------------------------

    if DEBUG_STATIONS:

        stations = stations[:DEBUG_STATIONS]

        print(
            f"DEBUG MODE: running only "
            f"{len(stations)} station(s)"
        )

    else:

        print(
            f"LIVE RUN: processing all "
            f"{len(stations)} station(s) "
            f"in {STATIONS_FILE}."
        )

    # --------------------------------------------------------
    # Run timestamp
    # --------------------------------------------------------

    run_timestamp = (
        datetime.datetime.utcnow()
        .strftime("%Y-%m-%d %H:%M:%S")
    )

    # --------------------------------------------------------
    # Playwright
    # --------------------------------------------------------

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=True
        )

        page = browser.new_page()

        # ----------------------------------------------------
        # Process stations
        # ----------------------------------------------------

        for station in stations:

            url_id = station["url_id"]
            name = station["name"]

            print(
                f"-> {name} "
                f"(station_index="
                f"{station.get('station_index')})"
            )

            filename = safe_filename(
                name,
                station.get("station_index"),
                url_id,
            )

            filepath = os.path.join(
                OUTPUT_DIR,
                filename,
            )

            # ------------------------------------------------
            # Debug tag
            # ------------------------------------------------

            debug_tag = None

            if DEBUG_STATIONS:

                clean_name_for_tag = re.sub(
                    r"[^\w-]+",
                    "_",
                    name,
                )

                debug_tag = (
                    f"{url_id}_"
                    f"{clean_name_for_tag}"
                )

            # ------------------------------------------------
            # Process station
            # ------------------------------------------------

            try:

                history = process_station(
                    page,
                    station,
                    debug_tag=debug_tag,
                )

            except Exception as e:

                print(
                    f"   failed: {e}"
                )

                history = []

            # ------------------------------------------------
            # Print result
            # ------------------------------------------------

            print(
                f"   read "
                f"{len(history)} row(s): "
                f"{history}"
            )

            # ------------------------------------------------
            # Write Excel
            # ------------------------------------------------

            if history:

                append_all_rows(
                    filepath,
                    history,
                    run_timestamp,
                )

                print(
                    f"   wrote "
                    f"{len(history)} row(s) "
                    f"(dedup disabled)"
                )

            else:

                print(
                    "   nothing parsed, "
                    "nothing written"
                )

            time.sleep(0.5)

        # ----------------------------------------------------
        # Close browser
        # ----------------------------------------------------

        browser.close()

    print("Done.")


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()