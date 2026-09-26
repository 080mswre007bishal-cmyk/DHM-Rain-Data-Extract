"""
DHM Rainfall Watch Map -> Daily station rainfall -> Excel

Workflow:
1. Open https://dhm.gov.np/hydrology/rainfall-watch-map
2. Click "24 Hour".
3. Click the station markers on the map and read the popup. The popup is
   expected to contain a link to /hydrology/rainfallSingle/<id>.
4. Open that station page.
5. Set Date = yesterday (Nepal time), Period = Daily, click VIEW.
6. Read the Accumulated Rainfall table.
7. Append only dates not already present in output/*.xlsx.

Install:
    pip install playwright openpyxl
    playwright install chromium

Useful environment variables:
    DEBUG=1             save screenshots/html for troubleshooting
    MAX_STATIONS=10     test only the first N discovered stations
    HEADLESS=0          show the browser locally
"""

import datetime as dt
import json
import os
import re
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


MAP_URL = "https://dhm.gov.np/hydrology/rainfall-watch-map"
STATION_URL = "https://dhm.gov.np/hydrology/rainfallSingle/{}"

OUTPUT_DIR = Path("output")
DEBUG_DIR = OUTPUT_DIR / "debug"

DEBUG = os.getenv("DEBUG", "0") == "1"
HEADLESS = os.getenv("HEADLESS", "1") != "0"
MAX_STATIONS = int(os.getenv("MAX_STATIONS", "0"))  # 0 = all


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

NPT = dt.timezone(dt.timedelta(hours=5, minutes=45))


def yesterday_npt():
    return (dt.datetime.now(NPT).date() - dt.timedelta(days=1))


def yesterday_iso():
    return yesterday_npt().strftime("%Y-%m-%d")


DATE_FORMATS = [
    "%a, %b %d, %Y, %I:%M %p",  # Fri, Sep 25, 2026, 12:00 AM
    "%a, %b %d, %Y",
    "%Y-%m-%d",
    "%d %b, %Y",
    "%d %B, %Y",
    "%d %b %Y",
    "%d %B %Y",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%d-%m-%Y",
]


def parse_date(text):
    text = text.strip()
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass

    # Search inside a longer table cell.
    patterns = [
        r"[A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4}(?:,\s+\d{1,2}:\d{2}\s*[AP]M)?",
        r"\d{4}-\d{2}-\d{2}",
        r"\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{4}",
        r"\d{1,2}/\d{1,2}/\d{4}",
        r"\d{1,2}-\d{1,2}-\d{4}",
    ]
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            for fmt in DATE_FORMATS:
                try:
                    return dt.datetime.strptime(m.group(0), fmt).strftime("%Y-%m-%d")
                except ValueError:
                    pass
    return None


def parse_number(text):
    text = text.replace(",", "").strip()
    if not text:
        return None

    # Prefer a decimal/integer token.
    m = re.search(r"-?\d+(?:\.\d+)?", text)
    if not m:
        return None

    try:
        return float(m.group(0))
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# General browser helpers
# ---------------------------------------------------------------------------

def save_debug(page, tag):
    if not DEBUG:
        return
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w.-]+", "_", tag)
    page.screenshot(path=str(DEBUG_DIR / f"{safe}.png"), full_page=True)
    (DEBUG_DIR / f"{safe}.html").write_text(
        page.content(), encoding="utf-8"
    )


def click_text(page, text, timeout=8000):
    """Click exact visible text, with a few fallbacks."""
    candidates = [
        page.get_by_text(text, exact=True).first,
        page.locator(f"text={text}").first,
        page.get_by_role("button", name=text, exact=True).first,
        page.locator(f"a:has-text('{text}')").first,
        page.locator(f"button:has-text('{text}')").first,
    ]

    for loc in candidates:
        try:
            loc.wait_for(state="visible", timeout=timeout)
            loc.click()
            return True
        except Exception:
            continue

    return False


# ---------------------------------------------------------------------------
# 1. Open map and click 24 Hour
# ---------------------------------------------------------------------------

def open_24_hour_map(page):
    print("Opening DHM Rainfall Watch Map...")
    page.goto(MAP_URL, wait_until="domcontentloaded", timeout=60000)

    # The page currently exposes the six tabs:
    # 1 Hour, 3 Hour, 6 Hour, 12 Hour, 24 Hour, Tabular View.
    if not click_text(page, "24 Hour", timeout=15000):
        raise RuntimeError("Could not click the '24 Hour' tab.")

    # Give the map layer time to redraw after switching duration.
    page.wait_for_timeout(2500)

    # Wait until the map itself exists.
    try:
        page.locator("#map").wait_for(state="visible", timeout=15000)
    except Exception:
        # Some versions may use a different map container.
        if not page.locator(".leaflet-container").count():
            raise RuntimeError("DHM map container was not found.")

    save_debug(page, "00_24_hour_map")
    print("24 Hour map loaded.")


# ---------------------------------------------------------------------------
# 2. Click station markers and collect rainfallSingle links
# ---------------------------------------------------------------------------

def marker_count(page):
    selectors = [
        "#map .leaflet-marker-icon",
        ".leaflet-container .leaflet-marker-icon",
        "#map .leaflet-interactive",
        ".leaflet-container .leaflet-interactive",
    ]
    count = 0
    for selector in selectors:
        try:
            count = max(count, page.locator(selector).count())
        except Exception:
            pass
    return count


def station_id_from_url(url):
    if not url:
        return None
    m = re.search(r"/hydrology/rainfallSingle/(\d+)", url)
    return m.group(1) if m else None


def extract_station_id_from_text(text):
    """Find a rainfallSingle ID or common station-id attribute in HTML/JS."""
    if not text:
        return None

    patterns = [
        r"/hydrology/rainfallSingle/(\d+)",
        r"rainfallSingle[\"'\s:/]+(\d+)",
        r"(?:station[_ -]?id|stationId|stationID|station_number|stationNumber)\s*[=:]\s*[\"']?(\d+)",
        r"(?:station[_ -]?index|stationIndex)\s*[=:]\s*[\"']?(\d+)",
    ]

    for pattern in patterns:
        m = re.search(pattern, text, flags=re.I)
        if m:
            return m.group(1)
    return None


def extract_station_name(popup_text, station_id):
    lines = [x.strip() for x in popup_text.splitlines() if x.strip()]
    bad = {
        "VIEW", "View", "DETAIL", "Details", "Station",
        "Rainfall", "Rainfall Watch Map", "24 Hour",
    }

    for line in lines:
        if line in bad:
            continue
        if re.fullmatch(r"[\d.,\-\s]+", line):
            continue
        if "rainfallSingle" in line:
            continue
        if len(line) > 1:
            return line[:150]

    return f"Station_{station_id}"


def add_station(stations, station_id, name=None):
    if not station_id:
        return False
    station_id = str(station_id)
    if station_id in stations:
        return False

    stations[station_id] = {
        "url_id": station_id,
        "name": name or f"Station_{station_id}",
        "url": STATION_URL.format(station_id),
    }
    print(f"  [{len(stations)}] {stations[station_id]['name']} -> {stations[station_id]['url']}")
    return True


def collect_station_links_from_dom(page, stations):
    """Look everywhere in the rendered DOM, not only inside a popup."""
    try:
        html = page.content()
    except Exception:
        html = ""

    # Direct rainfallSingle URLs in HTML/inline JavaScript.
    for sid in re.findall(r"/hydrology/rainfallSingle/(\d+)", html):
        add_station(stations, sid)

    # Inspect links, buttons and elements carrying URL/data attributes.
    try:
        elements = page.locator("a,button,[onclick],[data-url],[data-href],[data-id],[data-station-id]")
        for i in range(min(elements.count(), 5000)):
            el = elements.nth(i)
            try:
                attrs = el.evaluate("""e => ({
                    text: e.innerText || '',
                    href: e.getAttribute('href') || '',
                    onclick: e.getAttribute('onclick') || '',
                    dataUrl: e.getAttribute('data-url') || '',
                    dataHref: e.getAttribute('data-href') || '',
                    dataId: e.getAttribute('data-id') || '',
                    dataStationId: e.getAttribute('data-station-id') || ''
                })""")
                blob = " ".join(str(v) for v in attrs.values())
                sid = extract_station_id_from_text(blob)
                if sid:
                    add_station(stations, sid, attrs.get("text", "").strip() or None)
            except Exception:
                continue
    except Exception:
        pass


def click_popup_station_control(page, popup, stations):
    """Some DHM versions put the station navigation in a JS button rather
    than an <a href>. Click those controls and capture the resulting URL."""
    controls = popup.locator("a,button,[role='button'],[onclick],[data-url],[data-href]")

    for i in range(controls.count()):
        control = controls.nth(i)
        try:
            attrs = control.evaluate("""e => ({
                text: e.innerText || '',
                href: e.getAttribute('href') || '',
                onclick: e.getAttribute('onclick') || '',
                dataUrl: e.getAttribute('data-url') || '',
                dataHref: e.getAttribute('data-href') || ''
            })""")
            blob = " ".join(str(v) for v in attrs.values())
            sid = extract_station_id_from_text(blob)
            if sid:
                add_station(stations, sid, attrs.get("text", "").strip() or None)
                return True

            text = attrs.get("text", "").strip().lower()
            if not any(x in text for x in ("view", "detail", "station")):
                continue

            old_url = page.url
            control.click(force=True, timeout=3000)
            page.wait_for_timeout(500)

            sid = station_id_from_url(page.url)
            if sid:
                add_station(stations, sid)
                if page.url != old_url:
                    page.go_back(wait_until="domcontentloaded", timeout=15000)
                return True
        except Exception:
            continue

    return False


def discover_stations_from_map(page):
    """
    Discover station IDs from the 24-hour map.

    Important: the DHM map markers do NOT necessarily contain a normal
    rainfallSingle href inside the Leaflet popup. Earlier code assumed they
    did, which is why 729 markers could be found but 0 stations discovered.

    This version uses several layers:
      1. rendered DOM/inline JS scan for rainfallSingle URLs;
      2. actual marker click + popup inspection;
      3. clicking View/Details controls inside the popup;
      4. a Tabular View fallback if the map popup does not expose IDs.
    """
    print("Discovering stations from the 24 Hour map...")
    stations = {}

    # First check whether the page already contains station URLs in its
    # rendered HTML or JavaScript data.
    collect_station_links_from_dom(page, stations)
    if stations:
        print(f"Found {len(stations)} station URL(s) in the map DOM/data.")

    # Wait for marker rendering.
    for _ in range(30):
        if marker_count(page) > 0:
            break
        page.wait_for_timeout(500)

    total = marker_count(page)
    print(f"Found about {total} map marker element(s).")

    # Prefer one marker selector. Using both .leaflet-marker-icon and
    # .leaflet-interactive previously caused duplicate/overlapping clicks.
    selector = "#map .leaflet-marker-icon"
    if page.locator(selector).count() == 0:
        selector = ".leaflet-container .leaflet-marker-icon"

    loc = page.locator(selector)
    count = loc.count()
    print(f"Scanning {count} marker icon(s)...")

    # Do not make a 729-marker scan unnecessarily slow if we already have all
    # stations from page data. Still allow MAX_STATIONS testing.
    marker_limit = count
    if MAX_STATIONS:
        marker_limit = min(marker_limit, max(MAX_STATIONS * 5, 30))

    for i in range(marker_limit):
        try:
            marker = loc.nth(i)
            if not marker.is_visible():
                continue

            marker.scroll_into_view_if_needed(timeout=2000)
            marker.click(force=True, timeout=3000)
            page.wait_for_timeout(350)

            popup = page.locator(".leaflet-popup").last
            if not popup.is_visible():
                # Try a DOM click event as a fallback for Leaflet markers whose
                # native click is intercepted by an overlay.
                try:
                    marker.dispatch_event("click")
                    page.wait_for_timeout(350)
                except Exception:
                    pass

            if not popup.is_visible():
                continue

            popup_html = popup.inner_html()
            popup_text = popup.inner_text()

            # 1) direct URL / ID in popup HTML.
            sid = extract_station_id_from_text(popup_html)
            if sid:
                add_station(stations, sid, extract_station_name(popup_text, sid))

            # 2) inspect/click any popup navigation control.
            click_popup_station_control(page, popup, stations)

            # 3) popup may reveal a station name only. Scan the whole DOM
            # after the click because DHM can create a hidden link/modal.
            collect_station_links_from_dom(page, stations)

        except Exception:
            pass

        finally:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass

        if MAX_STATIONS and len(stations) >= MAX_STATIONS:
            break

    # Fallback: DHM exposes a Tabular View on the same page. On versions where
    # Leaflet popups contain only rainfall values, the station rows may expose
    # the station URL/data-id instead.
    if not stations:
        print("No station IDs found from markers; trying Tabular View fallback...")
        try:
            if click_text(page, "Tabular View", timeout=8000):
                page.wait_for_timeout(2000)
                collect_station_links_from_dom(page, stations)

                # Inspect table rows for href/data attributes and station text.
                tables = page.locator("table")
                for ti in range(tables.count()):
                    table = tables.nth(ti)
                    rows = table.locator("tr")
                    for ri in range(rows.count()):
                        row = rows.nth(ri)
                        try:
                            html = row.inner_html()
                            sid = extract_station_id_from_text(html)
                            if sid:
                                add_station(stations, sid, row.inner_text().splitlines()[0].strip())
                        except Exception:
                            continue
        except Exception as exc:
            print(f"Tabular View fallback failed: {exc}")

    save_debug(page, "01_after_station_discovery")

    result = list(stations.values())
    if MAX_STATIONS:
        result = result[:MAX_STATIONS]

    print(f"Discovered {len(result)} unique station(s).")
    return result


# ---------------------------------------------------------------------------
# 3. Set Date = yesterday and Period = Daily
# ---------------------------------------------------------------------------

def find_date_input(page):
    """Find the Date input without assuming a fixed element ID."""
    inputs = page.locator("input")
    candidates = []

    for i in range(inputs.count()):
        try:
            el = inputs.nth(i)
            attrs = el.evaluate(
                """e => ({
                    type: e.type || '',
                    id: e.id || '',
                    name: e.name || '',
                    placeholder: e.placeholder || '',
                    value: e.value || '',
                    aria: e.getAttribute('aria-label') || ''
                })"""
            )
            blob = " ".join(str(v).lower() for v in attrs.values())

            score = 0
            if attrs["type"].lower() == "date":
                score += 100
            if "date" in blob:
                score += 30
            if "from" in blob or "start" in blob:
                score += 5

            candidates.append((score, i, attrs))

        except Exception:
            pass

    candidates.sort(reverse=True)

    if candidates and candidates[0][0] > 0:
        return inputs.nth(candidates[0][1])

    return None


def set_date_yesterday(page):
    target = yesterday_iso()
    inp = find_date_input(page)

    if inp is None:
        raise RuntimeError("Could not locate the Date input on the station page.")

    # First try Playwright fill.
    try:
        inp.fill(target, timeout=3000)
    except Exception:
        pass

    # Then explicitly set the DOM value and fire events. This handles date
    # pickers/framework-bound inputs that ignore a normal fill.
    try:
        inp.evaluate(
            """(e, value) => {
                const setter =
                  Object.getOwnPropertyDescriptor(
                    HTMLInputElement.prototype, 'value'
                  ).set;
                setter.call(e, value);
                e.dispatchEvent(new Event('input', {bubbles: true}));
                e.dispatchEvent(new Event('change', {bubbles: true}));
                e.blur();
            }""",
            target,
        )
    except Exception:
        pass

    # Verify what the browser actually has.
    try:
        actual = inp.input_value()
    except Exception:
        actual = ""

    if actual != target:
        raise RuntimeError(
            f"Date input did not accept yesterday. Expected {target}, got {actual!r}"
        )

    return target


def set_period_daily(page):
    selects = page.locator("select")

    for i in range(selects.count()):
        select = selects.nth(i)
        try:
            options = select.locator("option").all_inner_texts()
            if any(o.strip().lower() == "daily" for o in options):
                select.select_option(label="Daily")
                return True
        except Exception:
            pass

    # Fallback if the site uses a custom dropdown.
    for selector in [
        "label:has-text('Daily')",
        "text=Daily",
        "[value='daily']",
    ]:
        try:
            loc = page.locator(selector).first
            loc.click(timeout=2000)
            return True
        except Exception:
            pass

    return False


def click_view(page):
    for selector in [
        "button:has-text('VIEW')",
        "input[value='VIEW']",
        "a:has-text('VIEW')",
        "text=VIEW",
    ]:
        try:
            loc = page.locator(selector).first
            loc.wait_for(state="visible", timeout=3000)
            loc.click()
            return True
        except Exception:
            pass

    return False


# ---------------------------------------------------------------------------
# 4. Read the Accumulated Rainfall table
# ---------------------------------------------------------------------------

def read_rainfall_table(page):
    """
    Select the table whose header contains Date + Daily. This is safer than
    simply reading the first <table>, because the station page can contain
    other tables/HTML elements.
    """
    tables = page.locator("table")

    for i in range(tables.count()):
        table = tables.nth(i)
        try:
            text = table.inner_text()
            low = text.lower()

            if "date" not in low or "daily" not in low:
                continue

            rows = table.locator("tr")
            extracted = []

            for r in range(rows.count()):
                cells = rows.nth(r).locator("th, td")
                vals = [x.strip() for x in cells.all_inner_texts()]

                if len(vals) < 2:
                    continue

                date_str = parse_date(vals[0])
                if not date_str:
                    continue

                daily = parse_number(vals[1])
                total = parse_number(vals[2]) if len(vals) >= 3 else None

                if daily is not None:
                    extracted.append((date_str, daily, total))

            if extracted:
                return extracted

        except Exception:
            continue

    return []


def extract_station_history(page, debug_tag):
    target = set_date_yesterday()
    daily_ok = set_period_daily()

    if not daily_ok:
        raise RuntimeError("Could not set Period to Daily.")

    if not click_view(page):
        raise RuntimeError("Could not click VIEW.")

    # Wait until the table actually appears.
    try:
        page.locator("table").filter(
            has_text="Daily"
        ).first.wait_for(state="visible", timeout=15000)
    except Exception:
        page.wait_for_timeout(3000)

    # Give AJAX-rendered table rows time to settle.
    page.wait_for_timeout(1000)

    rows = read_rainfall_table(page)

    if DEBUG:
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        (DEBUG_DIR / f"{debug_tag}.txt").write_text(
            f"target_date={target}\n"
            f"period_daily={daily_ok}\n"
            f"rows={rows}\n\n"
            f"--- PAGE TEXT ---\n"
            + page.locator("body").inner_text(),
            encoding="utf-8",
        )
        save_debug(page, debug_tag)

    if not rows:
        raise RuntimeError("No Daily rainfall table rows were extracted.")

    return rows


# ---------------------------------------------------------------------------
# 5. Excel output
# ---------------------------------------------------------------------------

def safe_filename(name, station_id):
    clean = re.sub(r"[^\w\-]+", "_", name).strip("_")
    return f"{station_id}_{clean or 'station'}.xlsx"


def existing_dates(path):
    if not path.exists():
        return set()

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    found = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        value = row[0]
        if isinstance(value, dt.datetime):
            found.add(value.strftime("%Y-%m-%d"))
        elif isinstance(value, dt.date):
            found.add(value.strftime("%Y-%m-%d"))
        else:
            parsed = parse_date(str(value))
            if parsed:
                found.add(parsed)
            else:
                found.add(str(value))

    wb.close()
    return found


def append_rows(path, rows):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if path.exists():
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Daily Rainfall"
        ws.append([
            "Date",
            "Daily Rainfall (mm)",
            "Running Total on page (mm)",
        ])

    existing = existing_dates(path)

    added = 0
    for date_str, daily, total in sorted(rows):
        if date_str in existing:
            continue

        ws.append([date_str, daily, total])
        added += 1
        existing.add(date_str)

    # Store actual Excel dates for easier sorting/filtering.
    for cell in ws["A"][1:]:
        if isinstance(cell.value, str):
            try:
                cell.value = dt.datetime.strptime(
                    cell.value, "%Y-%m-%d"
                ).date()
                cell.number_format = "yyyy-mm-dd"
            except ValueError:
                pass

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)

    return added


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    OUTPUT_DIR.mkdir(exist_ok=True)

    run_time = dt.datetime.now(dt.timezone.utc).strftime(
        "%Y-%m-%d %H:%M:%S UTC"
    )

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
            ],
        )

        context = browser.new_context(
            viewport={"width": 1440, "height": 1000},
            locale="en-US",
            timezone_id="Asia/Kathmandu",
        )

        map_page = context.new_page()

        try:
            open_24_hour_map(map_page)
            stations = discover_stations_from_map(map_page)

            if not stations:
                raise RuntimeError(
                    "No station links were discovered by clicking the map. "
                    "Run with DEBUG=1 HEADLESS=0 to inspect the map."
                )

            print("\nStarting station extraction...")
            print(f"Yesterday in Nepal: {yesterday_iso()}")
            print(f"Run time: {run_time}\n")

            for number, station in enumerate(stations, start=1):
                station_id = station["url_id"]
                name = station["name"]
                url = station["url"]

                print(f"[{number}/{len(stations)}] {name} ({station_id})")

                station_page = context.new_page()

                try:
                    station_page.goto(
                        url,
                        wait_until="domcontentloaded",
                        timeout=60000,
                    )

                    station_page.wait_for_timeout(1000)

                    rows = extract_station_history(
                        station_page,
                        f"station_{station_id}",
                    )

                    path = OUTPUT_DIR / safe_filename(name, station_id)
                    added = append_rows(path, rows)

                    print(
                        f"    extracted {len(rows)} row(s); "
                        f"added {added} new row(s) -> {path}"
                    )

                except Exception as exc:
                    print(f"    FAILED: {exc}")
                    save_debug(station_page, f"ERROR_station_{station_id}")

                finally:
                    station_page.close()

                # Small delay so the site is not hammered.
                time.sleep(0.75)

        finally:
            map_page.close()
            browser.close()

    print("\nDone.")


if __name__ == "__main__":
    main()
