import datetime as dt
import os
import re
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIGURATION
# ============================================================

MAP_URL = "https://dhm.gov.np/hydrology/rainfall-watch-map"

OUTPUT_DIR = Path("output")
DEBUG_DIR = OUTPUT_DIR / "debug"

DEBUG = os.getenv("DEBUG", "0") == "1"
HEADLESS = os.getenv("HEADLESS", "1") != "0"
MAX_STATIONS = int(os.getenv("MAX_STATIONS", "0"))

# DEBUG always processes at most these many markers.
DEBUG_MAX_STATIONS = 10


# ============================================================
# DATE
# ============================================================

def today_iso():
    return dt.date.today().strftime("%Y-%m-%d")


def yesterday_iso():
    return (dt.date.today() - dt.timedelta(days=1)).strftime("%Y-%m-%d")


# ============================================================
# DATE PARSING
# ============================================================

DATE_FORMATS = [
    "%a, %b %d, %Y, %I:%M %p",
    "%a, %b %d, %Y, %I:%M:%S %p",
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
    if not text:
        return None

    text = " ".join(str(text).split())

    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass

    patterns = [
        r"[A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4}(?:,\s+\d{1,2}:\d{2}(?::\d{2})?\s*[AP]M)?",
        r"\d{4}-\d{2}-\d{2}",
        r"\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{4}",
        r"\d{1,2}/\d{1,2}/\d{4}",
        r"\d{1,2}-\d{1,2}-\d{4}",
    ]

    for pattern in patterns:
        match = re.search(pattern, text)
        if not match:
            continue

        token = match.group(0)
        for fmt in DATE_FORMATS:
            try:
                return dt.datetime.strptime(token, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass

    return None


def parse_number(text):
    if text is None:
        return None

    text = str(text).replace(",", "").strip()

    if not text or text in {"-", "—", "N/A", "NA"}:
        return None

    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None

    try:
        return float(match.group(0))
    except ValueError:
        return None


# ============================================================
# DEBUG
# ============================================================

def save_debug(page, tag, popup=None):
    if not DEBUG:
        return

    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w.-]+", "_", tag)

    try:
        page.screenshot(
            path=str(DEBUG_DIR / f"{safe}_page.png"),
            full_page=True,
        )
    except Exception:
        pass

    try:
        (DEBUG_DIR / f"{safe}_page.html").write_text(
            page.content(), encoding="utf-8"
        )
    except Exception:
        pass

    try:
        (DEBUG_DIR / f"{safe}_page.txt").write_text(
            page.locator("body").inner_text(), encoding="utf-8"
        )
    except Exception:
        pass

    if popup is not None:
        try:
            (DEBUG_DIR / f"{safe}_popup.html").write_text(
                popup.evaluate("el => el.outerHTML"),
                encoding="utf-8",
            )
        except Exception:
            pass

        try:
            (DEBUG_DIR / f"{safe}_popup.txt").write_text(
                popup.inner_text(), encoding="utf-8"
            )
        except Exception:
            pass


# ============================================================
# MAP / 24 HOUR
# ============================================================

def click_24_hour(page):
    print("Opening DHM Rainfall Watch Map...")

    page.goto(
        MAP_URL,
        wait_until="domcontentloaded",
        timeout=60000,
    )

    page.wait_for_timeout(3000)

    candidates = [
        page.get_by_text("24 Hour", exact=True).first,
        page.get_by_role("button", name="24 Hour", exact=True).first,
        page.locator("a:has-text('24 Hour')").first,
        page.locator("button:has-text('24 Hour')").first,
    ]

    for locator in candidates:
        try:
            locator.wait_for(state="visible", timeout=5000)
            locator.click(force=True)
            page.wait_for_timeout(4000)
            print("24 Hour map loaded.")
            return
        except Exception:
            pass

    raise RuntimeError("Could not click DHM '24 Hour'.")


# ============================================================
# LEAFLET MARKERS
# ============================================================

def marker_locator(page):
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


def wait_for_markers(page, timeout=30):
    deadline = time.time() + timeout
    while time.time() < deadline:
        loc = marker_locator(page)
        if loc is not None:
            try:
                if loc.count() > 0:
                    return loc
            except Exception:
                pass
        page.wait_for_timeout(500)
    raise RuntimeError("No Leaflet station markers were found.")


# ============================================================
# ACTIVE LEAFLET POPUP
# ============================================================

def get_active_popup(page):
    """Return the currently visible Leaflet popup."""
    popups = page.locator(".leaflet-popup")

    for i in range(popups.count() - 1, -1, -1):
        popup = popups.nth(i)
        try:
            if popup.is_visible():
                return popup
        except Exception:
            pass

    return None


def wait_for_popup(page, timeout=15):
    """Wait until a Leaflet popup exists and has some text."""
    deadline = time.time() + timeout

    while time.time() < deadline:
        popup = get_active_popup(page)
        if popup is not None:
            try:
                if popup.inner_text().strip():
                    return popup
            except Exception:
                pass
        page.wait_for_timeout(250)

    raise RuntimeError("Marker was clicked but no Leaflet popup appeared.")


def wait_for_station_info(popup, timeout=12):
    """
    Wait specifically for Station and Station Index to appear INSIDE
    the active Leaflet popup.
    """
    deadline = time.time() + timeout

    while time.time() < deadline:
        try:
            text = popup.inner_text()
        except Exception:
            text = ""

        if re.search(r"station\s*index", text, re.I) and re.search(
            r"(?:station\s*(?:name)?\s*:|station\s+name\b)",
            text,
            re.I,
        ):
            return text

        page = popup.page
        page.wait_for_timeout(300)

    raise RuntimeError("Leaflet popup appeared, but Station/Station Index did not render inside it.")


# ============================================================
# STATION NAME + INDEX — POPUP ONLY
# ============================================================

def extract_station_info(popup, sequence):
    """Extract Station Name and actual Station Index ONLY from popup."""

    try:
        text = popup.inner_text()
    except Exception:
        text = ""

    normalized_text = "\n".join(
        re.sub(r"\s+", " ", line).strip()
        for line in text.splitlines()
        if line.strip()
    )

    # ----------------------------
    # Station Index
    # ----------------------------
    index_patterns = [
        r"station\s*index\s*(?:code|number|no\.?)*\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9_-]*)",
        r"station\s*(?:index|number|no\.?)\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9_-]*)",
    ]

    station_index = None

    for pattern in index_patterns:
        match = re.search(pattern, normalized_text, re.I)
        if match:
            station_index = match.group(1).strip()
            break

    # Handle label/value split across DOM/text lines.
    if not station_index:
        lines = [x.strip() for x in text.splitlines() if x.strip()]
        for i, line in enumerate(lines):
            if re.fullmatch(r"station\s*(?:index|number|no\.?)", line, re.I):
                if i + 1 < len(lines):
                    candidate = lines[i + 1].strip()
                    m = re.search(r"^[A-Za-z0-9][A-Za-z0-9_-]*$", candidate)
                    if m:
                        station_index = candidate
                        break

    # ----------------------------
    # Station Name
    # ----------------------------
    station_name = None

    # Typical DHM popup:
    # Station: Thakmarpha | District: MUSTANG
    name_patterns = [
        r"station\s*name\s*[:\-]\s*([^\n|]+)",
        r"(?:^|\n)station\s*[:\-]\s*([^\n|]+)",
    ]

    for pattern in name_patterns:
        match = re.search(pattern, normalized_text, re.I)
        if match:
            candidate = match.group(1).strip()
            if candidate:
                station_name = candidate
                break

    # Handle station name/value split over lines.
    if not station_name:
        lines = [x.strip() for x in text.splitlines() if x.strip()]
        for i, line in enumerate(lines):
            if re.fullmatch(r"station\s*name|station", line, re.I):
                if i + 1 < len(lines):
                    candidate = lines[i + 1].strip()
                    candidate = re.split(
                        r"\s*\|\s*(?:district|province|basin|river)\s*[:\-]",
                        candidate,
                        maxsplit=1,
                        flags=re.I,
                    )[0].strip()
                    if candidate:
                        station_name = candidate
                        break

    if not station_index:
        raise RuntimeError(
            f"Could not extract Station Index from the active popup for marker sequence {sequence}."
        )

    if not station_name:
        raise RuntimeError(
            f"Could not extract Station Name from the active popup for marker sequence {sequence}."
        )

    return station_name[:180], station_index


# ============================================================
# POPUP CONTROLS — POPUP ONLY
# ============================================================

def find_period_select(popup):
    """Find the native Period select INSIDE the active popup."""
    selects = popup.locator("select")

    for i in range(selects.count()):
        select = selects.nth(i)
        try:
            if not select.is_visible():
                continue

            options = select.locator("option")
            values = []
            for j in range(options.count()):
                option = options.nth(j)
                values.append((
                    option.inner_text().strip(),
                    (option.get_attribute("value") or "").strip(),
                ))

            if any(text.lower() == "daily" for text, _ in values):
                return select, values
        except Exception:
            pass

    return None, []


def set_period_daily(popup, timeout=15):
    """Set Period=Daily using controls INSIDE the active popup only."""
    deadline = time.time() + timeout

    while time.time() < deadline:
        select, options = find_period_select(popup)

        if select is not None:
            print(f"    Found Period <select> in popup: {options}")

            daily = next(
                (item for item in options if item[0].strip().lower() == "daily"),
                None,
            )

            if daily:
                value = daily[1] or "3"

                try:
                    select.select_option(value=value)
                except Exception:
                    select.select_option(label="Daily")

                popup.page.wait_for_timeout(250)

                try:
                    selected = select.locator("option:checked").inner_text().strip()
                    if selected.lower() == "daily":
                        print("    Period successfully set to Daily in popup.")
                        return
                except Exception:
                    pass

                # JS fallback on THIS popup select.
                try:
                    ok = select.evaluate(
                        """
                        (el) => {
                            const option = [...el.options].find(
                                o => o.textContent.trim().toLowerCase() === 'daily'
                            );
                            if (!option) return false;
                            el.value = option.value;
                            el.dispatchEvent(new Event('input', {bubbles:true}));
                            el.dispatchEvent(new Event('change', {bubbles:true}));
                            return el.value === option.value;
                        }
                        """
                    )
                    if ok:
                        popup.page.wait_for_timeout(250)
                        selected = select.locator("option:checked").inner_text().strip()
                        if selected.lower() == "daily":
                            print("    Period set to Daily via popup JavaScript.")
                            return
                except Exception:
                    pass

        # Custom dropdown fallback, still restricted to popup.
        for selector in [
            "[role='combobox']",
            "[aria-haspopup='listbox']",
            ".select2-selection",
            ".dropdown-toggle",
        ]:
            try:
                candidates = popup.locator(selector)
                for i in range(candidates.count()):
                    candidate = candidates.nth(i)
                    if not candidate.is_visible():
                        continue

                    label = (
                        candidate.inner_text().strip()
                        + " "
                        + str(candidate.get_attribute("aria-label") or "")
                    ).lower()

                    if not any(x in label for x in ["period", "hour", "daily"]):
                        continue

                    candidate.click(force=True)
                    popup.page.wait_for_timeout(200)

                    options = popup.locator(
                        "[role='option'], [role='menuitem'], .dropdown-item, li"
                    )
                    for j in range(options.count()):
                        option = options.nth(j)
                        try:
                            if option.is_visible() and option.inner_text().strip().lower() == "daily":
                                option.click(force=True)
                                popup.page.wait_for_timeout(300)
                                print("    Period set to Daily using popup custom dropdown.")
                                return
                        except Exception:
                            pass
            except Exception:
                pass

        popup.page.wait_for_timeout(400)

    raise RuntimeError("Period dropdown was not found or could not be set to Daily inside the popup.")


def click_view(popup, timeout=10):
    """Click VIEW INSIDE the active popup only."""
    deadline = time.time() + timeout

    while time.time() < deadline:
        selectors = [
            "button",
            "input[type='button']",
            "input[type='submit']",
            "input",
            "a",
            "[role='button']",
            ".btn",
            ".button",
        ]

        for selector in selectors:
            try:
                candidates = popup.locator(selector)
                for i in range(candidates.count()):
                    element = candidates.nth(i)
                    if not element.is_visible():
                        continue

                    text = " ".join([
                        element.inner_text().strip() if selector != "input" else "",
                        str(element.get_attribute("value") or "").strip(),
                        str(element.get_attribute("aria-label") or "").strip(),
                        str(element.get_attribute("title") or "").strip(),
                    ]).strip()

                    if text.upper() != "VIEW":
                        continue

                    print(f"    Found VIEW button inside popup: text={text!r}")

                    try:
                        element.scroll_into_view_if_needed(timeout=3000)
                    except Exception:
                        pass

                    try:
                        element.click(timeout=5000)
                    except Exception:
                        element.click(force=True, timeout=5000)

                    popup.page.wait_for_timeout(500)
                    print("    VIEW clicked successfully inside popup.")
                    return
            except Exception:
                pass

        popup.page.wait_for_timeout(300)

    raise RuntimeError("VIEW button was not found inside the active popup.")


# ============================================================
# TABLE — POPUP ONLY
# ============================================================

def read_daily_table(popup):
    """Read Date/Daily/Total table strictly from the active popup."""
    tables = popup.locator("table")

    for i in range(tables.count()):
        table = tables.nth(i)

        try:
            if not table.is_visible():
                continue

            rows = table.locator("tr")
            if rows.count() < 2:
                continue

            header_index = None
            date_col = None
            daily_col = None
            total_col = None

            for r in range(min(rows.count(), 8)):
                cells = rows.nth(r).locator("th, td")
                values = [x.strip() for x in cells.all_inner_texts()]
                normalized = [re.sub(r"\s+", " ", x).lower() for x in values]

                d = next((c for c, x in enumerate(normalized) if x == "date" or x.startswith("date ")), None)
                daily = next((c for c, x in enumerate(normalized) if x == "daily" or "daily" in x), None)
                total = next((c for c, x in enumerate(normalized) if x == "total" or "total" in x), None)

                if d is not None and daily is not None:
                    header_index = r
                    date_col = d
                    daily_col = daily
                    total_col = total
                    break

            if header_index is None:
                continue

            extracted = []

            for r in range(header_index + 1, rows.count()):
                cells = rows.nth(r).locator("th, td")
                values = [x.strip() for x in cells.all_inner_texts()]

                needed = max(date_col, daily_col, total_col if total_col is not None else 0)
                if len(values) <= needed:
                    continue

                date_value = parse_date(values[date_col])
                daily_value = parse_number(values[daily_col])
                total_value = parse_number(values[total_col]) if total_col is not None else None

                if date_value is None or daily_value is None:
                    continue

                extracted.append((date_value, daily_value, total_value))

            if extracted:
                print(
                    f"    Found Daily rainfall table inside popup: {len(extracted)} row(s)."
                )
                return extracted

        except Exception:
            pass

    return []


# ============================================================
# STATION RAINFALL WORKFLOW — POPUP ONLY
# ============================================================

def extract_popup_rainfall(popup):
    """
    Complete station workflow INSIDE THE ACTIVE POPUP:

      Period -> Daily
      VIEW
      read Date/Daily/Total table
      remove today's row
      keep every other returned row
    """
    page = popup.page
    today = today_iso()

    set_period_daily(popup)
    click_view(popup)

    print("    Waiting for Daily rainfall table inside popup...")

    deadline = time.time() + 15

    while time.time() < deadline:
        # Reacquire the active popup because DHM may replace its inner DOM
        # after VIEW is clicked.
        current_popup = get_active_popup(page)
        if current_popup is not None:
            popup = current_popup

        rows = read_daily_table(popup)

        if rows:
            filtered = [row for row in rows if row[0] != today]

            print(f"    Today's row excluded: {today}")
            print(f"    Historical rows retained: {len(filtered)}")

            if filtered:
                return filtered

        page.wait_for_timeout(500)

    raise RuntimeError(
        "Daily rainfall table was not found inside the active popup, "
        "or it contained no rows other than today's date."
    )


# ============================================================
# EXCEL
# ============================================================

def safe_filename(station_index, station_name):
    clean = re.sub(r"[^\w\-]+", "_", station_name).strip("_")
    if not clean:
        clean = "Unknown_Station"
    return f"{station_index}_{clean}.xlsx"


def existing_dates(path):
    if not path.exists():
        return set()

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    dates = set()

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        value = row[0]
        if isinstance(value, dt.datetime):
            dates.add(value.strftime("%Y-%m-%d"))
        elif isinstance(value, dt.date):
            dates.add(value.strftime("%Y-%m-%d"))
        else:
            parsed = parse_date(str(value))
            if parsed:
                dates.add(parsed)

    wb.close()
    return dates


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
            "Total (mm)",
        ])

    existing = existing_dates(path)
    added = 0

    for date_value, daily, total in sorted(rows, key=lambda x: x[0]):
        if date_value in existing:
            continue

        ws.append([date_value, daily, total])
        existing.add(date_value)
        added += 1

    for cell in ws["A"][1:]:
        if isinstance(cell.value, str):
            try:
                cell.value = dt.datetime.strptime(cell.value, "%Y-%m-%d").date()
            except ValueError:
                pass
        if isinstance(cell.value, dt.date):
            cell.number_format = "yyyy-mm-dd"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)

    return added


# ============================================================
# PROCESS ONE MARKER
# ============================================================

def process_marker(page, marker, sequence):
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass

    try:
        marker.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass

    try:
        marker.click(force=True, timeout=5000)
    except Exception:
        try:
            marker.dispatch_event("click")
        except Exception as exc:
            raise RuntimeError(f"Could not click Leaflet marker: {exc}")

    # Wait for popup.
    popup = wait_for_popup(page)

    # IMPORTANT: station name + index are extracted ONLY from popup.
    try:
        wait_for_station_info(popup)
    except Exception:
        save_debug(page, f"marker_{sequence:04d}_station_info_wait", popup)
        raise

    station_name, station_index = extract_station_info(popup, sequence)

    print(f"    popup station: {station_name}")
    print(f"    actual Station Index: {station_index}")

    save_debug(page, f"marker_{sequence:04d}_before_daily", popup)

    # IMPORTANT: Period, VIEW and table are all handled through this popup.
    rows = extract_popup_rainfall(popup)

    popup_after = get_active_popup(page)
    save_debug(page, f"marker_{sequence:04d}_after_daily", popup_after or popup)

    return station_name, station_index, rows


# ============================================================
# MAIN
# ============================================================

def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print("======================================")
    print("DHM Daily Rainfall Extractor")
    print("======================================")
    print(f"Today: {today_iso()}")
    print(f"Yesterday: {yesterday_iso()}")
    print("Station name/index: ACTIVE LEAFLET POPUP ONLY")
    print("Period: ACTIVE LEAFLET POPUP ONLY")
    print("VIEW: ACTIVE LEAFLET POPUP ONLY")
    print("Rainfall table: ACTIVE LEAFLET POPUP ONLY")
    print("Today's row will be excluded; all other Daily rows are saved.")
    print("======================================")

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
        )

        page = context.new_page()

        try:
            click_24_hour(page)

            markers = wait_for_markers(page)
            total = markers.count()
            print(f"Found {total} Leaflet station marker(s).")

            if DEBUG:
                total_to_process = min(total, DEBUG_MAX_STATIONS)
            elif MAX_STATIONS:
                total_to_process = min(total, MAX_STATIONS)
            else:
                total_to_process = total

            print(f"Will process {total_to_process} station(s).")

            successful = 0
            failed = 0

            for sequence in range(1, total_to_process + 1):
                print()
                print("--------------------------------------")
                print(f"[{sequence}/{total_to_process}] Station marker sequence {sequence}")
                print("Clicking station marker...")

                try:
                    # Reacquire because the DOM can change after each popup/AJAX request.
                    markers = marker_locator(page)
                    if markers is None:
                        raise RuntimeError("Leaflet marker locator disappeared.")

                    count = markers.count()
                    if sequence > count:
                        raise RuntimeError(
                            f"Marker count changed: {count}; cannot access marker {sequence}."
                        )

                    marker = markers.nth(sequence - 1)

                    station_name, station_index, rows = process_marker(
                        page, marker, sequence
                    )

                    filename = safe_filename(station_index, station_name)
                    path = OUTPUT_DIR / filename
                    added = append_rows(path, rows)

                    print(f"    Rainfall rows retained: {len(rows)}")
                    print(f"    New rows written: {added}")
                    print(f"    Excel: {path}")

                    successful += 1

                except Exception as exc:
                    failed += 1
                    print(f"    FAILED: {exc}")
                    save_debug(page, f"ERROR_marker_{sequence:04d}", get_active_popup(page))

                finally:
                    try:
                        close = page.locator(".leaflet-popup-close-button").last
                        if close.is_visible():
                            close.click(force=True)
                            page.wait_for_timeout(300)
                    except Exception:
                        pass

                time.sleep(0.5)

            print()
            print("======================================")
            print("Finished")
            print("======================================")
            print(f"Successful stations: {successful}")
            print(f"Failed stations: {failed}")
            print(f"Output folder: {OUTPUT_DIR.resolve()}")
            print("======================================")

        finally:
            page.close()
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
