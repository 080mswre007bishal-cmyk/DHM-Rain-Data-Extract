
"""
DHM Rainfall Watch Map -> click each station popup -> Daily rainfall -> Excel

IMPORTANT:
The DHM map popup itself contains the station information and the rainfall
controls. Therefore this script does NOT try to discover station IDs or
rainfallSingle URLs.

Workflow:
1. Open https://dhm.gov.np/hydrology/rainfall-watch-map
2. Click "24 Hour".
3. Find the Leaflet station markers.
4. Click each marker ONE BY ONE.
5. Read the station name/details from the popup.
6. Inside that popup, set:
       Date = yesterday (Nepal time)
       Period = Daily
   then click VIEW.
7. Read the Daily rainfall table shown by that popup.
8. Save/append the rows to an Excel file named after the station.
9. Move to the next marker.

Install:
    pip install playwright openpyxl
    playwright install chromium

Test locally first:
    MAX_STATIONS=3 DEBUG=1 HEADLESS=0 python main.py

For GitHub Actions:
    MAX_STATIONS=3 DEBUG=1 HEADLESS=1 python main.py

MAX_STATIONS=0 means all stations.
"""

import datetime as dt
import os
import re
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright


MAP_URL = "https://dhm.gov.np/hydrology/rainfall-watch-map"

OUTPUT_DIR = Path("output")
DEBUG_DIR = OUTPUT_DIR / "debug"

DEBUG = os.getenv("DEBUG", "0") == "1"
HEADLESS = os.getenv("HEADLESS", "1") != "0"
MAX_STATIONS = int(os.getenv("MAX_STATIONS", "0"))  # 0 = all


# ============================================================
# Nepal date
# ============================================================

NPT = dt.timezone(dt.timedelta(hours=5, minutes=45))


def yesterday_npt():
    return dt.datetime.now(NPT).date() - dt.timedelta(days=1)


def yesterday_iso():
    return yesterday_npt().strftime("%Y-%m-%d")


# ============================================================
# Parsing
# ============================================================

DATE_FORMATS = [
    "%a, %b %d, %Y, %I:%M %p",
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
        r"[A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4}(?:,\s+\d{1,2}:\d{2}\s*[AP]M)?",
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

    # Ignore empty / dash / non-numeric cells.
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
# Debug
# ============================================================

def save_debug(page, tag):
    if not DEBUG:
        return

    DEBUG_DIR.mkdir(parents=True, exist_ok=True)

    safe = re.sub(r"[^\w.-]+", "_", tag)

    try:
        page.screenshot(
            path=str(DEBUG_DIR / f"{safe}.png"),
            full_page=True,
        )
    except Exception:
        pass

    try:
        (DEBUG_DIR / f"{safe}.html").write_text(
            page.content(),
            encoding="utf-8",
        )
    except Exception:
        pass

    try:
        (DEBUG_DIR / f"{safe}.txt").write_text(
            page.locator("body").inner_text(),
            encoding="utf-8",
        )
    except Exception:
        pass


# ============================================================
# Map
# ============================================================

def click_24_hour(page):
    print("Opening DHM Rainfall Watch Map...")

    page.goto(
        MAP_URL,
        wait_until="domcontentloaded",
        timeout=60000,
    )

    # Wait for the page JavaScript to initialise.
    page.wait_for_timeout(3000)

    candidates = [
        page.get_by_text("24 Hour", exact=True).first,
        page.get_by_role("button", name="24 Hour", exact=True).first,
        page.locator("a:has-text('24 Hour')").first,
        page.locator("button:has-text('24 Hour')").first,
    ]

    clicked = False

    for locator in candidates:
        try:
            locator.wait_for(state="visible", timeout=10000)
            locator.click(force=True)
            clicked = True
            break
        except Exception:
            pass

    if not clicked:
        raise RuntimeError("Could not click the DHM '24 Hour' button.")

    # Give Leaflet enough time to redraw the 24-hour markers.
    page.wait_for_timeout(4000)

    try:
        page.locator(".leaflet-container").wait_for(
            state="visible",
            timeout=15000,
        )
    except Exception:
        pass

    print("24 Hour map loaded.")


def get_marker_locator(page):
    """
    DHM uses Leaflet. We deliberately use ONLY .leaflet-marker-icon.

    Do not combine .leaflet-marker-icon with .leaflet-interactive:
    that can count different elements belonging to the same station.
    """

    selectors = [
        "#map .leaflet-marker-icon",
        ".leaflet-container .leaflet-marker-icon",
    ]

    for selector in selectors:
        locator = page.locator(selector)

        try:
            if locator.count() > 0:
                return locator
        except Exception:
            pass

    return None


def marker_count(page):
    locator = get_marker_locator(page)
    return locator.count() if locator else 0


# ============================================================
# Popup helpers
# ============================================================

def visible_popup(page):
    """
    Return the last visible Leaflet popup.

    We intentionally do not require a station ID or station URL.
    """

    popups = page.locator(".leaflet-popup")

    visible = []

    for i in range(popups.count()):
        popup = popups.nth(i)

        try:
            if popup.is_visible():
                visible.append(popup)
        except Exception:
            pass

    return visible[-1] if visible else None


def station_name_from_popup(popup, sequence):
    """
    Extract the station name from the popup.

    First look for explicit labels such as:
        Station Name
        Station
        Name

    If DHM changes the labels, fall back to the first meaningful text line.
    """

    try:
        text = popup.inner_text()
    except Exception:
        return f"Station_{sequence:04d}"

    lines = [
        re.sub(r"\s+", " ", x).strip()
        for x in text.splitlines()
        if x.strip()
    ]

    # Look for:
    # Station Name: ABC
    # Station: ABC
    # Name: ABC
    label_pattern = re.compile(
        r"^(?:station\s*name|station|name)\s*[:\-]\s*(.+)$",
        re.I,
    )

    for line in lines:
        match = label_pattern.match(line)
        if match:
            value = match.group(1).strip()
            if value:
                return value[:180]

    # Sometimes the label and value are separate lines:
    # Station Name
    # ABC
    for i, line in enumerate(lines[:-1]):
        if re.fullmatch(
            r"(?:station\s*name|station|name)",
            line,
            re.I,
        ):
            candidate = lines[i + 1].strip()

            if candidate and not re.search(
                r"rainfall|daily|total|date|period|view",
                candidate,
                re.I,
            ):
                return candidate[:180]

    # Fallback: first useful non-numeric line.
    ignored = {
        "VIEW",
        "View",
        "DAILY",
        "Daily",
        "Date",
        "Period",
        "Rainfall",
        "Accumulated Rainfall",
        "24 Hour",
    }

    for line in lines:
        if line in ignored:
            continue

        if re.fullmatch(r"[\d.,\-\s]+", line):
            continue

        if len(line) >= 2:
            return line[:180]

    return f"Station_{sequence:04d}"


# ============================================================
# Date / Period / VIEW INSIDE POPUP
# ============================================================

def popup_date_input(popup):
    inputs = popup.locator("input")

    candidates = []

    for i in range(inputs.count()):
        inp = inputs.nth(i)

        try:
            attrs = inp.evaluate(
                """e => ({
                    type: e.type || '',
                    id: e.id || '',
                    name: e.name || '',
                    placeholder: e.placeholder || '',
                    aria: e.getAttribute('aria-label') || ''
                })"""
            )

            blob = " ".join(
                str(v).lower()
                for v in attrs.values()
            )

            score = 0

            if attrs["type"].lower() == "date":
                score += 100

            if "date" in blob:
                score += 30

            candidates.append((score, i))

        except Exception:
            pass

    candidates.sort(reverse=True)

    if candidates and candidates[0][0] > 0:
        return inputs.nth(candidates[0][1])

    return None


def set_popup_date_yesterday(popup):
    target = yesterday_iso()

    inp = popup_date_input(popup)

    if inp is None:
        return False

    # Normal Playwright fill.
    try:
        inp.fill(target, timeout=3000)
    except Exception:
        pass

    # Framework-safe DOM value setter.
    try:
        inp.evaluate(
            """(e, value) => {
                const setter =
                  Object.getOwnPropertyDescriptor(
                    HTMLInputElement.prototype,
                    'value'
                  ).set;

                setter.call(e, value);

                e.dispatchEvent(
                    new Event('input', {bubbles: true})
                );

                e.dispatchEvent(
                    new Event('change', {bubbles: true})
                );
            }""",
            target,
        )
    except Exception:
        pass

    try:
        actual = inp.input_value()
        return actual == target
    except Exception:
        return False


def set_popup_period_daily(popup):
    selects = popup.locator("select")

    for i in range(selects.count()):
        select = selects.nth(i)

        try:
            options = select.locator("option").all_inner_texts()

            if any(
                option.strip().lower() == "daily"
                for option in options
            ):
                select.select_option(label="Daily")
                return True

        except Exception:
            pass

    # Some versions may use a custom dropdown.
    for selector in [
        "[value='daily']",
        "option:has-text('Daily')",
    ]:
        try:
            popup.locator(selector).first.click(timeout=2000)
            return True
        except Exception:
            pass

    return False


def click_popup_view(popup):
    candidates = [
        popup.get_by_role(
            "button",
            name="VIEW",
            exact=True,
        ).first,

        popup.locator(
            "button:has-text('VIEW')"
        ).first,

        popup.locator(
            "input[value='VIEW']"
        ).first,

        popup.locator(
            "a:has-text('VIEW')"
        ).first,

        popup.get_by_text(
            "VIEW",
            exact=True,
        ).first,
    ]

    for locator in candidates:
        try:
            locator.wait_for(
                state="visible",
                timeout=2000,
            )

            locator.click(force=True)

            return True

        except Exception:
            pass

    return False


# ============================================================
# Read table INSIDE popup
# ============================================================

def read_rainfall_table_from_popup(popup):
    tables = popup.locator("table")

    for i in range(tables.count()):
        table = tables.nth(i)

        try:
            text = table.inner_text()
            low = text.lower()

            if "date" not in low:
                continue

            if "daily" not in low:
                continue

            rows = table.locator("tr")
            extracted = []

            for r in range(rows.count()):
                cells = rows.nth(r).locator("th, td")

                values = [
                    x.strip()
                    for x in cells.all_inner_texts()
                ]

                if len(values) < 2:
                    continue

                date_value = parse_date(values[0])

                if not date_value:
                    continue

                daily_value = parse_number(values[1])

                if daily_value is None:
                    continue

                total_value = (
                    parse_number(values[2])
                    if len(values) >= 3
                    else None
                )

                extracted.append(
                    (
                        date_value,
                        daily_value,
                        total_value,
                    )
                )

            if extracted:
                return extracted

        except Exception:
            pass

    return []


def extract_popup_rainfall(popup):
    """
    Complete operation on one clicked station popup:

        yesterday -> Daily -> VIEW -> table
    """

    date_ok = set_popup_date_yesterday(popup)

    if not date_ok:
        raise RuntimeError(
            "Date input was not found or could not be set to yesterday."
        )

    daily_ok = set_popup_period_daily(popup)

    if not daily_ok:
        raise RuntimeError(
            "Period dropdown was not found or could not be set to Daily."
        )

    if not click_popup_view(popup):
        raise RuntimeError(
            "VIEW button was not found inside the station popup."
        )

    # AJAX table rendering.
    for _ in range(20):
        rows = read_rainfall_table_from_popup(popup)

        if rows:
            return rows

        time.sleep(0.5)

    return []


# ============================================================
# Excel
# ============================================================

def safe_filename(name, sequence):
    clean = re.sub(
        r"[^\w\-]+",
        "_",
        name,
    ).strip("_")

    if not clean:
        clean = f"Station_{sequence:04d}"

    return f"{sequence:04d}_{clean}.xlsx"


def existing_dates(path):
    if not path.exists():
        return set()

    wb = load_workbook(
        path,
        read_only=True,
        data_only=True,
    )

    ws = wb.active
    dates = set()

    for row in ws.iter_rows(
        min_row=2,
        values_only=True,
    ):
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
    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    if path.exists():
        wb = load_workbook(path)
        ws = wb.active

    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Daily Rainfall"

        ws.append(
            [
                "Date",
                "Daily Rainfall (mm)",
                "Running Total on page (mm)",
            ]
        )

    existing = existing_dates(path)

    added = 0

    for date_value, daily, total in sorted(rows):
        if date_value in existing:
            continue

        ws.append(
            [
                date_value,
                daily,
                total,
            ]
        )

        existing.add(date_value)
        added += 1

    # Convert Date column to real Excel dates.
    for cell in ws["A"][1:]:
        if isinstance(cell.value, str):
            try:
                cell.value = dt.datetime.strptime(
                    cell.value,
                    "%Y-%m-%d",
                ).date()

                cell.number_format = "yyyy-mm-dd"

            except ValueError:
                pass

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(path)

    return added


# ============================================================
# Process ONE marker
# ============================================================

def process_marker(page, marker, sequence):
    """
    Click exactly one Leaflet marker and work entirely from its popup.

    No station ID is required.
    """

    # Close any previous popup.
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass

    try:
        marker.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass

    # Normal click first.
    clicked = False

    try:
        marker.click(
            force=True,
            timeout=5000,
        )
        clicked = True
    except Exception:
        pass

    if not clicked:
        try:
            marker.dispatch_event("click")
            clicked = True
        except Exception:
            pass

    if not clicked:
        raise RuntimeError(
            "Could not click marker."
        )

    # Wait for popup.
    popup = None

    for _ in range(20):
        popup = visible_popup(page)

        if popup is not None:
            break

        time.sleep(0.25)

    if popup is None:
        raise RuntimeError(
            "Marker was clicked but no Leaflet popup appeared."
        )

    station_name = station_name_from_popup(
        popup,
        sequence,
    )

    if DEBUG:
        save_debug(
            page,
            f"marker_{sequence:04d}_before_view",
        )

    print(
        f"    popup station: {station_name}"
    )

    rows = extract_popup_rainfall(popup)

    if DEBUG:
        save_debug(
            page,
            f"marker_{sequence:04d}_after_view",
        )

    if not rows:
        raise RuntimeError(
            "VIEW completed but no Daily rainfall table "
            "was found inside the popup."
        )

    return station_name, rows


# ============================================================
# Main
# ============================================================

def main():
    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        f"Yesterday in Nepal: {yesterday_iso()}"
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
            viewport={
                "width": 1440,
                "height": 1000,
            },
            locale="en-US",
            timezone_id="Asia/Kathmandu",
        )

        page = context.new_page()

        try:
            # --------------------------------------------------
            # Map
            # --------------------------------------------------

            click_24_hour(page)

            # Wait until Leaflet markers appear.
            marker_locator = None

            for _ in range(40):
                marker_locator = get_marker_locator(page)

                if marker_locator and marker_locator.count() > 0:
                    break

                page.wait_for_timeout(500)

            if marker_locator is None:
                raise RuntimeError(
                    "No Leaflet station markers were found."
                )

            total = marker_locator.count()

            print(
                f"Found about {total} map station marker(s)."
            )

            if MAX_STATIONS:
                total_to_process = min(
                    total,
                    MAX_STATIONS,
                )
            else:
                total_to_process = total

            print(
                f"Will process {total_to_process} station(s)."
            )

            # --------------------------------------------------
            # Station-by-station popup extraction
            # --------------------------------------------------

            successful = 0
            failed = 0

            for sequence in range(
                1,
                total_to_process + 1,
            ):

                print(
                    f"\n[{sequence}/{total_to_process}] "
                    f"Clicking station marker..."
                )

                try:
                    # Re-create the locator each time because the DOM can
                    # change after a popup AJAX update.
                    marker_locator = get_marker_locator(page)

                    if marker_locator is None:
                        raise RuntimeError(
                            "Leaflet marker locator disappeared."
                        )

                    if sequence > marker_locator.count():
                        raise RuntimeError(
                            "Marker count changed while processing."
                        )

                    marker = marker_locator.nth(
                        sequence - 1
                    )

                    station_name, rows = process_marker(
                        page,
                        marker,
                        sequence,
                    )

                    filename = safe_filename(
                        station_name,
                        sequence,
                    )

                    path = OUTPUT_DIR / filename

                    added = append_rows(
                        path,
                        rows,
                    )

                    print(
                        f"    rainfall rows: {len(rows)}"
                    )

                    print(
                        f"    added: {added} new row(s)"
                    )

                    print(
                        f"    Excel: {path}"
                    )

                    successful += 1

                except Exception as exc:

                    failed += 1

                    print(
                        f"    FAILED: {exc}"
                    )

                    save_debug(
                        page,
                        f"ERROR_marker_{sequence:04d}",
                    )

                finally:
                    # Close popup before moving to the next marker.
                    try:
                        close_button = page.locator(
                            ".leaflet-popup-close-button"
                        ).last

                        if close_button.is_visible():
                            close_button.click(
                                force=True
                            )
                            page.wait_for_timeout(
                                250
                            )

                    except Exception:
                        pass

                # Be polite to DHM.
                time.sleep(0.4)

            print(
                "\n======================================"
            )

            print(
                "Finished."
            )

            print(
                f"Successful stations: {successful}"
            )

            print(
                f"Failed stations: {failed}"
            )

            print(
                f"Output folder: {OUTPUT_DIR.resolve()}"
            )

        finally:
            page.close()
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
