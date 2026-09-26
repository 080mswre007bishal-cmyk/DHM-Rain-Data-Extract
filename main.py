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

# 0 = process all stations
MAX_STATIONS = int(os.getenv("MAX_STATIONS", "0"))


# ============================================================
# DATE
# ============================================================

def yesterday():
    """
    Return yesterday according to the normal AD/Gregorian
    calendar of the machine running the script.

    No Nepal timezone conversion is performed.
    """
    return dt.date.today() - dt.timedelta(days=1)


def yesterday_iso():
    """
    Return yesterday as YYYY-MM-DD.
    """
    return yesterday().strftime("%Y-%m-%d")


# ============================================================
# DATE PARSING
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
    """
    Convert a date string from the DHM table into YYYY-MM-DD.
    """

    if not text:
        return None

    text = " ".join(str(text).split())

    # Try complete known formats first.
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(
                text,
                fmt,
            ).strftime("%Y-%m-%d")
        except ValueError:
            pass

    # Try extracting a date from a larger string.
    patterns = [
        r"[A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4}(?:,\s+\d{1,2}:\d{2}\s*[AP]M)?",
        r"\d{4}-\d{2}-\d{2}",
        r"\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{4}",
        r"\d{1,2}/\d{1,2}/\d{4}",
        r"\d{1,2}-\d{1,2}-\d{4}",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
        )

        if not match:
            continue

        token = match.group(0)

        for fmt in DATE_FORMATS:

            try:
                return dt.datetime.strptime(
                    token,
                    fmt,
                ).strftime("%Y-%m-%d")

            except ValueError:
                pass

    return None


# ============================================================
# NUMBER PARSING
# ============================================================

def parse_number(text):
    """
    Convert rainfall text into a number.
    """

    if text is None:
        return None

    text = str(text).replace(
        ",",
        "",
    ).strip()

    # Ignore empty / dash / non-numeric cells.
    if not text or text in {
        "-",
        "—",
        "N/A",
        "NA",
    }:
        return None

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text,
    )

    if not match:
        return None

    try:
        return float(match.group(0))

    except ValueError:
        return None


# ============================================================
# DEBUG
# ============================================================

def save_debug(page, tag):
    """
    Save screenshot, HTML and text when DEBUG=1.
    """

    if not DEBUG:
        return

    DEBUG_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    safe = re.sub(
        r"[^\w.-]+",
        "_",
        tag,
    )

    # Screenshot
    try:

        page.screenshot(
            path=str(
                DEBUG_DIR / f"{safe}.png"
            ),
            full_page=True,
        )

    except Exception:
        pass

    # HTML
    try:

        (
            DEBUG_DIR / f"{safe}.html"
        ).write_text(
            page.content(),
            encoding="utf-8",
        )

    except Exception:
        pass

    # Visible text
    try:

        (
            DEBUG_DIR / f"{safe}.txt"
        ).write_text(
            page.locator("body").inner_text(),
            encoding="utf-8",
        )

    except Exception:
        pass


# ============================================================
# MAP
# ============================================================

def click_24_hour(page):
    """
    Open the DHM Rainfall Watch Map and click 24 Hour.
    """

    print(
        "Opening DHM Rainfall Watch Map..."
    )

    page.goto(
        MAP_URL,
        wait_until="domcontentloaded",
        timeout=60000,
    )

    # Give the page JavaScript time to initialize.
    page.wait_for_timeout(3000)

    candidates = [
        page.get_by_text(
            "24 Hour",
            exact=True,
        ).first,

        page.get_by_role(
            "button",
            name="24 Hour",
            exact=True,
        ).first,

        page.locator(
            "a:has-text('24 Hour')"
        ).first,

        page.locator(
            "button:has-text('24 Hour')"
        ).first,
    ]

    clicked = False

    for locator in candidates:

        try:

            locator.wait_for(
                state="visible",
                timeout=10000,
            )

            locator.click(
                force=True
            )

            clicked = True
            break

        except Exception:
            pass

    if not clicked:

        raise RuntimeError(
            "Could not click the DHM '24 Hour' button."
        )

    # Give Leaflet enough time to redraw.
    page.wait_for_timeout(4000)

    try:

        page.locator(
            ".leaflet-container"
        ).wait_for(
            state="visible",
            timeout=15000,
        )

    except Exception:
        pass

    print(
        "24 Hour map loaded."
    )


# ============================================================
# LEAFLET MARKERS
# ============================================================

def get_marker_locator(page):
    """
    DHM uses Leaflet.

    Only use .leaflet-marker-icon so that the same station
    is not accidentally counted twice.
    """

    selectors = [
        "#map .leaflet-marker-icon",
        ".leaflet-container .leaflet-marker-icon",
    ]

    for selector in selectors:

        locator = page.locator(
            selector
        )

        try:

            if locator.count() > 0:
                return locator

        except Exception:
            pass

    return None


def marker_count(page):
    locator = get_marker_locator(page)

    return (
        locator.count()
        if locator
        else 0
    )


# ============================================================
# POPUP
# ============================================================

def visible_popup(page):
    """
    Return the last visible Leaflet popup.
    """

    popups = page.locator(
        ".leaflet-popup"
    )

    visible = []

    for i in range(
        popups.count()
    ):

        popup = popups.nth(i)

        try:

            if popup.is_visible():
                visible.append(
                    popup
                )

        except Exception:
            pass

    return (
        visible[-1]
        if visible
        else None
    )


# ============================================================
# STATION NAME
# ============================================================

def station_name_from_popup(
    popup,
    sequence,
):
    """
    Extract station name from the popup.
    """

    try:

        text = popup.inner_text()

    except Exception:

        return (
            f"Station_{sequence:04d}"
        )

    lines = [
        re.sub(
            r"\s+",
            " ",
            x,
        ).strip()

        for x in text.splitlines()

        if x.strip()
    ]

    # --------------------------------------------------------
    # Look for:
    #
    # Station Name: ABC
    # Station: ABC
    # Name: ABC
    # --------------------------------------------------------

    label_pattern = re.compile(
        r"^(?:station\s*name|station|name)"
        r"\s*[:\-]\s*(.+)$",
        re.I,
    )

    for line in lines:

        match = label_pattern.match(
            line
        )

        if match:

            value = (
                match.group(1)
                .strip()
            )

            if value:
                return value[:180]

    # --------------------------------------------------------
    # Sometimes label and value are separate:
    #
    # Station Name
    # ABC
    # --------------------------------------------------------

    for i, line in enumerate(
        lines[:-1]
    ):

        if re.fullmatch(
            r"(?:station\s*name|station|name)",
            line,
            re.I,
        ):

            candidate = (
                lines[i + 1]
                .strip()
            )

            if (
                candidate
                and not re.search(
                    r"rainfall|daily|total|date|period|view",
                    candidate,
                    re.I,
                )
            ):

                return candidate[:180]

    # --------------------------------------------------------
    # Fallback: first useful non-numeric line.
    # --------------------------------------------------------

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

        if re.fullmatch(
            r"[\d.,\-\s]+",
            line,
        ):
            continue

        if len(line) >= 2:
            return line[:180]

    return (
        f"Station_{sequence:04d}"
    )


# ============================================================
# PERIOD
# ============================================================

def _visible(locator):
    """Safely check whether a Playwright locator is visible."""
    try:
        return locator.count() > 0 and locator.first.is_visible()
    except Exception:
        return False


def _try_select_daily(root, label=""):
    """
    Look for a real HTML <select> containing Daily.
    This is the most reliable method when DHM uses a native select.
    """
    selects = root.locator("select")

    for i in range(selects.count()):
        select = selects.nth(i)

        try:
            if not select.is_visible():
                continue

            options = select.locator("option").all()
            option_info = []

            for option in options:
                try:
                    option_info.append({
                        "text": option.inner_text().strip(),
                        "value": option.get_attribute("value"),
                    })
                except Exception:
                    pass

            daily = next(
                (
                    x for x in option_info
                    if x["text"].strip().lower() == "daily"
                    or str(x["value"] or "").strip().lower() == "daily"
                ),
                None,
            )

            if daily is None:
                continue

            print(
                f"    Found Period <select>"
                f"{' in ' + label if label else ''}: "
                f"{option_info}"
            )

            # First try normal Playwright selection.
            try:
                if daily["value"]:
                    select.select_option(value=daily["value"])
                else:
                    select.select_option(label="Daily")
            except Exception:
                select.select_option(label="Daily")

            # Verify.
            try:
                selected_text = select.locator(
                    "option:checked"
                ).inner_text().strip()

                selected_value = (
                    select.input_value().strip().lower()
                )

                if (
                    selected_text.lower() == "daily"
                    or selected_value == "daily"
                    or (
                        daily["value"]
                        and selected_value == str(daily["value"]).lower()
                    )
                ):
                    print("    Period successfully set to Daily.")
                    return True
            except Exception:
                pass

            # Some DHM pages require the native change event.
            try:
                select.evaluate(
                    """
                    (el, value) => {
                        el.value = value;
                        el.dispatchEvent(
                            new Event('input', {bubbles: true})
                        );
                        el.dispatchEvent(
                            new Event('change', {bubbles: true})
                        );
                    }
                    """,
                    daily["value"] or "daily",
                )

                time.sleep(0.2)

                selected_text = select.locator(
                    "option:checked"
                ).inner_text().strip()

                if selected_text.lower() == "daily":
                    print("    Period set to Daily via native change event.")
                    return True

            except Exception:
                pass

        except Exception:
            continue

    return False


def _try_custom_daily(root, label=""):
    """
    Handle Bootstrap/custom dropdowns where there is no usable <select>.
    """
    # First click an element that visibly represents the current Period.
    trigger_selectors = [
        "[aria-haspopup='listbox']",
        "[role='combobox']",
        ".select2-selection",
        ".selectpicker",
        ".dropdown-toggle",
        "button:has-text('Hourly')",
        "button:has-text('Daily')",
        "button:has-text('Period')",
        "[class*='select']",
        "[class*='dropdown']",
    ]

    for selector in trigger_selectors:
        try:
            loc = root.locator(selector)

            for i in range(min(loc.count(), 10)):
                item = loc.nth(i)

                if not item.is_visible():
                    continue

                text_value = (
                    item.inner_text().strip()
                    + " "
                    + str(item.get_attribute("aria-label") or "")
                    + " "
                    + str(item.get_attribute("title") or "")
                ).lower()

                # Avoid clicking unrelated map dropdowns.
                if not any(
                    word in text_value
                    for word in [
                        "period",
                        "hour",
                        "daily",
                        "observ",
                    ]
                ):
                    continue

                item.click(force=True)
                time.sleep(0.2)

                # Search globally from the root for the opened Daily option.
                daily_candidates = root.locator(
                    """
                    [role='option'],
                    [role='menuitem'],
                    .dropdown-menu a,
                    .dropdown-menu button,
                    .dropdown-item,
                    li,
                    option
                    """
                )

                for j in range(daily_candidates.count()):
                    candidate = daily_candidates.nth(j)

                    try:
                        if not candidate.is_visible():
                            continue

                        if candidate.inner_text().strip().lower() == "daily":
                            candidate.click(force=True)
                            time.sleep(0.2)
                            print(
                                f"    Custom Period dropdown set to Daily"
                                f"{' in ' + label if label else ''}."
                            )
                            return True
                    except Exception:
                        continue

        except Exception:
            continue

    # Direct Daily text fallback.
    for selector in [
        "[role='option']:has-text('Daily')",
        "[role='menuitem']:has-text('Daily')",
        ".dropdown-item:has-text('Daily')",
        "li:has-text('Daily')",
        "option:has-text('Daily')",
    ]:
        try:
            candidate = root.locator(selector).first

            if candidate.is_visible():
                candidate.click(force=True)
                time.sleep(0.2)
                print(
                    f"    Clicked Daily directly"
                    f"{' in ' + label if label else ''}."
                )
                return True

        except Exception:
            pass

    return False


def _debug_period_controls(root, label=""):
    """
    Print every visible select/combobox/button that may belong to Period.
    """
    if not DEBUG:
        return

    try:
        print(
            f"    DEBUG Period inspection"
            f"{' (' + label + ')' if label else ''}:"
        )

        elements = root.locator(
            "select, [role='combobox'], "
            "button, input, [aria-haspopup='listbox']"
        )

        for i in range(elements.count()):
            el = elements.nth(i)

            try:
                if not el.is_visible():
                    continue

                tag = el.evaluate("(e) => e.tagName")
                text_value = el.inner_text().strip() if tag not in {
                    "INPUT", "SELECT"
                } else ""

                print(
                    "      "
                    f"{tag} "
                    f"text={text_value!r} "
                    f"name={el.get_attribute('name')!r} "
                    f"id={el.get_attribute('id')!r} "
                    f"class={el.get_attribute('class')!r} "
                    f"value={el.get_attribute('value')!r}"
                )

            except Exception:
                pass

    except Exception:
        pass


def set_popup_period_daily(popup, page):
    """
    Set DHM Period = Daily.

    The Date field is deliberately NOT modified.

    DHM can render the station controls in different ways:
      1. directly inside the Leaflet popup,
      2. inside an iframe,
      3. as page-level controls associated with the popup,
      4. as a native <select>,
      5. as a custom Bootstrap/select2 dropdown.

    We try all of these and VERIFY that Daily is selected.
    """

    roots = [
        ("popup", popup),
        ("page", page),
    ]

    # ------------------------------------------------------------
    # Check if the popup contains an iframe.
    # ------------------------------------------------------------
    try:
        frames = popup.locator("iframe")

        for i in range(frames.count()):
            iframe = frames.nth(i)

            try:
                frame = iframe.content_frame

                if frame is not None:
                    roots.insert(0, (f"popup iframe {i}", frame))
            except Exception:
                pass

    except Exception:
        pass

    # ------------------------------------------------------------
    # DEBUG: show what Playwright can actually see.
    # ------------------------------------------------------------
    for label, root in roots:
        _debug_period_controls(root, label)

    # ------------------------------------------------------------
    # 1. Native <select>.
    # ------------------------------------------------------------
    for label, root in roots:
        if _try_select_daily(root, label):
            return True

    # ------------------------------------------------------------
    # 2. Custom dropdown.
    # ------------------------------------------------------------
    for label, root in roots:
        if _try_custom_daily(root, label):
            return True

    # ------------------------------------------------------------
    # 3. Last-resort JavaScript search.
    #
    # Find a visible SELECT anywhere in the document containing
    # an option whose text/value is Daily.
    # ------------------------------------------------------------
    try:
        result = page.evaluate(
            """
            () => {
                const selects = [...document.querySelectorAll("select")];

                for (const select of selects) {
                    if (!select.offsetParent) continue;

                    const option = [...select.options].find(
                        o =>
                            o.textContent.trim().toLowerCase() === "daily" ||
                            (o.value || "").trim().toLowerCase() === "daily"
                    );

                    if (!option) continue;

                    select.value = option.value;

                    select.dispatchEvent(
                        new Event("input", {bubbles: true})
                    );

                    select.dispatchEvent(
                        new Event("change", {bubbles: true})
                    );

                    return {
                        found: true,
                        value: select.value,
                        text: option.textContent.trim()
                    };
                }

                return {found: false};
            }
            """
        )

        if result.get("found") and (
            result.get("text", "").lower() == "daily"
            or result.get("value", "").lower() == "daily"
        ):
            print("    Period set to Daily using page JavaScript fallback.")
            return True

    except Exception:
        pass

    return False


# ============================================================
# VIEW
# ============================================================

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

            locator.click(
                force=True
            )

            return True

        except Exception:
            pass

    return False


# ============================================================
# READ DAILY RAINFALL TABLE
# ============================================================

def read_rainfall_table_from_popup(
    popup
):
    """
    Read all valid rainfall rows from the popup.
    """

    tables = popup.locator(
        "table"
    )

    for i in range(
        tables.count()
    ):

        table = tables.nth(i)

        try:

            text = table.inner_text()
            low = text.lower()

            # Make sure this is the relevant table.
            if "date" not in low:
                continue

            if "daily" not in low:
                continue

            rows = table.locator(
                "tr"
            )

            extracted = []

            for r in range(
                rows.count()
            ):

                cells = rows.nth(
                    r
                ).locator(
                    "th, td"
                )

                values = [
                    x.strip()
                    for x in cells.all_inner_texts()
                ]

                if len(values) < 2:
                    continue

                date_value = parse_date(
                    values[0]
                )

                if not date_value:
                    continue

                daily_value = parse_number(
                    values[1]
                )

                if daily_value is None:
                    continue

                total_value = (
                    parse_number(
                        values[2]
                    )
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


# ============================================================
# EXTRACT YESTERDAY
# ============================================================

def extract_popup_rainfall(
    popup,
    page,
):
    """
    Extract ONLY yesterday's rainfall.

    Workflow:

        1. Leave Date exactly as DHM opened it.
        2. Set Period = Daily.
        3. Click VIEW.
        4. Read the returned table.
        5. Determine yesterday using AD/Gregorian date.
        6. Keep only that row.
    """

    target_date = yesterday_iso()

    print(
        f"    Target AD date: {target_date}"
    )

    # --------------------------------------------------------
    # Period = Daily
    # --------------------------------------------------------

    daily_ok = (
        set_popup_period_daily(
            popup,
            page,
        )
    )

    if not daily_ok:

        raise RuntimeError(
            "Period dropdown was not found "
            "or could not be set to Daily."
        )

    # --------------------------------------------------------
    # VIEW
    # --------------------------------------------------------

    if not click_popup_view(
        popup
    ):

        raise RuntimeError(
            "VIEW button was not found "
            "inside the station popup."
        )

    # --------------------------------------------------------
    # Wait for AJAX table.
    # --------------------------------------------------------

    for _ in range(30):

        rows = (
            read_rainfall_table_from_popup(
                popup
            )
        )

        if rows:

            # ------------------------------------------------
            # Keep ONLY yesterday.
            # ------------------------------------------------

            yesterday_rows = [
                row
                for row in rows

                if row[0] == target_date
            ]

            if yesterday_rows:

                return yesterday_rows

        time.sleep(0.5)

    raise RuntimeError(
        "Daily rainfall table loaded, "
        f"but no row was found for "
        f"yesterday ({target_date})."
    )


# ============================================================
# EXCEL FILE NAME
# ============================================================

def safe_filename(
    station_index,
    station_name,
):
    """
    Create:

        stationindex_StationName.xlsx

    Examples:

        1_Tarakhola.xlsx
        2_Panchase.xlsx
        365_Tarakhola.xlsx
        366_Panchase.xlsx
    """

    clean = re.sub(
        r"[^\w\-]+",
        "_",
        station_name,
    ).strip("_")

    if not clean:
        clean = (
            "Unknown_Station"
        )

    return (
        f"{station_index}_"
        f"{clean}.xlsx"
    )


# ============================================================
# EXISTING EXCEL DATES
# ============================================================

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

        if (
            not row
            or row[0] is None
        ):
            continue

        value = row[0]

        if isinstance(
            value,
            dt.datetime,
        ):

            dates.add(
                value.strftime(
                    "%Y-%m-%d"
                )
            )

        elif isinstance(
            value,
            dt.date,
        ):

            dates.add(
                value.strftime(
                    "%Y-%m-%d"
                )
            )

        else:

            parsed = parse_date(
                str(value)
            )

            if parsed:
                dates.add(parsed)

    wb.close()

    return dates


# ============================================================
# APPEND TO EXCEL
# ============================================================

def append_rows(
    path,
    rows,
):

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    # --------------------------------------------------------
    # Existing workbook
    # --------------------------------------------------------

    if path.exists():

        wb = load_workbook(
            path
        )

        ws = wb.active

    # --------------------------------------------------------
    # New workbook
    # --------------------------------------------------------

    else:

        wb = Workbook()

        ws = wb.active

        ws.title = (
            "Daily Rainfall"
        )

        ws.append(
            [
                "Date",
                "Daily Rainfall (mm)",
                "Running Total on page (mm)",
            ]
        )

    existing = existing_dates(
        path
    )

    added = 0

    for (
        date_value,
        daily,
        total,
    ) in sorted(rows):

        if date_value in existing:
            continue

        ws.append(
            [
                date_value,
                daily,
                total,
            ]
        )

        existing.add(
            date_value
        )

        added += 1

    # --------------------------------------------------------
    # Convert Date column to actual Excel dates.
    # --------------------------------------------------------

    for cell in ws["A"][1:]:

        if isinstance(
            cell.value,
            str,
        ):

            try:

                cell.value = (
                    dt.datetime.strptime(
                        cell.value,
                        "%Y-%m-%d",
                    ).date()
                )

                cell.number_format = (
                    "yyyy-mm-dd"
                )

            except ValueError:
                pass

    ws.freeze_panes = "A2"

    ws.auto_filter.ref = (
        ws.dimensions
    )

    wb.save(path)

    return added


# ============================================================
# PROCESS ONE STATION
# ============================================================

def process_marker(
    page,
    marker,
    sequence,
):
    """
    Process one station marker.

    `sequence` is the station index and is also used
    in the Excel filename.
    """

    # --------------------------------------------------------
    # Close any previous popup.
    # --------------------------------------------------------

    try:

        page.keyboard.press(
            "Escape"
        )

    except Exception:
        pass

    # --------------------------------------------------------
    # Scroll marker into view.
    # --------------------------------------------------------

    try:

        marker.scroll_into_view_if_needed(
            timeout=2000
        )

    except Exception:
        pass

    # --------------------------------------------------------
    # Click marker.
    # --------------------------------------------------------

    clicked = False

    try:

        marker.click(
            force=True,
            timeout=5000,
        )

        clicked = True

    except Exception:
        pass

    # --------------------------------------------------------
    # Fallback click.
    # --------------------------------------------------------

    if not clicked:

        try:

            marker.dispatch_event(
                "click"
            )

            clicked = True

        except Exception:
            pass

    if not clicked:

        raise RuntimeError(
            "Could not click marker."
        )

    # --------------------------------------------------------
    # Wait for popup.
    # --------------------------------------------------------

    popup = None

    for _ in range(20):

        popup = visible_popup(
            page
        )

        if popup is not None:
            break

        time.sleep(0.25)

    if popup is None:

        raise RuntimeError(
            "Marker was clicked but "
            "no Leaflet popup appeared."
        )

    # --------------------------------------------------------
    # Station name.
    # --------------------------------------------------------

    station_name = (
        station_name_from_popup(
            popup,
            sequence,
        )
    )

    if DEBUG:

        save_debug(
            page,
            f"marker_{sequence:04d}_before_view",
        )

    print(
        f"    station index: {sequence}"
    )

    print(
        f"    popup station: {station_name}"
    )

    # --------------------------------------------------------
    # Extract yesterday.
    # --------------------------------------------------------

    rows = (
        extract_popup_rainfall(
            popup,
            page,
        )
    )

    if DEBUG:

        save_debug(
            page,
            f"marker_{sequence:04d}_after_view",
        )

    if not rows:

        raise RuntimeError(
            "VIEW completed but no "
            "Daily rainfall row was found "
            "for yesterday."
        )

    return (
        station_name,
        rows,
    )


# ============================================================
# MAIN
# ============================================================

def main():

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        "======================================"
    )

    print(
        "DHM Daily Rainfall Extractor"
    )

    print(
        "======================================"
    )

    print(
        f"Today's AD date: "
        f"{dt.date.today().strftime('%Y-%m-%d')}"
    )

    print(
        f"Yesterday's AD date: "
        f"{yesterday_iso()}"
    )

    print(
        "Date field will NOT be modified."
    )

    print(
        "DHM default date will be used."
    )

    print(
        "Period will be set to Daily."
    )

    print(
        "======================================"
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
        )

        page = context.new_page()

        try:

            # =================================================
            # MAP
            # =================================================

            click_24_hour(
                page
            )

            # =================================================
            # WAIT FOR MARKERS
            # =================================================

            marker_locator = None

            for _ in range(40):

                marker_locator = (
                    get_marker_locator(
                        page
                    )
                )

                if (
                    marker_locator
                    and marker_locator.count() > 0
                ):
                    break

                page.wait_for_timeout(
                    500
                )

            if marker_locator is None:

                raise RuntimeError(
                    "No Leaflet station markers "
                    "were found."
                )

            total = (
                marker_locator.count()
            )

            print(
                f"Found about {total} "
                f"map station marker(s)."
            )

            # =================================================
            # HOW MANY STATIONS?
            # =================================================

            if MAX_STATIONS:

                total_to_process = min(
                    total,
                    MAX_STATIONS,
                )

            else:

                total_to_process = total

            print(
                f"Will process "
                f"{total_to_process} station(s)."
            )

            print(
                f"Target date: "
                f"{yesterday_iso()}"
            )

            # =================================================
            # PROCESS STATIONS
            # =================================================

            successful = 0
            failed = 0

            for sequence in range(
                1,
                total_to_process + 1,
            ):

                print()
                print(
                    "--------------------------------------"
                )

                print(
                    f"[{sequence}/{total_to_process}] "
                    f"Station index {sequence}"
                )

                print(
                    "Clicking station marker..."
                )

                try:

                    # ------------------------------------------------
                    # Re-create marker locator every time because
                    # the DOM may change after AJAX updates.
                    # ------------------------------------------------

                    marker_locator = (
                        get_marker_locator(
                            page
                        )
                    )

                    if marker_locator is None:

                        raise RuntimeError(
                            "Leaflet marker locator "
                            "disappeared."
                        )

                    current_count = (
                        marker_locator.count()
                    )

                    if sequence > current_count:

                        raise RuntimeError(
                            "Marker count changed "
                            "while processing. "
                            f"Current count={current_count}, "
                            f"requested index={sequence}."
                        )

                    # ------------------------------------------------
                    # IMPORTANT:
                    #
                    # nth() is zero-based.
                    # sequence 1 -> nth(0)
                    # sequence 2 -> nth(1)
                    # etc.
                    # ------------------------------------------------

                    marker = (
                        marker_locator.nth(
                            sequence - 1
                        )
                    )

                    station_name, rows = (
                        process_marker(
                            page,
                            marker,
                            sequence,
                        )
                    )

                    # =================================================
                    # EXCEL FILE
                    # =================================================

                    filename = (
                        safe_filename(
                            sequence,
                            station_name,
                        )
                    )

                    path = (
                        OUTPUT_DIR
                        / filename
                    )

                    added = (
                        append_rows(
                            path,
                            rows,
                        )
                    )

                    print(
                        f"    rainfall rows: "
                        f"{len(rows)}"
                    )

                    print(
                        f"    added: "
                        f"{added} new row(s)"
                    )

                    print(
                        f"    Excel: "
                        f"{path}"
                    )

                    successful += 1

                except Exception as exc:

                    failed += 1

                    print(
                        f"    FAILED: {exc}"
                    )

                    # Save debugging information.
                    save_debug(
                        page,
                        f"ERROR_marker_{sequence:04d}",
                    )

                finally:

                    # =================================================
                    # CLOSE POPUP
                    # =================================================

                    try:

                        close_button = (
                            page.locator(
                                ".leaflet-popup-close-button"
                            ).last
                        )

                        if close_button.is_visible():

                            close_button.click(
                                force=True
                            )

                            page.wait_for_timeout(
                                250
                            )

                    except Exception:
                        pass

                # ----------------------------------------------------
                # Small delay between stations.
                # ----------------------------------------------------

                time.sleep(
                    0.4
                )

            # =================================================
            # SUMMARY
            # =================================================

            print()
            print(
                "======================================"
            )

            print(
                "Finished."
            )

            print(
                "======================================"
            )

            print(
                f"Target date: "
                f"{yesterday_iso()}"
            )

            print(
                f"Successful stations: "
                f"{successful}"
            )

            print(
                f"Failed stations: "
                f"{failed}"
            )

            print(
                f"Output folder: "
                f"{OUTPUT_DIR.resolve()}"
            )

            print(
                "======================================"
            )

        finally:

            page.close()
            context.close()
            browser.close()


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()