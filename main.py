import datetime as dt
import json
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

# GitHub Actions MUST remain headless.
HEADLESS = os.getenv("HEADLESS", "1") != "0"

# 0 = all stations
# Example:
# MAX_STATIONS=10 python main.py
MAX_STATIONS = int(os.getenv("MAX_STATIONS", "0"))

# How long to wait for the map/station interface.
MAP_TIMEOUT = 60
STATION_TIMEOUT = 20


# ============================================================
# TODAY
# ============================================================

def today_ad():
    """
    Ordinary machine AD/Gregorian date.

    This is ONLY used to remove today's row.

    The DHM Date field is NEVER modified.
    """
    return dt.date.today()


def today_iso():
    return today_ad().strftime("%Y-%m-%d")


# ============================================================
# DATE PARSING
# ============================================================

DATE_FORMATS = [
    "%a, %b %d, %Y, %I:%M %p",
    "%a, %b %d, %Y, %I:%M:%S %p",
    "%a, %b %d, %Y",
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%d %b, %Y",
    "%d %B, %Y",
    "%d %b %Y",
    "%d %B %Y",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%d-%m-%Y",
]


def parse_date(value):
    if value is None:
        return None

    text = " ".join(str(value).split()).strip()

    if not text:
        return None

    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(
                text,
                fmt,
            ).date()
        except ValueError:
            pass

    patterns = [
        r"[A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4}(?:,\s+\d{1,2}:\d{2}(?::\d{2})?\s*[AP]M)?",
        r"\d{4}-\d{2}-\d{2}",
        r"\d{4}/\d{2}/\d{2}",
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
                return dt.datetime.strptime(
                    token,
                    fmt,
                ).date()
            except ValueError:
                pass

    return None


# ============================================================
# NUMBER PARSING
# ============================================================

def parse_number(value):
    if value is None:
        return None

    text = str(value).replace(",", "").strip()

    if not text:
        return None

    if text.upper() in {
        "-",
        "—",
        "N/A",
        "NA",
        "NULL",
        "NONE",
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

def save_debug(page, name):
    if not DEBUG:
        return

    DEBUG_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    safe = re.sub(
        r"[^\w.-]+",
        "_",
        name,
    )

    try:
        page.screenshot(
            path=str(
                DEBUG_DIR / f"{safe}.png"
            ),
            full_page=True,
        )
    except Exception:
        pass

    try:
        (
            DEBUG_DIR / f"{safe}.html"
        ).write_text(
            page.content(),
            encoding="utf-8",
        )
    except Exception:
        pass

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
# OPEN MAP
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

        page.locator(
            "[role='tab']:has-text('24 Hour')"
        ).first,
    ]

    clicked = False

    for locator in candidates:
        try:
            locator.wait_for(
                state="visible",
                timeout=5000,
            )

            locator.click(
                force=True,
                timeout=5000,
            )

            clicked = True
            break

        except Exception:
            pass

    if not clicked:
        raise RuntimeError(
            "Could not click DHM 24 Hour tab."
        )

    print("24 Hour tab clicked.")

    # Wait for the DHM AJAX call / map redraw.
    page.wait_for_timeout(5000)

    try:
        page.locator(
            ".leaflet-container"
        ).wait_for(
            state="visible",
            timeout=15000,
        )
    except Exception:
        pass

    # Give station markers time to render.
    wait_for_markers(page)

    print("24 Hour map loaded.")


# ============================================================
# LEAFLET MARKERS
# ============================================================

def marker_locators(page):
    """
    Return possible rendered Leaflet station-marker locators.

    We do NOT replace Leaflet's L.map().

    We do NOT inspect private Leaflet internals.

    We simply use the markers that are actually rendered
    on the map.
    """

    return [
        page.locator(
            "#map .leaflet-marker-icon"
        ),

        page.locator(
            ".leaflet-container .leaflet-marker-icon"
        ),
    ]


def get_marker_locator(page):
    """
    Prefer the normal Leaflet marker-icon implementation.
    """

    for locator in marker_locators(page):
        try:
            count = locator.count()

            if count > 0:
                return locator

        except Exception:
            pass

    return None


def wait_for_markers(page):
    """
    Wait until DHM has actually rendered the station markers.
    """

    print("Waiting for station markers...")

    for attempt in range(60):

        locator = get_marker_locator(page)

        if locator is not None:

            count = locator.count()

            if count > 0:

                print(
                    f"Found {count} rendered Leaflet marker(s)."
                )

                return locator

        page.wait_for_timeout(500)

    # Debug information if the normal markers never appeared.
    if DEBUG:
        save_debug(
            page,
            "NO_MARKERS",
        )

    raise RuntimeError(
        "No rendered Leaflet station markers were found."
    )


# ============================================================
# CAPTURE MARKER POSITIONS
# ============================================================

def capture_markers(page):
    """
    Capture each marker's screen position BEFORE opening any
    station interface.

    This is important because the DOM can change after a
    station is opened.

    We keep the marker's actual screen coordinate and then
    perform a real mouse click at that coordinate.
    """

    locator = wait_for_markers(page)

    count = locator.count()

    markers = []

    print(
        f"Capturing {count} station marker position(s)..."
    )

    for i in range(count):

        marker = locator.nth(i)

        try:
            if not marker.is_visible():
                continue

            box = marker.bounding_box()

            if not box:
                continue

            x = box["x"] + box["width"] / 2
            y = box["y"] + box["height"] / 2

            # Ignore completely invalid coordinates.
            if x < 0 or y < 0:
                continue

            markers.append(
                {
                    "sequence": len(markers) + 1,
                    "x": x,
                    "y": y,
                    "width": box["width"],
                    "height": box["height"],
                }
            )

        except Exception:
            continue

    if not markers:
        raise RuntimeError(
            "Leaflet markers existed but their positions "
            "could not be captured."
        )

    print(
        f"Captured {len(markers)} clickable station position(s)."
    )

    return markers


# ============================================================
# STATION INTERFACE
# ============================================================

def element_is_inside_map(element):
    """
    Determine whether an element is inside Leaflet's map.

    The DHM station interface we want is OUTSIDE the map.
    """

    try:
        return element.evaluate(
            """
            el => {
                const map =
                    document.querySelector('.leaflet-container');

                return !!(
                    map &&
                    map.contains(el)
                );
            }
            """
        )
    except Exception:
        return False


def visible_outside_map_elements(page):
    """
    Return visible elements outside the Leaflet map.
    """

    elements = page.locator(
        "body *:visible"
    )

    result = []

    count = elements.count()

    # Avoid scanning an absurdly large DOM.
    # The useful station interface is normally fairly small.
    for i in range(count):

        element = elements.nth(i)

        try:

            if element_is_inside_map(element):
                continue

            result.append(element)

        except Exception:
            continue

    return result


def station_interface_candidates(page):
    """
    Find likely DHM station interfaces outside the Leaflet map.

    We deliberately do NOT use .leaflet-popup here.
    """

    selectors = [
        "[role='dialog']:visible",
        ".modal.show:visible",
        ".modal.in:visible",
        ".offcanvas.show:visible",
        ".offcanvas:visible",
        ".station-modal:visible",
        ".station-details:visible",
        ".station-detail:visible",
        ".rainfall-single:visible",
        ".rainfallSingle:visible",
        "#stationModal:visible",
        "#station-modal:visible",
    ]

    candidates = []

    for selector in selectors:

        loc = page.locator(selector)

        try:
            for i in range(loc.count()):

                item = loc.nth(i)

                if not element_is_inside_map(item):
                    candidates.append(item)

        except Exception:
            pass

    # --------------------------------------------------------
    # Also discover containers by their text.
    # --------------------------------------------------------

    for element in visible_outside_map_elements(page):

        try:

            text = element.inner_text(
                timeout=1000
            ).strip()

            if not text:
                continue

            low = text.lower()

            if (
                "station index" in low
                or "station name" in low
            ):

                candidates.append(element)

        except Exception:
            pass

    # --------------------------------------------------------
    # Remove duplicates and very deeply nested elements.
    # Prefer reasonably sized containers.
    # --------------------------------------------------------

    unique = []

    seen = set()

    for item in candidates:

        try:

            handle = item

            key = item.evaluate(
                "el => el"
            )

            # Playwright object itself isn't reliably hashable
            # across versions, so use its text + tag.
            identity = item.evaluate(
                """
                el => ({
                    tag: el.tagName,
                    id: el.id || '',
                    cls: el.className || '',
                    text: (el.innerText || '').slice(0, 500)
                })
                """
            )

            identity = json.dumps(
                identity,
                sort_keys=True,
                default=str,
            )

            if identity in seen:
                continue

            seen.add(identity)

            unique.append(item)

        except Exception:
            pass

    return unique


def find_station_interface(page):
    """
    Find the station-specific interface that appeared after
    clicking a marker.

    IMPORTANT:
    This is NOT the Leaflet popup.
    """

    # First wait briefly for DHM to create it.
    for _ in range(
        STATION_TIMEOUT * 2
    ):

        candidates = station_interface_candidates(
            page
        )

        # ----------------------------------------------------
        # Strong test:
        # interface contains Station Index.
        # ----------------------------------------------------

        for candidate in candidates:

            try:

                if not candidate.is_visible():
                    continue

                text = candidate.inner_text()

                if re.search(
                    r"station\s*index",
                    text,
                    re.I,
                ):
                    return candidate

            except Exception:
                pass

        # ----------------------------------------------------
        # Some DHM versions may display Station Name first.
        # ----------------------------------------------------

        for candidate in candidates:

            try:

                if not candidate.is_visible():
                    continue

                text = candidate.inner_text()

                if (
                    re.search(
                        r"station\s*name",
                        text,
                        re.I,
                    )
                    and (
                        "period" in text.lower()
                        or "view" in text.lower()
                        or "daily" in text.lower()
                    )
                ):
                    return candidate

            except Exception:
                pass

        page.wait_for_timeout(500)

    return None


# ============================================================
# STATION NAME / STATION INDEX
# ============================================================

def clean_station_name(value):
    value = " ".join(
        str(value).split()
    ).strip()

    value = re.sub(
        r"\s*\(\d+\)\s*$",
        "",
        value,
    ).strip()

    return value


def extract_labeled_value(text, labels):
    """
    Extract values from forms such as:

        Station Index: 0604

    or:

        Station Index
        0604
    """

    lines = [
        " ".join(line.split()).strip()
        for line in text.splitlines()
        if line.strip()
    ]

    # --------------------------------------------------------
    # Same-line form.
    # --------------------------------------------------------

    for label in labels:

        pattern = re.compile(
            rf"^{re.escape(label)}\s*[:\-]\s*(.+)$",
            re.I,
        )

        for line in lines:

            match = pattern.match(line)

            if match:
                value = match.group(1).strip()

                if value:
                    return value

    # --------------------------------------------------------
    # Label and value on separate lines.
    # --------------------------------------------------------

    for i, line in enumerate(lines):

        for label in labels:

            if re.fullmatch(
                re.escape(label),
                line,
                re.I,
            ):

                if i + 1 < len(lines):

                    value = lines[i + 1].strip()

                    if value:
                        return value

    return None


def extract_station_index(interface):
    """
    Extract the ACTUAL DHM Station Index.

    This is deliberately independent of the marker sequence.
    """

    text = interface.inner_text()

    # First use explicit Station Index label.
    value = extract_labeled_value(
        text,
        [
            "Station Index",
            "Station index",
            "Station Index No",
            "Station Index Number",
        ],
    )

    if value:

        # Station codes may contain leading zeroes.
        # Keep them as strings.
        match = re.search(
            r"\b\d{2,8}\b",
            value,
        )

        if match:
            return match.group(0)

    # --------------------------------------------------------
    # Search HTML text around "Station Index".
    # --------------------------------------------------------

    match = re.search(
        r"station\s*index"
        r"\s*(?:[:\-])?\s*"
        r"(?:</?\w+[^>]*>\s*)*"
        r"(\d{2,8})",
        interface.inner_html(),
        re.I,
    )

    if match:
        return match.group(1)

    # --------------------------------------------------------
    # Search visible text.
    # --------------------------------------------------------

    match = re.search(
        r"station\s*index"
        r"\s*[:\-]?\s*(\d{2,8})",
        text,
        re.I,
    )

    if match:
        return match.group(1)

    raise RuntimeError(
        "Actual Station Index was not found in the "
        "station interface."
    )


def extract_station_name(interface):
    text = interface.inner_text()

    value = extract_labeled_value(
        text,
        [
            "Station Name",
            "Station name",
            "Station",
            "Name",
        ],
    )

    if value:
        return clean_station_name(
            value
        )[:180]

    # --------------------------------------------------------
    # Look at headings.
    # --------------------------------------------------------

    for selector in [
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        ".modal-title",
        ".offcanvas-title",
        ".station-title",
    ]:

        loc = interface.locator(selector)

        try:

            for i in range(loc.count()):

                value = loc.nth(i).inner_text().strip()

                if not value:
                    continue

                if re.search(
                    r"station\s*index|period|rainfall|daily|view",
                    value,
                    re.I,
                ):
                    continue

                return clean_station_name(
                    value
                )[:180]

        except Exception:
            pass

    # --------------------------------------------------------
    # Look for station-like line.
    # --------------------------------------------------------

    lines = [
        " ".join(x.split()).strip()
        for x in text.splitlines()
        if x.strip()
    ]

    for line in lines:

        if len(line) < 2:
            continue

        if re.search(
            r"station\s*index|period|daily|rainfall|"
            r"date|total|view|hourly|point value",
            line,
            re.I,
        ):
            continue

        if re.fullmatch(
            r"[\d.,\-\s]+",
            line,
        ):
            continue

        # Don't accept generic form labels.
        if line.lower() in {
            "station",
            "name",
            "station name",
            "station index",
            "period",
        }:
            continue

        return clean_station_name(
            line
        )[:180]

    raise RuntimeError(
        "Station Name was not found in the "
        "station interface."
    )


# ============================================================
# PERIOD = DAILY
# ============================================================

def set_period_daily(interface):
    """
    Set Period = Daily.

    IMPORTANT:
    We do NOT modify the Date field.
    """

    selects = interface.locator(
        "select"
    )

    for i in range(
        selects.count()
    ):

        select = selects.nth(i)

        try:

            options = select.locator(
                "option"
            )

            option_texts = (
                options.all_inner_texts()
            )

            for option_text in option_texts:

                if option_text.strip().lower() == "daily":

                    # Try label first.
                    try:
                        select.select_option(
                            label="Daily"
                        )
                        return True
                    except Exception:
                        pass

                    # Then inspect option value.
                    for j in range(
                        options.count()
                    ):

                        option = options.nth(j)

                        text = option.inner_text().strip()

                        if text.lower() == "daily":

                            value = option.get_attribute(
                                "value"
                            )

                            if value is not None:

                                select.select_option(
                                    value=value
                                )

                                return True

        except Exception:
            pass

    # --------------------------------------------------------
    # Search for a select whose HTML mentions Daily.
    # --------------------------------------------------------

    try:

        daily_option = interface.locator(
            "select option"
        ).filter(
            has_text=re.compile(
                r"^\s*Daily\s*$",
                re.I,
            )
        ).first

        if daily_option.count():

            parent = daily_option.locator(
                ".."
            )

            parent.select_option(
                label="Daily"
            )

            return True

    except Exception:
        pass

    # --------------------------------------------------------
    # Custom dropdown.
    # --------------------------------------------------------

    custom_candidates = [
        interface.get_by_text(
            "Daily",
            exact=True,
        ).first,

        interface.locator(
            "[data-value='3']"
        ).first,

        interface.locator(
            "[value='3']"
        ).first,
    ]

    for candidate in custom_candidates:

        try:

            if candidate.is_visible():

                candidate.click(
                    force=True
                )

                page = interface.page

                page.wait_for_timeout(300)

                return True

        except Exception:
            pass

    return False


# ============================================================
# VIEW BUTTON
# ============================================================

def click_view(interface):
    candidates = [

        interface.get_by_role(
            "button",
            name="VIEW",
            exact=True,
        ).first,

        interface.locator(
            "button:has-text('VIEW')"
        ).first,

        interface.locator(
            "input[value='VIEW']"
        ).first,

        interface.locator(
            "input[value='View']"
        ).first,

        interface.locator(
            "a:has-text('VIEW')"
        ).first,

        interface.get_by_text(
            "VIEW",
            exact=True,
        ).first,
    ]

    for locator in candidates:

        try:

            if locator.is_visible():

                locator.click(
                    force=True,
                    timeout=5000,
                )

                return True

        except Exception:
            pass

    return False


# ============================================================
# TABLE
# ============================================================

def read_daily_table(interface):
    """
    Return:

        [
            (date, daily, total),
            ...
        ]

    from the station interface.
    """

    tables = interface.locator(
        "table"
    )

    best_rows = []

    for i in range(
        tables.count()
    ):

        table = tables.nth(i)

        try:

            if not table.is_visible():
                continue

            text = table.inner_text()

            low = text.lower()

            if "date" not in low:
                continue

            if "daily" not in low:
                continue

            rows = table.locator(
                "tr"
            )

            current = []

            for r in range(
                rows.count()
            ):

                cells = rows.nth(r).locator(
                    "th, td"
                )

                values = [
                    " ".join(x.split()).strip()
                    for x in cells.all_inner_texts()
                ]

                if len(values) < 2:
                    continue

                date_value = parse_date(
                    values[0]
                )

                if date_value is None:
                    continue

                daily = parse_number(
                    values[1]
                )

                if daily is None:
                    continue

                total = None

                if len(values) >= 3:
                    total = parse_number(
                        values[2]
                    )

                current.append(
                    (
                        date_value,
                        daily,
                        total,
                    )
                )

            if len(current) > len(best_rows):
                best_rows = current

        except Exception:
            pass

    return best_rows


# ============================================================
# WAIT FOR TABLE
# ============================================================

def wait_for_daily_table(interface):
    """
    Wait for DHM's VIEW/AJAX response.

    Then return ALL daily rows.
    """

    for _ in range(60):

        rows = read_daily_table(
            interface
        )

        if rows:
            return rows

        interface.page.wait_for_timeout(
            500
        )

    raise RuntimeError(
        "VIEW was clicked but the Daily rainfall table "
        "did not appear."
    )


# ============================================================
# FILTER TODAY
# ============================================================

def remove_today(rows):
    """
    Keep every returned row except today's AD/Gregorian date.

    IMPORTANT:
    We do NOT keep only yesterday.

    If DHM returns:

        Sep 26
        Sep 25
        Sep 24
        Sep 23
        ...

    we save:

        Sep 25
        Sep 24
        Sep 23
        ...
    """

    today = today_ad()

    result = []

    for row in rows:

        date_value = row[0]

        if date_value == today:
            continue

        result.append(row)

    return result


# ============================================================
# EXCEL
# ============================================================

def safe_filename(station_index, station_name):
    station_index = str(
        station_index
    ).strip()

    clean_name = re.sub(
        r"[^\w\-.]+",
        "_",
        station_name,
    ).strip("_")

    if not clean_name:
        clean_name = "Unknown_Station"

    return (
        f"{station_index}_{clean_name}.xlsx"
    )


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

        if not row:
            continue

        value = row[0]

        if value is None:
            continue

        if isinstance(
            value,
            dt.datetime,
        ):
            dates.add(
                value.date()
            )

        elif isinstance(
            value,
            dt.date,
        ):
            dates.add(value)

        else:

            parsed = parse_date(
                str(value)
            )

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

        wb = load_workbook(
            path
        )

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

    # Format date column.
    for cell in ws["A"][1:]:

        if isinstance(
            cell.value,
            dt.datetime,
        ):
            cell.number_format = "yyyy-mm-dd"

        elif isinstance(
            cell.value,
            dt.date,
        ):
            cell.number_format = "yyyy-mm-dd"

    ws.freeze_panes = "A2"

    ws.auto_filter.ref = ws.dimensions

    wb.save(path)

    return added


# ============================================================
# CLOSE STATION INTERFACE
# ============================================================

def close_station_interface(page):
    """
    Close whatever DHM station interface is currently open.

    We deliberately don't assume it is a Leaflet popup.
    """

    candidates = [

        page.get_by_role(
            "button",
            name=re.compile(
                r"close",
                re.I,
            ),
        ).last,

        page.locator(
            ".modal.show .btn-close"
        ).last,

        page.locator(
            ".modal.show [data-dismiss='modal']"
        ).last,

        page.locator(
            ".offcanvas.show .btn-close"
        ).last,

        page.locator(
            "[aria-label='Close']:visible"
        ).last,

        page.locator(
            "button.close:visible"
        ).last,
    ]

    for locator in candidates:

        try:

            if locator.is_visible():

                locator.click(
                    force=True,
                    timeout=2000,
                )

                page.wait_for_timeout(
                    300
                )

                return

        except Exception:
            pass

    # Escape is only a final fallback.
    try:
        page.keyboard.press(
            "Escape"
        )

        page.wait_for_timeout(
            300
        )

    except Exception:
        pass


# ============================================================
# CLICK ONE REAL MAP MARKER
# ============================================================

def click_marker_position(
    page,
    marker,
):
    """
    Perform an actual browser mouse click at the captured
    Leaflet marker position.

    This is not dispatch_event('click').

    It is a real mouse move + mouse down/up.
    """

    x = marker["x"]
    y = marker["y"]

    print(
        f"    marker screen position: "
        f"x={x:.1f}, y={y:.1f}"
    )

    page.mouse.move(
        x,
        y,
        steps=10,
    )

    page.wait_for_timeout(
        100
    )

    page.mouse.down()

    page.wait_for_timeout(
        80
    )

    page.mouse.up()


# ============================================================
# PROCESS ONE STATION
# ============================================================

def process_station(
    page,
    marker,
):
    sequence = marker["sequence"]

    print()
    print(
        "----------------------------------------"
    )

    print(
        f"[{sequence}] Clicking station marker..."
    )

    click_marker_position(
        page,
        marker,
    )

    # --------------------------------------------------------
    # Wait for station-specific interface.
    # --------------------------------------------------------

    interface = find_station_interface(
        page
    )

    if interface is None:

        save_debug(
            page,
            f"ERROR_station_{sequence:04d}_no_interface",
        )

        raise RuntimeError(
            "Marker was clicked, but the station-specific "
            "interface outside the map was not found."
        )

    # --------------------------------------------------------
    # Station Name.
    # --------------------------------------------------------

    station_name = extract_station_name(
        interface
    )

    # --------------------------------------------------------
    # ACTUAL Station Index.
    #
    # IMPORTANT:
    # This is NOT `sequence`.
    # --------------------------------------------------------

    station_index = extract_station_index(
        interface
    )

    print(
        f"    Station Name: {station_name}"
    )

    print(
        f"    Station Index: {station_index}"
    )

    # --------------------------------------------------------
    # Set Period = Daily.
    #
    # Date is NEVER changed.
    # --------------------------------------------------------

    print(
        "    Setting Period = Daily..."
    )

    if not set_period_daily(
        interface
    ):
        raise RuntimeError(
            "Could not set Period to Daily."
        )

    # --------------------------------------------------------
    # VIEW
    # --------------------------------------------------------

    print(
        "    Clicking VIEW..."
    )

    if not click_view(
        interface
    ):
        raise RuntimeError(
            "VIEW button was not found."
        )

    # --------------------------------------------------------
    # Daily table.
    # --------------------------------------------------------

    print(
        "    Waiting for Daily table..."
    )

    rows = wait_for_daily_table(
        interface
    )

    print(
        f"    DHM returned {len(rows)} row(s)."
    )

    # --------------------------------------------------------
    # Remove ONLY today's row.
    # --------------------------------------------------------

    filtered = remove_today(
        rows
    )

    print(
        f"    Rows after removing today "
        f"({today_iso()}): {len(filtered)}"
    )

    if not filtered:
        raise RuntimeError(
            "No rows remained after removing today's row."
        )

    # --------------------------------------------------------
    # Excel filename uses ACTUAL Station Index.
    # --------------------------------------------------------

    filename = safe_filename(
        station_index,
        station_name,
    )

    path = (
        OUTPUT_DIR
        / filename
    )

    added = append_rows(
        path,
        filtered,
    )

    print(
        f"    Added {added} row(s)"
    )

    print(
        f"    Excel: {path}"
    )

    if DEBUG:

        save_debug(
            page,
            f"station_{sequence:04d}_{station_index}_complete",
        )

    return {
        "sequence": sequence,
        "station_index": station_index,
        "station_name": station_name,
        "rows": filtered,
        "path": path,
    }


# ============================================================
# MAIN
# ============================================================

def main():

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print()
    print(
        "=========================================="
    )
    print(
        "DHM RAINFALL DAILY EXTRACTOR"
    )
    print(
        "=========================================="
    )

    print(
        f"Today AD: {today_iso()}"
    )

    print(
        "Date field: NOT MODIFIED"
    )

    print(
        "Period: Daily"
    )

    print(
        "Today's row: EXCLUDED"
    )

    print(
        f"DEBUG={DEBUG}"
    )

    print(
        f"HEADLESS={HEADLESS}"
    )

    print(
        f"MAX_STATIONS={MAX_STATIONS}"
    )

    print(
        "=========================================="
    )

    with sync_playwright() as p:

        browser = p.chromium.launch(

            headless=HEADLESS,

            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ],
        )

        context = browser.new_context(

            viewport={
                "width": 1440,
                "height": 1000,
            },

            locale="en-US",

            # Keep browser behavior stable.
            device_scale_factor=1,
        )

        page = context.new_page()

        try:

            # =================================================
            # 1. OPEN 24 HOUR MAP
            # =================================================

            click_24_hour(
                page
            )

            # =================================================
            # 2. CAPTURE ACTUAL RENDERED MARKERS
            # =================================================

            markers = capture_markers(
                page
            )

            total = len(markers)

            print()
            print(
                f"Total rendered station markers: {total}"
            )

            # =================================================
            # 3. DEBUG LIMIT
            # =================================================

            if MAX_STATIONS > 0:

                markers = markers[
                    :MAX_STATIONS
                ]

                print(
                    f"DEBUG LIMIT ACTIVE: "
                    f"processing first {len(markers)} station(s)."
                )

            else:

                print(
                    f"Processing all {len(markers)} station(s)."
                )

            # =================================================
            # 4. PROCESS
            # =================================================

            successful = 0
            failed = 0

            results = []

            for marker in markers:

                sequence = marker[
                    "sequence"
                ]

                try:

                    result = process_station(
                        page,
                        marker,
                    )

                    results.append(
                        result
                    )

                    successful += 1

                except Exception as exc:

                    failed += 1

                    print(
                        f"    FAILED station "
                        f"{sequence}: {exc}"
                    )

                    save_debug(
                        page,
                        f"ERROR_station_{sequence:04d}",
                    )

                finally:

                    # Close station interface before
                    # moving to the next marker.
                    close_station_interface(
                        page
                    )

                    # Allow map to settle.
                    page.wait_for_timeout(
                        500
                    )

            # =================================================
            # 5. SUMMARY
            # =================================================

            print()
            print(
                "=========================================="
            )
            print(
                "FINISHED"
            )
            print(
                "=========================================="
            )

            print(
                f"Rendered markers: {total}"
            )

            print(
                f"Processed: {len(markers)}"
            )

            print(
                f"Successful: {successful}"
            )

            print(
                f"Failed: {failed}"
            )

            print(
                f"Today excluded: {today_iso()}"
            )

            print(
                f"Output: {OUTPUT_DIR.resolve()}"
            )

            print(
                "=========================================="
            )

        finally:

            try:
                page.close()
            except Exception:
                pass

            try:
                context.close()
            except Exception:
                pass

            try:
                browser.close()
            except Exception:
                pass


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()