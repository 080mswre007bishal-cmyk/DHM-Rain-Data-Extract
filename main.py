import datetime as dt
import json
import os
import re
import time
from pathlib import Path

import cv2
import numpy as np
from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright

MAP_URL = "https://dhm.gov.np/hydrology/rainfall-watch-map"
OUTPUT_DIR = Path("output")
DEBUG_DIR = OUTPUT_DIR / "debug"
NETWORK_LOG = DEBUG_DIR / "network.log"

DEBUG = os.getenv("DEBUG", "0") == "1"
HEADLESS = os.getenv("HEADLESS", "1") != "0"
MAX_STATIONS = int(os.getenv("MAX_STATIONS", "0"))
DEBUG_MAX_STATIONS = 10

# The DHM screenshot shows small filled circles, roughly 18-26 px across.
# Hough settings below are deliberately conservative to avoid roads/rivers.
HOUGH_SETTINGS = [
    # param2, minDist, minRadius, maxRadius
    (22, 16, 7, 15),
    (20, 16, 7, 15),
    (24, 15, 7, 15),
    (18, 18, 7, 15),
]

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


def today_iso():
    return dt.date.today().strftime("%Y-%m-%d")


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
        m = re.search(pattern, text)
        if not m:
            continue
        token = m.group(0)
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
    m = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(m.group(0)) if m else None


def save_debug(page, tag, extra_text=None):
    if not DEBUG:
        return
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w.-]+", "_", tag)
    try:
        page.screenshot(path=str(DEBUG_DIR / f"{safe}.png"), full_page=True)
    except Exception:
        pass
    try:
        (DEBUG_DIR / f"{safe}.html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass
    try:
        (DEBUG_DIR / f"{safe}.txt").write_text(
            page.locator("body").inner_text(), encoding="utf-8"
        )
    except Exception:
        pass
    if extra_text is not None:
        try:
            (DEBUG_DIR / f"{safe}.json").write_text(
                json.dumps(extra_text, indent=2, default=str), encoding="utf-8"
            )
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Network/console logging. This is diagnostic only. We still click the map
# markers with a real mouse and extract rainfall through the DHM interface.
# ---------------------------------------------------------------------------
def attach_network_logger(page):
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    log = NETWORK_LOG.open("a", encoding="utf-8")

    def write(obj):
        try:
            log.write(json.dumps(obj, default=str) + "\n")
            log.flush()
        except Exception:
            pass

    page.on("console", lambda msg: write({"console": msg.type, "text": msg.text[:500]}))
    page.on("pageerror", lambda err: write({"pageerror": str(err)[:1000]}))

    def on_response(resp):
        try:
            req = resp.request
            if req.resource_type not in ("xhr", "fetch"):
                return
            if "dhm.gov.np" not in resp.url:
                return
            item = {
                "method": req.method,
                "url": resp.url,
                "post_data": req.post_data,
                "status": resp.status,
            }
            try:
                body = resp.json()
                if isinstance(body, dict):
                    item["json_keys"] = list(body.keys())[:20]
                    if "data" in body and isinstance(body["data"], list):
                        item["data_length"] = len(body["data"])
                elif isinstance(body, list):
                    item["json_type"] = f"list[{len(body)}]"
            except Exception:
                pass
            write(item)
        except Exception:
            pass

    page.on("response", on_response)
    return log


# ---------------------------------------------------------------------------
# Page/map setup
# ---------------------------------------------------------------------------
def click_24_hour(page):
    page.goto(MAP_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(3500)

    candidates = [
        page.get_by_text("24 Hour", exact=True).first,
        page.get_by_role("button", name="24 Hour", exact=True).first,
        page.locator("a:has-text('24 Hour')").first,
        page.locator("button:has-text('24 Hour')").first,
    ]

    for loc in candidates:
        try:
            loc.wait_for(state="visible", timeout=4000)
            loc.click(force=True)
            page.wait_for_timeout(5000)
            return
        except Exception:
            continue

    raise RuntimeError("Could not click the DHM '24 Hour' tab.")


def get_map_box(page):
    map_el = page.locator("#map").first
    map_el.wait_for(state="visible", timeout=15000)
    box = map_el.bounding_box()
    if not box or box["width"] < 100 or box["height"] < 100:
        raise RuntimeError(f"DHM map has invalid bounding box: {box}")
    return map_el, box


# ---------------------------------------------------------------------------
# Screenshot-based station marker detector.
#
# IMPORTANT: We do NOT monkey-patch Leaflet and we do NOT depend on
# .leaflet-marker-icon. DHM's rainfall circles are rendered visually on the
# map, and this detector finds those circles from the rendered image.
# ---------------------------------------------------------------------------
def marker_color_score(hsv, x, y, r):
    h, w = hsv.shape[:2]
    rr = max(3, int(r * 0.55))
    x0, x1 = max(0, x - rr), min(w, x + rr + 1)
    y0, y1 = max(0, y - rr), min(h, y + rr + 1)
    patch = hsv[y0:y1, x0:x1]
    if patch.size == 0:
        return 0.0, 0.0, 999.0

    # Use the middle portion of the circle. White rainfall text occupies only
    # a small fraction of the center, so median HSV is a good fill estimate.
    sat = float(np.median(patch[:, :, 1]))
    val = float(np.median(patch[:, :, 2]))
    gray_std = float(np.std(patch[:, :, 2]))

    # Colored markers: moderate/high saturation and non-white fill.
    colored = sat >= 22 and val < 245

    # Gray markers: low saturation, but a dark/medium fill rather than map
    # background/white labels.
    gray_marker = sat < 22 and 70 <= val <= 205 and gray_std < 55

    if colored or gray_marker:
        return 1.0, sat, val
    return 0.0, sat, val


def dedupe_points(points, min_distance=13):
    # Prefer larger/stronger circles first so duplicate Hough detections keep
    # one stable center.
    points = sorted(points, key=lambda p: (-p[2], p[1], p[0]))
    kept = []
    for p in points:
        x, y, r = p
        duplicate = False
        for q in kept:
            dx = x - q[0]
            dy = y - q[1]
            if (dx * dx + dy * dy) ** 0.5 < min_distance:
                duplicate = True
                break
        if not duplicate:
            kept.append(p)
    return sorted(kept, key=lambda p: (p[1], p[0]))


def detect_marker_points(map_png):
    arr = cv2.imdecode(np.frombuffer(map_png, dtype=np.uint8), cv2.IMREAD_COLOR)
    if arr is None:
        raise RuntimeError("Could not decode the map screenshot.")

    h, w = arr.shape[:2]
    gray = cv2.cvtColor(arr, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (5, 5), 1.1)
    hsv = cv2.cvtColor(arr, cv2.COLOR_BGR2HSV)

    all_points = []

    for param2, min_dist, min_radius, max_radius in HOUGH_SETTINGS:
        circles = cv2.HoughCircles(
            gray,
            cv2.HOUGH_GRADIENT,
            dp=1.2,
            minDist=min_dist,
            param1=80,
            param2=param2,
            minRadius=min_radius,
            maxRadius=max_radius,
        )
        if circles is None:
            continue

        for xf, yf, rf in np.round(circles[0]).astype(int):
            x, y, r = int(xf), int(yf), int(rf)

            # Keep only circles well inside the map. The top-left Leaflet
            # zoom controls are excluded explicitly.
            if x < 10 or y < 12 or x >= w - 8 or y >= h - 8:
                continue
            if y < 75 and x < 80:
                continue

            score, sat, val = marker_color_score(hsv, x, y, r)
            if score < 1:
                continue

            # DHM station circles are approximately 18-30 px in diameter.
            # This additional radius test removes small map-artifacts.
            if r < 7 or r > 15:
                continue

            all_points.append((x, y, r, sat, val))

    # Dedupe the same visual marker detected by multiple Hough settings.
    raw = [(p[0], p[1], p[2]) for p in all_points]
    points = dedupe_points(raw, min_distance=14)

    return points


def find_map_markers(page):
    map_el, box = get_map_box(page)

    # Wait for the rainfall layer. We intentionally do not wait for a specific
    # Leaflet DOM selector because the DHM circles are not reliably represented
    # by .leaflet-marker-icon.
    last_count = 0
    last_png = None
    for attempt in range(1, 21):
        try:
            png = map_el.screenshot(type="png")
            points = detect_marker_points(png)
            last_count = len(points)
            last_png = png
            print(f"  Marker image scan {attempt}/20: {last_count} circular candidates")

            # The live DHM map normally has many stations. Five or more is
            # enough to prove that the rainfall layer has rendered.
            if last_count >= 5:
                break
        except Exception as exc:
            print(f"  Marker image scan {attempt}/20 failed: {exc}")
        page.wait_for_timeout(1000)

    if not last_png or last_count == 0:
        raise RuntimeError(
            "No circular DHM rainfall markers were detected from the rendered map screenshot. "
            "Set DEBUG=1 and inspect output/debug/ for the map screenshot."
        )

    # Re-read the box after the final screenshot in case the page resized.
    _, box = get_map_box(page)
    arr = cv2.imdecode(np.frombuffer(last_png, dtype=np.uint8), cv2.IMREAD_COLOR)
    img_h, img_w = arr.shape[:2]
    scale_x = box["width"] / img_w
    scale_y = box["height"] / img_h

    viewport_points = []
    for x, y, r in points:
        vx = box["x"] + x * scale_x
        vy = box["y"] + y * scale_y
        viewport_points.append({
            "map_x": x,
            "map_y": y,
            "radius": r,
            "x": vx,
            "y": vy,
        })

    # Save the exact detected points when DEBUG is enabled.
    if DEBUG:
        try:
            annotated = arr.copy()
            for p in viewport_points:
                cv2.circle(
                    annotated,
                    (int(p["map_x"]), int(p["map_y"])),
                    int(p["radius"]),
                    (0, 0, 255),
                    2,
                )
                cv2.circle(
                    annotated,
                    (int(p["map_x"]), int(p["map_y"])),
                    2,
                    (0, 255, 0),
                    -1,
                )
            cv2.imwrite(str(DEBUG_DIR / "detected_markers.png"), annotated)
            (DEBUG_DIR / "detected_markers.json").write_text(
                json.dumps(viewport_points, indent=2), encoding="utf-8"
            )
        except Exception:
            pass

    print(f"Detected {len(viewport_points)} rainfall marker(s) from the rendered map.")
    return viewport_points


# ---------------------------------------------------------------------------
# Station interface discovery.
# ---------------------------------------------------------------------------
def is_inside_leaflet_map(locator):
    try:
        return bool(locator.evaluate("el => !!el.closest('.leaflet-container')"))
    except Exception:
        return False


def interface_candidates(page):
    """
    Find a station interface outside the Leaflet map.

    We locate the visible 'Station Index' text and climb ancestors until we
    find a container that also contains the Period form / VIEW control.
    This avoids depending on one particular modal/offcanvas class.
    """
    candidates = []

    try:
        labels = page.get_by_text(re.compile(r"^\s*Station\s*Index\s*:?\s*$", re.I))
        count = labels.count()
    except Exception:
        count = 0
        labels = None

    for i in range(count):
        label = labels.nth(i)
        try:
            if not label.is_visible() or is_inside_leaflet_map(label):
                continue

            current = label
            for level in range(1, 10):
                current = current.locator("xpath=..")
                try:
                    if not current.is_visible():
                        continue
                    if is_inside_leaflet_map(current):
                        continue

                    text = current.inner_text(timeout=500)
                    low = text.lower()
                    if "station index" not in low:
                        continue
                    if "view" not in low:
                        continue
                    if current.locator("select").count() == 0 and current.locator(
                        "button:has-text('VIEW'), input[value='VIEW'], a:has-text('VIEW')"
                    ).count() == 0:
                        continue

                    candidates.append(current)
                    break
                except Exception:
                    continue
        except Exception:
            continue

    # Additional common containers as fallback.
    selectors = [
        "[role='dialog']:visible",
        ".modal.show:visible",
        ".modal.in:visible",
        ".offcanvas.show:visible",
        "[class*='station']:visible",
        "[id*='station']:visible",
        "[class*='rainfall']:visible",
        "[id*='rainfall']:visible",
        "[class*='observation']:visible",
        "[id*='observation']:visible",
    ]
    for selector in selectors:
        try:
            loc = page.locator(selector)
            for i in range(loc.count()):
                el = loc.nth(i)
                if not el.is_visible() or is_inside_leaflet_map(el):
                    continue
                text = el.inner_text(timeout=500)
                low = text.lower()
                if "station index" not in low or "view" not in low:
                    continue
                if el.locator("select").count() == 0 and el.locator(
                    "button:has-text('VIEW'), input[value='VIEW'], a:has-text('VIEW')"
                ).count() == 0:
                    continue
                candidates.append(el)
        except Exception:
            continue

    # Deduplicate by outerHTML signature.
    result = []
    seen = set()
    for el in candidates:
        try:
            sig = el.evaluate(
                "el => el.tagName + '|' + (el.id || '') + '|' + (typeof el.className === 'string' ? el.className : '')"
            )
            if sig in seen:
                continue
            seen.add(sig)
            result.append(el)
        except Exception:
            pass
    return result


def wait_for_station_interface(page, timeout_seconds=20):
    end = time.time() + timeout_seconds
    while time.time() < end:
        candidates = interface_candidates(page)
        if candidates:
            # Prefer the smallest container that contains the form. This avoids
            # selecting the whole page/sidebar when a nested station panel exists.
            ranked = []
            for el in candidates:
                try:
                    txt_len = len(el.inner_text(timeout=500))
                    ranked.append((txt_len, el))
                except Exception:
                    pass
            if ranked:
                ranked.sort(key=lambda x: x[0])
                return ranked[0][1]
        time.sleep(0.25)
    return None


def normalize_text(text):
    return " ".join(str(text or "").split())


def extract_station_name(text):
    text = normalize_text(text)
    patterns = [
        r"\bStation\s*:\s*(.+?)(?=\s+District\s*:)",
        r"\bStation\s*:\s*(.+?)(?=\s+Station\s+Index\s*:)",
        r"\bStation\s*Name\s*[:\-]\s*(.+?)(?=\s+(?:District|Station\s+Index)\s*:|$)",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if m:
            value = m.group(1).strip(" |:-")
            if value and value.lower() not in {"name", "index"}:
                return value[:180]

    # Line-based fallback.
    lines = [x.strip() for x in str(text).splitlines() if x.strip()]
    for i, line in enumerate(lines):
        if re.fullmatch(r"Station(?:\s+Name)?\s*[:\-]?", line, re.I):
            if i + 1 < len(lines):
                value = lines[i + 1].strip()
                if value and not re.match(
                    r"^(District|Station\s+Index|Latitude|Longitude|Date|Period|VIEW)\b",
                    value,
                    re.I,
                ):
                    return value[:180]
    return None


def extract_station_index(text):
    text = normalize_text(text)
    m = re.search(
        r"\bStation\s*Index\s*[:\-]?\s*([A-Za-z0-9_-]*\d[A-Za-z0-9_-]*)",
        text,
        re.I,
    )
    return m.group(1).strip() if m else None


def station_details(page, panel):
    sources = []
    try:
        sources.append(panel.inner_text())
    except Exception:
        pass

    # The name and Station Index can be split into separate children. Include
    # a few ancestors so both values can be captured together.
    try:
        ancestors = panel.locator("xpath=ancestor::*")
        for i in range(min(ancestors.count(), 7)):
            el = ancestors.nth(i)
            try:
                if el.is_visible() and not is_inside_leaflet_map(el):
                    sources.append(el.inner_text(timeout=500))
            except Exception:
                pass
    except Exception:
        pass

    name = None
    index = None
    for text in sources:
        if not index:
            index = extract_station_index(text)
        if not name:
            name = extract_station_name(text)
        if name and index:
            break

    if not name or not index:
        # Search visible labels directly and use their nearby container.
        try:
            station_labels = page.get_by_text(re.compile(r"Station\s*(?:Name)?\s*:", re.I))
            for i in range(station_labels.count()):
                el = station_labels.nth(i)
                if not el.is_visible() or is_inside_leaflet_map(el):
                    continue
                parent = el.locator("xpath=..").first
                text = parent.inner_text(timeout=500)
                name = name or extract_station_name(text)
                index = index or extract_station_index(text)
                if name and index:
                    break
        except Exception:
            pass

    if not name:
        raise RuntimeError("Station name was not found in the station interface.")
    if not index:
        raise RuntimeError("Actual Station Index was not found in the station interface.")
    return name, index


# ---------------------------------------------------------------------------
# Daily period + rainfall table.
# ---------------------------------------------------------------------------
def set_period_daily(panel):
    selects = panel.locator("select")
    available = []

    for i in range(selects.count()):
        sel = selects.nth(i)
        try:
            options = [x.strip() for x in sel.locator("option").all_inner_texts()]
            available.append(options)
            if any(x.lower() == "daily" for x in options):
                sel.select_option(label="Daily")
                selected = sel.locator("option:checked").inner_text().strip()
                if selected.lower() != "daily":
                    raise RuntimeError(
                        f"Period selection failed; selected value is {selected!r}."
                    )
                print("    Period: Daily")
                return
        except Exception:
            continue

    raise RuntimeError(f"Daily option was not found. Available selects/options: {available}")


def click_view(panel):
    candidates = [
        panel.get_by_role("button", name="VIEW", exact=True).first,
        panel.locator("button:has-text('VIEW')").first,
        panel.locator("input[value='VIEW']").first,
        panel.locator("a:has-text('VIEW')").first,
        panel.get_by_text("VIEW", exact=True).first,
    ]
    for loc in candidates:
        try:
            if loc.count() and loc.is_visible():
                loc.click(force=True)
                print("    VIEW clicked")
                return
        except Exception:
            continue
    raise RuntimeError("VIEW button was not found in the station interface.")


def parse_table(table):
    rows = table.locator("tr")
    if rows.count() < 2:
        return []

    header = [normalize_text(x).lower() for x in rows.nth(0).locator("th,td").all_inner_texts()]
    if "date" not in header or "daily" not in header:
        return []

    date_i = header.index("date")
    daily_i = header.index("daily")
    total_i = header.index("total") if "total" in header else None

    output = []
    for r in range(1, rows.count()):
        cells = [normalize_text(x) for x in rows.nth(r).locator("th,td").all_inner_texts()]
        if len(cells) <= max(date_i, daily_i):
            continue

        date_value = parse_date(cells[date_i])
        daily = parse_number(cells[daily_i])
        total = None
        if total_i is not None and len(cells) > total_i:
            total = parse_number(cells[total_i])

        if date_value and daily is not None:
            output.append((date_value, daily, total))

    return output


def read_daily_table(page):
    tables = page.locator("table:visible")
    for i in range(tables.count()):
        try:
            rows = parse_table(tables.nth(i))
            if rows:
                return rows
        except Exception:
            continue
    return []


def extract_rainfall(page, panel):
    set_period_daily(panel)
    click_view(panel)

    # DHM's response can take a while. Poll for the actual Daily table.
    deadline = time.time() + 90
    last_count = 0
    while time.time() < deadline:
        rows = read_daily_table(page)
        if rows:
            last_count = len(rows)
            kept = [row for row in rows if row[0] != today_iso()]
            print(f"    Daily table rows found: {last_count}")
            print(f"    Rows kept after removing today ({today_iso()}): {len(kept)}")
            return kept
        time.sleep(0.5)

    raise RuntimeError(
        f"No Daily rainfall table appeared within 90 seconds. Last detected row count: {last_count}"
    )


# ---------------------------------------------------------------------------
# Closing/cleanup.
# ---------------------------------------------------------------------------
def close_station_interface(page, panel=None):
    scopes = [panel] if panel is not None else []
    scopes.append(page)

    selectors = [
        ".btn-close:visible",
        "button[title='Close']:visible",
        "[aria-label='Close']:visible",
        "button:has-text('Close'):visible",
        ".leaflet-popup-close-button:visible",
    ]

    for scope in scopes:
        if scope is None:
            continue
        for selector in selectors:
            try:
                loc = scope.locator(selector).first
                if loc.count() and loc.is_visible():
                    loc.click(force=True)
                    page.wait_for_timeout(500)
                    return
            except Exception:
                continue

    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(500)
    except Exception:
        pass


def station_interface_still_visible(page):
    return wait_for_station_interface(page, timeout_seconds=0.3) is not None


def ensure_interface_closed(page):
    for _ in range(4):
        panel = wait_for_station_interface(page, timeout_seconds=0.3)
        if panel is None:
            return True
        close_station_interface(page, panel)
        page.wait_for_timeout(400)
    return not station_interface_still_visible(page)


# ---------------------------------------------------------------------------
# Excel output.
# ---------------------------------------------------------------------------
def safe_filename(station_index, station_name):
    index = re.sub(r"[^A-Za-z0-9_-]+", "_", str(station_index)).strip("_") or "UNKNOWN"
    name = re.sub(r"[^\w\-]+", "_", str(station_name)).strip("_") or "Unknown_Station"
    return f"{index}_{name}.xlsx"


def existing_dates(path):
    if not path.exists():
        return set()

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    result = set()

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        value = row[0]
        if isinstance(value, (dt.datetime, dt.date)):
            result.add(value.strftime("%Y-%m-%d"))
        else:
            parsed = parse_date(str(value))
            if parsed:
                result.add(parsed)

    wb.close()
    return result


def append_rows(path, rows):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if path.exists():
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Daily Rainfall"
        ws.append(["Date", "Daily", "Total"])

    existing = existing_dates(path)
    added = 0

    for date_value, daily, total in sorted(rows):
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
        if isinstance(cell.value, (dt.date, dt.datetime)):
            cell.number_format = "yyyy-mm-dd"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    wb.save(path)
    return added


# ---------------------------------------------------------------------------
# One real mouse click -> station interface -> Daily -> VIEW -> Excel.
# ---------------------------------------------------------------------------
def click_marker_with_mouse(page, point):
    x = float(point["x"])
    y = float(point["y"])

    print(f"    Mouse click at ({x:.1f}, {y:.1f})")
    page.mouse.move(x, y, steps=12)
    page.wait_for_timeout(100)
    page.mouse.down()
    page.wait_for_timeout(80)
    page.mouse.up()


def process_marker(page, point, sequence):
    if not ensure_interface_closed(page):
        raise RuntimeError("Previous station interface could not be closed.")

    click_marker_with_mouse(page, point)

    panel = wait_for_station_interface(page, timeout_seconds=20)
    if panel is None:
        raise RuntimeError(
            "Mouse clicked a rainfall circle, but no station interface appeared outside the map."
        )

    station_name, station_index = station_details(page, panel)
    print(f"    Station: {station_name}")
    print(f"    Station Index: {station_index}")

    if DEBUG:
        save_debug(
            page,
            f"station_{station_index}_before_view",
            {
                "sequence": sequence,
                "station_name": station_name,
                "station_index": station_index,
                "marker": point,
            },
        )

    rows = extract_rainfall(page, panel)

    if DEBUG:
        save_debug(
            page,
            f"station_{station_index}_after_view",
            {"rows": rows},
        )

    return station_name, station_index, rows, panel


# ---------------------------------------------------------------------------
# Main.
# ---------------------------------------------------------------------------
def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)

    print("===============================================")
    print("DHM RAINFALL WATCH MAP - DAILY STATION SCRAPER")
    print("===============================================")
    print(f"Today (AD): {today_iso()}")
    print("Date field is NOT modified.")
    print("Period is set to Daily, then VIEW is clicked.")
    print("All returned Daily rows are saved EXCEPT today's row.")
    print("Markers are found from the rendered circular rainfall symbols.")
    print(f"DEBUG={DEBUG} HEADLESS={HEADLESS} MAX_STATIONS={MAX_STATIONS}")

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
            device_scale_factor=1,
        )

        page = context.new_page()
        netlog = attach_network_logger(page)

        successful = 0
        failed = 0
        processed_indices = set()

        try:
            print("\nOpening DHM Rainfall Watch Map...")
            click_24_hour(page)
            print("24 Hour tab selected.")

            # Give the rainfall layer a moment to finish painting.
            page.wait_for_timeout(2000)

            markers = find_map_markers(page)

            # Debug mode means exactly the first 10 detected circles.
            limit = len(markers)
            if DEBUG:
                limit = min(limit, DEBUG_MAX_STATIONS)
            if MAX_STATIONS:
                limit = min(limit, MAX_STATIONS)

            markers = markers[:limit]
            print(f"\nWill process {len(markers)} detected marker(s).")

            if not markers:
                raise RuntimeError("No rainfall markers available after detection.")

            for sequence, point in enumerate(markers, start=1):
                print("\n" + "-" * 65)
                print(f"[{sequence}/{len(markers)}] Processing detected rainfall circle")
                print("-" * 65)

                panel = None
                try:
                    station_name, station_index, rows, panel = process_marker(
                        page, point, sequence
                    )

                    if station_index in processed_indices:
                        print(f"    Station Index {station_index} was already processed; skipping duplicate.")
                        successful += 1
                        continue

                    processed_indices.add(station_index)
                    filename = safe_filename(station_index, station_name)
                    path = OUTPUT_DIR / filename
                    added = append_rows(path, rows)

                    print(f"    Rows kept: {len(rows)}")
                    print(f"    New rows written: {added}")
                    print(f"    Excel: {path}")
                    successful += 1

                except Exception as exc:
                    failed += 1
                    print(f"    FAILED: {exc}")
                    save_debug(
                        page,
                        f"ERROR_marker_{sequence:04d}",
                        {"marker": point, "error": str(exc)},
                    )

                finally:
                    try:
                        close_station_interface(page, panel)
                    except Exception:
                        pass
                    try:
                        ensure_interface_closed(page)
                    except Exception:
                        pass
                    page.wait_for_timeout(500)

            print("\n===============================================")
            print("FINISHED")
            print("===============================================")
            print(f"Successful stations: {successful}")
            print(f"Failed marker attempts: {failed}")
            print(f"Unique Station Index values saved: {len(processed_indices)}")
            print(f"Output directory: {OUTPUT_DIR.resolve()}")
            print(f"Debug directory: {DEBUG_DIR.resolve()}")
            print(f"Network log: {NETWORK_LOG.resolve()}")

        finally:
            try:
                netlog.close()
            except Exception:
                pass
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


if __name__ == "__main__":
    main()
