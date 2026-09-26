import datetime as dt
import json
import os
import re
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

MAP_URL = "https://dhm.gov.np/hydrology/rainfall-watch-map"
OUTPUT_DIR = Path("output")
DEBUG_DIR = OUTPUT_DIR / "debug"
NETWORK_LOG = DEBUG_DIR / "network.log"

DEBUG = os.getenv("DEBUG", "0") == "1"
HEADLESS = os.getenv("HEADLESS", "1") != "0"
MAX_STATIONS = int(os.getenv("MAX_STATIONS", "0"))
DEBUG_MAX_STATIONS = 10

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
        if m:
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


def save_debug(page, tag, panel=None):
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
        (DEBUG_DIR / f"{safe}.txt").write_text(page.locator("body").inner_text(), encoding="utf-8")
    except Exception:
        pass
    if panel is not None:
        try:
            (DEBUG_DIR / f"{safe}_panel.html").write_text(
                panel.evaluate("el => el.outerHTML"), encoding="utf-8"
            )
            (DEBUG_DIR / f"{safe}_panel.txt").write_text(panel.inner_text(), encoding="utf-8")
        except Exception:
            pass


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
            if resp.request.resource_type in ("xhr", "fetch") and "dhm.gov.np" in resp.url:
                item = {
                    "method": resp.request.method,
                    "url": resp.url,
                    "post_data": resp.request.post_data,
                    "status": resp.status,
                }
                try:
                    body = resp.json()
                    item["json_keys"] = list(body)[:20] if isinstance(body, dict) else f"list[{len(body)}]"
                except Exception:
                    pass
                write(item)
        except Exception:
            pass

    page.on("response", on_response)
    return log


def click_24_hour(page):
    page.goto(MAP_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(3000)

    selectors = [
        "#myTab a:has-text('24 Hour')",
        "a:has-text('24 Hour')",
        "button:has-text('24 Hour')",
        "[role='tab']:has-text('24 Hour')",
    ]
    for selector in selectors:
        loc = page.locator(selector).first
        try:
            loc.wait_for(state="visible", timeout=5000)
            loc.scroll_into_view_if_needed()
            loc.click(timeout=5000)
            page.wait_for_timeout(5000)
            return
        except Exception:
            continue
    raise RuntimeError("Could not click the DHM '24 Hour' tab.")


def _is_station_svg(el):
    """Return True for a small, filled Leaflet SVG shape likely to be a circle marker."""
    try:
        if not el.is_visible():
            return False
        box = el.bounding_box()
        if not box:
            return False

        w = float(box["width"])
        h = float(box["height"])
        # Station circle markers are small and approximately square.
        if w < 4 or h < 4 or w > 40 or h > 40:
            return False
        if min(w, h) / max(w, h) < 0.65:
            return False

        fill = (el.get_attribute("fill") or "").strip().lower()
        stroke = (el.get_attribute("stroke") or "").strip().lower()
        style = (el.get_attribute("style") or "").lower()
        cls = (el.get_attribute("class") or "").lower()
        d = el.get_attribute("d") or ""

        # Exclude obvious non-filled line geometry.
        if fill in {"none", "", "transparent"} and "fill:" not in style:
            return False

        # Leaflet's CircleMarker is normally a path with a closed/curved d.
        # Keep explicit circles, and paths with a non-trivial d.
        if "circle" in cls:
            return True
        if el.evaluate("el => el.tagName.toLowerCase()") == "circle":
            return True
        if d and len(d) >= 10:
            return True

        # A few renderers expose style rather than a fill attribute.
        return bool(fill or stroke or style)
    except Exception:
        return False


def station_marker_candidates(page):
    """
    Find the actual small Leaflet SVG station markers.

    We deliberately do NOT use locator.nth(i) as a long-lived marker handle.
    The page can redraw the map after every click. Instead, each pass obtains
    fresh DOM elements and we click their screen coordinates like a human.
    """
    selectors = [
        ".leaflet-container .leaflet-overlay-pane svg path.leaflet-interactive",
        ".leaflet-container .leaflet-overlay-pane svg circle.leaflet-interactive",
        "#map .leaflet-overlay-pane svg path.leaflet-interactive",
        "#map .leaflet-overlay-pane svg circle.leaflet-interactive",
        ".leaflet-container .leaflet-marker-icon",
    ]

    seen = set()
    candidates = []
    for selector in selectors:
        loc = page.locator(selector)
        try:
            count = loc.count()
        except Exception:
            continue
        for i in range(count):
            el = loc.nth(i)
            try:
                if not el.is_visible():
                    continue
                box = el.bounding_box()
                if not box:
                    continue
                if "leaflet-overlay-pane" in selector and not _is_station_svg(el):
                    continue

                cx = round(box["x"] + box["width"] / 2, 1)
                cy = round(box["y"] + box["height"] / 2, 1)
                key = (round(cx), round(cy))
                if key in seen:
                    continue
                seen.add(key)
                candidates.append({
                    "selector": selector,
                    "index": i,
                    "x": cx,
                    "y": cy,
                    "w": round(box["width"], 1),
                    "h": round(box["height"], 1),
                    "class": el.get_attribute("class") or "",
                })
            except Exception:
                continue

    candidates.sort(key=lambda x: (x["y"], x["x"]))
    return candidates


def marker_diagnostic(page):
    data = {}
    selectors = [
        ".leaflet-container .leaflet-marker-icon",
        ".leaflet-container .leaflet-overlay-pane svg path.leaflet-interactive",
        ".leaflet-container .leaflet-overlay-pane svg circle.leaflet-interactive",
        ".leaflet-container .leaflet-overlay-pane svg path",
        ".leaflet-container .leaflet-overlay-pane svg circle",
    ]
    for selector in selectors:
        try:
            data[selector] = page.locator(selector).count()
        except Exception:
            data[selector] = -1
    data["station_candidates"] = len(station_marker_candidates(page))
    return data


def _normal_text(text):
    return " ".join(str(text or "").split())


def _extract_station_name(text):
    text = _normal_text(text)
    patterns = [
        r"\bStation\s*:\s*(.+?)(?=\s*\|?\s*District\s*:)",
        r"\bStation\s*:\s*(.+?)(?=\s+Station\s+Index\s*:)",
        r"\bStation\s*Name\s*[:\-]\s*(.+?)(?=\s+(?:District|Station\s+Index)\s*:|$)",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if m:
            value = m.group(1).strip(" |:-")
            if value:
                return value[:180]

    lines = [x.strip() for x in str(text).splitlines() if x.strip()]
    for i, line in enumerate(lines):
        if re.fullmatch(r"Station(?:\s+Name)?\s*[:\-]?", line, re.I) and i + 1 < len(lines):
            value = lines[i + 1].strip()
            if value and not re.match(
                r"^(District|Station\s+Index|Latitude|Longitude|Date|Period|VIEW)\b",
                value,
                re.I,
            ):
                return value[:180]
    return None


def _extract_station_index(text):
    text = _normal_text(text)
    m = re.search(
        r"\bStation\s*Index\s*[:\-]?\s*([A-Za-z0-9_-]*\d[A-Za-z0-9_-]*)",
        text,
        re.I,
    )
    return m.group(1).strip() if m else None


def candidate_station_interfaces(page):
    """Find the station interface OUTSIDE the Leaflet map."""
    selectors = [
        "[role='dialog']:visible",
        ".modal.show:visible",
        ".modal.in:visible",
        ".offcanvas.show:visible",
        "div[class*='station']:visible",
        "section[class*='station']:visible",
        "div[id*='station']:visible",
        "section[id*='station']:visible",
        "div[class*='rainfall']:visible",
        "section[class*='rainfall']:visible",
        "div[class*='observation']:visible",
        "section[class*='observation']:visible",
    ]

    result = []
    seen = set()
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
                    text = el.inner_text(timeout=1000)
                    low = text.lower()
                    if "station index" not in low:
                        continue
                    has_view = el.locator(
                        "button:has-text('VIEW'), input[value='VIEW'], a:has-text('VIEW')"
                    ).count() > 0
                    has_select = el.locator("select").count() > 0
                    if not (has_view or has_select):
                        continue
                    sig = el.evaluate(
                        "el => el.tagName + '|' + (el.id || '') + '|' + (el.className || '')"
                    )
                    if sig in seen:
                        continue
                    seen.add(sig)
                    result.append(el)
                except Exception:
                    continue
        except Exception:
            continue
    return result


def visible_station_panel(page):
    panels = candidate_station_interfaces(page)
    if not panels:
        return None

    # Prefer the smallest visible OUTSIDE-map element containing both station
    # identity and the controls needed for Period/View.
    ranked = []
    for panel in panels:
        try:
            text = panel.inner_text()
            idx = _extract_station_index(text)
            name = _extract_station_name(text)
            if idx and name:
                ranked.append((len(text), panel))
        except Exception:
            pass
    if ranked:
        ranked.sort(key=lambda x: x[0])
        return ranked[0][1]

    return panels[0]


def station_details(page, panel):
    # Deliberately DO NOT use page-level text. Station identity must come from
    # the station interface associated with the marker that was clicked.
    sources = []
    try:
        sources.append(panel.inner_text())
    except Exception:
        pass
    try:
        ancestors = panel.locator("xpath=ancestor::*")
        for i in range(min(6, ancestors.count())):
            el = ancestors.nth(i)
            try:
                if el.is_visible():
                    text = el.inner_text()
                    if "station index" in text.lower():
                        sources.append(text)
            except Exception:
                pass
    except Exception:
        pass

    station_name = None
    station_index = None
    for text in sources:
        idx = _extract_station_index(text)
        name = _extract_station_name(text)
        if idx and name:
            station_name, station_index = name, idx
            break

    if not station_index:
        for text in sources:
            station_index = _extract_station_index(text)
            if station_index:
                break
    if not station_name:
        for text in sources:
            station_name = _extract_station_name(text)
            if station_name:
                break

    if not station_name:
        raise RuntimeError("Station name not found in the clicked station interface.")
    if not station_index:
        raise RuntimeError("Station Index not found in the clicked station interface.")
    return station_name, station_index


def set_period_daily(panel):
    selects = panel.locator("select")
    for i in range(selects.count()):
        sel = selects.nth(i)
        try:
            options = [x.strip() for x in sel.locator("option").all_inner_texts()]
            for option in options:
                if option.lower() == "daily":
                    sel.select_option(label=option)
                    selected = sel.locator("option:checked").inner_text().strip()
                    if selected.lower() != "daily":
                        raise RuntimeError(f"Period did not become Daily: {selected!r}")
                    return
        except Exception:
            continue

    raise RuntimeError("Daily option not found in the clicked station interface.")


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
            loc.wait_for(state="visible", timeout=2000)
            loc.scroll_into_view_if_needed()
            loc.click(timeout=5000)
            return
        except Exception:
            continue
    raise RuntimeError("VIEW button was not found in the clicked station interface.")


def _table_rows(table):
    rows = table.locator("tr")
    if rows.count() == 0:
        return []
    header = [x.strip().lower() for x in rows.nth(0).locator("th,td").all_inner_texts()]
    if "date" not in header or "daily" not in header:
        return []
    date_i = header.index("date")
    daily_i = header.index("daily")
    total_i = header.index("total") if "total" in header else None

    out = []
    for r in range(1, rows.count()):
        vals = [x.strip() for x in rows.nth(r).locator("th,td").all_inner_texts()]
        if len(vals) <= max(date_i, daily_i):
            continue
        d = parse_date(vals[date_i])
        daily = parse_number(vals[daily_i])
        total = parse_number(vals[total_i]) if total_i is not None and len(vals) > total_i else None
        if d and daily is not None:
            out.append((d, daily, total))
    return out


def read_table(panel, page):
    # Prefer tables inside the station interface. DHM may place the table in a
    # sibling container, so use the page only as a tightly validated fallback.
    scopes = [panel, page]
    for scope in scopes:
        try:
            tables = scope.locator("table:visible")
            for i in range(tables.count()):
                rows = _table_rows(tables.nth(i))
                if rows:
                    return rows
        except Exception:
            continue
    return []


def extract_rainfall(page, panel):
    set_period_daily(panel)
    click_view(panel)

    deadline = time.time() + 90
    while time.time() < deadline:
        rows = read_table(panel, page)
        if rows:
            return [row for row in rows if row[0] != today_iso()]
        time.sleep(0.5)
    raise RuntimeError("No Daily rainfall table appeared after clicking VIEW.")


def close_station_panel(page, panel=None):
    scopes = [panel] if panel is not None else []
    scopes.append(page)
    selectors = [
        ".btn-close:visible",
        "button[title='Close']:visible",
        "[aria-label='Close']:visible",
        "button:has-text('Close'):visible",
    ]
    for scope in scopes:
        for selector in selectors:
            try:
                loc = scope.locator(selector).first
                if loc.count() and loc.is_visible():
                    loc.click(force=True)
                    page.wait_for_timeout(400)
                    return
            except Exception:
                continue
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(400)
    except Exception:
        pass


def wait_panel_gone(page):
    for _ in range(20):
        if visible_station_panel(page) is None:
            return
        time.sleep(0.2)


def click_marker_like_human(page, marker):
    """Click the marker's screen center with Playwright mouse coordinates."""
    x = marker["x"]
    y = marker["y"]
    # A real pointer click rather than dispatch_event/DOM click.
    page.mouse.move(x, y, steps=4)
    page.wait_for_timeout(100)
    page.mouse.down()
    page.wait_for_timeout(80)
    page.mouse.up()


def process_marker(page, marker, sequence):
    close_station_panel(page, visible_station_panel(page))
    wait_panel_gone(page)

    # The marker dictionary contains coordinates captured immediately before
    # the click, so this behaves like a person clicking the visible circle.
    click_marker_like_human(page, marker)

    panel = None
    for _ in range(40):
        panel = visible_station_panel(page)
        if panel is not None:
            break
        time.sleep(0.25)

    if panel is None:
        raise RuntimeError(
            f"Human click on marker {sequence} produced no outside-map station interface."
        )

    station_name, station_index = station_details(page, panel)
    print(f"    Station: {station_name}")
    print(f"    Station Index: {station_index}")

    if DEBUG:
        save_debug(page, f"station_{station_index}_opened", panel)

    rows = extract_rainfall(page, panel)
    return station_name, station_index, rows, panel


def safe_filename(station_index, station_name):
    clean = re.sub(r"[^\w\-]+", "_", station_name).strip("_") or "Unknown_Station"
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
        if isinstance(value, (dt.datetime, dt.date)):
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
        ws.append(["Date", "Daily Rainfall (mm)", "Running Total on page (mm)"])

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
                cell.number_format = "yyyy-mm-dd"
            except ValueError:
                pass

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)
    return added


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print("======================================")
    print("DHM Daily Rainfall Extractor - Human Click")
    print("======================================")
    print(f"Today (AD): {today_iso()}")
    print("Date field is NOT modified.")
    print("Period is set to Daily, then VIEW is clicked.")
    print("All returned rows are saved EXCEPT today's row.")
    print(f"DEBUG={DEBUG} HEADLESS={HEADLESS}")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        )
        context = browser.new_context(
            viewport={"width": 1440, "height": 1000},
            locale="en-US",
        )
        page = context.new_page()
        netlog = attach_network_logger(page)

        try:
            click_24_hour(page)

            # Wait for the actual station circles to be rendered.
            markers = []
            for _ in range(60):
                markers = station_marker_candidates(page)
                if markers:
                    break
                page.wait_for_timeout(500)

            if not markers:
                save_debug(page, "NO_MARKERS")
                print("Leaflet diagnostic:", marker_diagnostic(page))
                raise RuntimeError(
                    "No visible circular station markers were found. "
                    "See output/debug/NO_MARKERS.* and network.log."
                )

            limit = len(markers)
            if DEBUG:
                limit = min(limit, DEBUG_MAX_STATIONS)
            if MAX_STATIONS:
                limit = min(limit, MAX_STATIONS)

            print(f"Found {len(markers)} visible station marker(s). Processing {limit}.")

            processed = set()
            clicked_positions = set()
            successful = 0
            failed = 0

            # We deliberately reacquire the visible marker list before EVERY
            # click. The map can redraw after a station is opened/closed.
            # We select an unprocessed coordinate from the fresh list and then
            # click that coordinate immediately, rather than trusting nth(i).
            max_attempts = max(limit * 4, 20)
            attempts = 0

            while len(processed) < limit and attempts < max_attempts:
                attempts += 1
                current = station_marker_candidates(page)
                if not current:
                    page.wait_for_timeout(1000)
                    continue

                # Prefer the first marker whose coordinate has not already been
                # used in this run. Coordinates are only used to avoid repeated
                # clicks; station identity comes from the station interface.
                target = None
                for marker in current:
                    key = (round(marker["x"]), round(marker["y"]))
                    if key not in clicked_positions:
                        target = marker
                        break
                if target is None:
                    # Every currently visible coordinate has already been clicked.
                    # If the map has not moved, there is nothing more to do.
                    # If it redraws with shifted coordinates, the next iteration
                    # will see fresh coordinates.
                    print("    No unclicked marker coordinate in current map view; refreshing marker list.")
                    page.wait_for_timeout(500)
                    continue

                clicked_positions.add((round(target["x"]), round(target["y"])))

                # Do not use coordinate as station identity. The actual station
                # index returned by the interface is authoritative.
                print(
                    f"\n[attempt {attempts}] clicking visible marker at "
                    f"({target['x']:.0f}, {target['y']:.0f})"
                )

                try:
                    station_name, station_index, rows, panel = process_marker(
                        page, target, attempts
                    )

                    if station_index in processed:
                        print(f"    Already processed Station Index {station_index}; closing.")
                        close_station_panel(page, panel)
                        wait_panel_gone(page)
                        continue

                    processed.add(station_index)
                    path = OUTPUT_DIR / safe_filename(station_index, station_name)
                    added = append_rows(path, rows)
                    print(f"    Rows kept (excluding today): {len(rows)}")
                    print(f"    New rows added: {added}")
                    print(f"    Excel: {path}")
                    successful += 1

                    close_station_panel(page, panel)
                    wait_panel_gone(page)
                    page.wait_for_timeout(500)

                except Exception as exc:
                    failed += 1
                    print(f"    FAILED: {exc}")
                    save_debug(page, f"ERROR_attempt_{attempts:04d}")
                    try:
                        close_station_panel(page, visible_station_panel(page))
                        wait_panel_gone(page)
                    except Exception:
                        pass
                    page.wait_for_timeout(700)

            print("\n======================================")
            print("Finished.")
            print(f"Successful stations: {successful}")
            print(f"Failed attempts: {failed}")
            print(f"Unique stations saved: {len(processed)}")
            print(f"Output folder: {OUTPUT_DIR.resolve()}")
            print(f"Network log: {NETWORK_LOG}")
            print("======================================")

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
