import datetime as dt
import json
import os
import re
import time
from pathlib import Path

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

DATE_FORMATS = [
    "%a, %b %d, %Y, %I:%M %p", "%a, %b %d, %Y",
    "%Y-%m-%d", "%d %b, %Y", "%d %B, %Y",
    "%d %b %Y", "%d %B %Y", "%d/%m/%Y",
    "%m/%d/%Y", "%d-%m-%Y",
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
    for suffix, fn in (
        ("png", lambda: page.screenshot(path=str(DEBUG_DIR / f"{safe}.png"), full_page=True)),
        ("html", lambda: (DEBUG_DIR / f"{safe}.html").write_text(page.content(), encoding="utf-8")),
        ("txt", lambda: (DEBUG_DIR / f"{safe}.txt").write_text(page.locator("body").inner_text(), encoding="utf-8")),
    ):
        try:
            fn()
        except Exception:
            pass
    if panel is not None:
        try:
            (DEBUG_DIR / f"{safe}_panel.html").write_text(panel.evaluate("el => el.outerHTML"), encoding="utf-8")
            (DEBUG_DIR / f"{safe}_panel.txt").write_text(panel.inner_text(), encoding="utf-8")
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Network logging: always on (cheap). This is how you discover the real API
# endpoints for the future requests-only version of this scraper.
# ---------------------------------------------------------------------------
def attach_network_logger(page):
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    log = NETWORK_LOG.open("a", encoding="utf-8")

    def on_response(resp):
        try:
            if resp.request.resource_type in ("xhr", "fetch") and "dhm.gov.np" in resp.url:
                line = {"method": resp.request.method, "url": resp.url,
                        "post_data": resp.request.post_data, "status": resp.status}
                try:
                    body = resp.json()
                    line["json_keys"] = list(body)[:10] if isinstance(body, dict) else f"list[{len(body)}]"
                except Exception:
                    pass
                log.write(json.dumps(line, default=str) + "\n")
                log.flush()
        except Exception:
            pass

    page.on("response", on_response)
    return log


def click_24_hour(page):
    page.goto(MAP_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(3000)
    candidates = [
        page.get_by_text("24 Hour", exact=True).first,
        page.get_by_role("button", name="24 Hour", exact=True).first,
        page.locator("a:has-text('24 Hour')").first,
        page.locator("button:has-text('24 Hour')").first,
    ]
    for loc in candidates:
        try:
            loc.wait_for(state="visible", timeout=5000)
            loc.click(force=True)
            page.wait_for_timeout(4000)
            return
        except Exception:
            pass
    raise RuntimeError("Could not click the DHM '24 Hour' tab.")


def get_marker_locator(page):
    """
    Return clickable DHM rainfall station layers.

    Handles both image/divIcon markers (.leaflet-marker-icon) and SVG
    interactive layers (circle/path). Large SVG shapes (basin/river
    boundaries) are filtered out by bounding box.
    """
    selectors = [
        "#map .leaflet-marker-icon",
        ".leaflet-container .leaflet-marker-icon",
        "#map .leaflet-overlay-pane svg path.leaflet-interactive",
        ".leaflet-container .leaflet-overlay-pane svg path.leaflet-interactive",
        "#map .leaflet-overlay-pane svg circle.leaflet-interactive",
        ".leaflet-container .leaflet-overlay-pane svg circle.leaflet-interactive",
    ]
    for selector in selectors:
        loc = page.locator(selector)
        try:
            count = loc.count()
            if count > 0:
                if "leaflet-overlay-pane" in selector:
                    usable = []
                    for i in range(count):
                        el = loc.nth(i)
                        try:
                            if not el.is_visible():
                                continue
                            box = el.bounding_box()
                            if not box:
                                continue
                            if box["width"] <= 45 and box["height"] <= 45:
                                usable.append(i)
                        except Exception:
                            pass
                    if usable:
                        return _indexed_locator(loc, usable)
                else:
                    return loc
        except Exception:
            pass
    return None


def _indexed_locator(base, indexes):
    """Union locator containing only the supplied zero-based indexes."""
    if not indexes:
        return None
    result = base.nth(indexes[0])
    for idx in indexes[1:]:
        result = result.or_(base.nth(idx))
    return result


def leaflet_dom_diagnostic(page):
    try:
        return {
            "marker_icons": page.locator(".leaflet-container .leaflet-marker-icon").count(),
            "interactive_paths": page.locator(".leaflet-container .leaflet-overlay-pane svg path.leaflet-interactive").count(),
            "interactive_circles": page.locator(".leaflet-container .leaflet-overlay-pane svg circle.leaflet-interactive").count(),
            "leaflet_paths": page.locator(".leaflet-container .leaflet-overlay-pane svg path").count(),
            "leaflet_circles": page.locator(".leaflet-container .leaflet-overlay-pane svg circle").count(),
        }
    except Exception:
        return {}


def _candidate_panels(root):
    selectors = [
        "[role='dialog']:visible",
        ".modal.show:visible", ".modal.in:visible",
        ".offcanvas.show:visible",
        ".leaflet-popup:visible",                       # CHANGED: accept in-map popups too
        "div[class*='station']:visible", "section[class*='station']:visible",
        "div[id*='station']:visible", "section[id*='station']:visible",
        "div[class*='rainfall']:visible", "section[class*='rainfall']:visible",
        "div[class*='observation']:visible", "section[class*='observation']:visible",
        "div[class*='popup']:visible", "section[class*='popup']:visible",
    ]
    seen = set()
    result = []
    for selector in selectors:
        try:
            loc = root.locator(selector)
            for i in range(loc.count()):
                el = loc.nth(i)
                try:
                    if not el.is_visible():
                        continue
                    txt = el.inner_text(timeout=1000)
                    low = txt.lower()
                    # CHANGED: broader acceptance - "station index" is the
                    # key marker; require an action or period hint too.
                    if "station index" in low and ("view" in low or "daily" in low or "period" in low):
                        sig = (el.evaluate("el => el.tagName + '|' + (el.id || '') + '|' + (el.className || '')"), len(txt))
                        if sig in seen:
                            continue
                        seen.add(sig)
                        inside_map = el.evaluate("el => !!el.closest('.leaflet-container')")
                        result.append((inside_map, el))
                except Exception:
                    pass
        except Exception:
            pass
    return result


def visible_station_panel(page):
    panels = _candidate_panels(page)
    if not panels:
        return None
    # CHANGED: prefer panels OUTSIDE the map (smallest text first), but
    # accept an in-map Leaflet popup as a fallback instead of returning None.
    outside = [(len(p.inner_text()), p) for inside, p in panels if not inside]
    inside = [(len(p.inner_text()), p) for is_in, p in panels if is_in]
    for group in (outside, inside):
        if group:
            group.sort(key=lambda x: x[0])
            return group[0][1]
    return None


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
            if value and not re.match(r"^(District|Station\s+Index|Latitude|Longitude|Date|Period|VIEW)\b", value, re.I):
                return value[:180]
    return None


def _extract_station_index(text):
    text = _normal_text(text)
    m = re.search(r"\bStation\s*Index\s*[:\-]?\s*([A-Za-z0-9_-]+)", text, re.I)
    return m.group(1).strip() if m else None


def station_details(page, panel, sequence):
    # CHANGED: only fall back to page-level text as a last resort, and never
    # let a page-level match override panel/ancestor matches. The sidebar
    # "Sort & Filters" area also contains the words "Station Name" which
    # previously could hijack the regexes.
    sources = []
    try:
        sources.append(("panel", panel.inner_text()))
    except Exception:
        pass
    try:
        ancestors = panel.locator("xpath=ancestor::*")
        for i in range(min(ancestors.count(), 8)):
            el = ancestors.nth(i)
            try:
                if el.is_visible():
                    sources.append((f"ancestor_{i}", el.inner_text()))
            except Exception:
                pass
    except Exception:
        pass

    station_name = None
    station_index = None
    for _, text in sources:
        idx = _extract_station_index(text)
        name = _extract_station_name(text)
        if idx and name:
            station_name, station_index = name, idx
            break
    if not station_index:
        for _, text in sources:
            station_index = _extract_station_index(text)
            if station_index:
                break
    if not station_name:
        for _, text in sources:
            station_name = _extract_station_name(text)
            if station_name:
                break

    # Last resort: page-level text. Require "Station Index" to be present in
    # the same text so the filter sidebar can't pollute the match.
    if not station_name or not station_index:
        try:
            body_text = page.locator("body").inner_text()
            if _extract_station_index(body_text):
                station_name = station_name or _extract_station_name(body_text)
                station_index = station_index or _extract_station_index(body_text)
        except Exception:
            pass

    if not station_name:
        raise RuntimeError(f"Station name not found after marker {sequence}.")
    if not station_index:
        raise RuntimeError(f"Station Index not found for marker {sequence}.")
    return station_name, station_index


def set_period_daily(panel):
    selects = panel.locator("select")
    for i in range(selects.count()):
        sel = selects.nth(i)
        try:
            opts = [o.strip() for o in sel.locator("option").all_inner_texts()]
            if any(o.lower() == "daily" for o in opts):
                sel.select_option(label="Daily")
                selected = sel.locator("option:checked").inner_text().strip()
                if selected.lower() != "daily":
                    raise RuntimeError(f"Period did not become Daily; selected={selected!r}")
                return
        except Exception:
            pass
    # CHANGED: no more blind value="3" fallback. Show the real options so
    # the failure message is actionable.
    all_opts = []
    try:
        for i in range(selects.count()):
            all_opts.append(panel.locator("select").nth(i).locator("option").all_inner_texts())
    except Exception:
        pass
    raise RuntimeError(f"Daily option not found. Available selects/options: {all_opts}")


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
            loc.click(force=True)
            return
        except Exception:
            pass
    raise RuntimeError("VIEW button was not found in the station interface.")


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


def read_table(page):
    # CHANGED: search the whole page, not just the panel. DHM may inject
    # the daily table into a container outside the popup element.
    tables = page.locator("table:visible")
    for i in range(tables.count()):
        try:
            out = _table_rows(tables.nth(i))
            if out:
                return out
        except Exception:
            pass
    return []


def extract_rainfall(page, panel):
    set_period_daily(panel)
    click_view(panel)
    # CHANGED: 90s budget - DHM's server is slow and 20s was too short.
    for _ in range(180):
        rows = read_table(page)
        if rows:
            return [row for row in rows if row[0] != today_iso()]
        time.sleep(0.5)
    raise RuntimeError("No Daily rainfall table appeared after clicking VIEW.")


def close_station_panel(page, panel=None):
    # CHANGED: scope the close button to the panel first. Clicking the
    # LAST page-level close button could close unrelated modals.
    scopes = []
    if panel is not None:
        scopes.append(panel)
    scopes.append(page)
    selectors = [".btn-close:visible", "button[title='Close']:visible",
                 "[aria-label='Close']:visible", "button:has-text('Close'):visible",
                 ".leaflet-popup-close-button:visible"]
    for scope in scopes:
        for selector in selectors:
            try:
                loc = scope.locator(selector).first
                if loc.count() > 0 and loc.is_visible():
                    loc.click(force=True)
                    page.wait_for_timeout(300)
                    return
            except Exception:
                pass
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)
    except Exception:
        pass


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


def process_marker(page, marker):
    close_station_panel(page)
    try:
        marker.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    try:
        marker.click(force=True, timeout=5000)
    except Exception:
        marker.dispatch_event("click")

    panel = None
    for _ in range(40):
        panel = visible_station_panel(page)
        if panel is not None:
            break
        time.sleep(0.25)
    if panel is None:
        raise RuntimeError("Marker clicked, but no station popup appeared (checked both outside-map containers and Leaflet popups).")

    station_name, station_index = station_details(page, panel, 0)
    print(f"    Station: {station_name}")
    print(f"    Station Index: {station_index}")
    if DEBUG:
        save_debug(page, f"station_{station_index}_panel", panel)

    rows = extract_rainfall(page, panel)
    if DEBUG:
        save_debug(page, f"station_{station_index}_after_view", panel)
    return station_name, station_index, rows


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print("======================================")
    print("DHM Daily Rainfall Extractor (v2)")
    print("======================================")
    print(f"Today (AD): {today_iso()}")
    print("All returned rows are saved EXCEPT today's row.")
    print(f"DEBUG={DEBUG} HEADLESS={HEADLESS}")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        )
        context = browser.new_context(viewport={"width": 1440, "height": 1000}, locale="en-US")
        page = context.new_page()
        netlog = attach_network_logger(page)
        try:
            click_24_hour(page)
            marker_locator = None
            for _ in range(40):
                marker_locator = get_marker_locator(page)
                if marker_locator and marker_locator.count() > 0:
                    break
                page.wait_for_timeout(500)
            if not marker_locator:
                save_debug(page, "NO_MARKERS")
                counts = leaflet_dom_diagnostic(page)
                print(f"Leaflet DOM diagnostic: {counts}")
                netlog.close()
                if counts.get("leaflet_paths", 0) > 0 or counts.get("leaflet_circles", 0) > 0:
                    raise RuntimeError(
                        "No clickable station layers found, but Leaflet shapes exist - the map "
                        "likely uses the CANVAS renderer (no DOM elements to click). Check "
                        f"{NETWORK_LOG} for the XHR endpoints and switch to direct API calls."
                    )
                raise RuntimeError("No clickable Leaflet rainfall station layers were found. See output/debug/NO_MARKERS.*")

            total = marker_locator.count()
            if DEBUG:
                total = min(total, DEBUG_MAX_STATIONS)
            if MAX_STATIONS:
                total = min(total, MAX_STATIONS)
            print(f"Found {total} marker(s) to process.")

            # CHANGED: index-shift-safe loop. The map may re-render, changing
            # DOM order, so we never trust an ordinal across iterations.
            # We walk indexes 0..count-1, skipping stations already processed,
            # and repeat full passes until one complete pass yields nothing new.
            processed = set()
            successful = failed = 0
            max_passes = 3
            for pass_no in range(1, max_passes + 1):
                new_in_pass = 0
                markers = get_marker_locator(page)
                if not markers:
                    break
                count = markers.count()
                if MAX_STATIONS and len(processed) >= MAX_STATIONS:
                    break
                for i in range(count):
                    if MAX_STATIONS and len(processed) >= MAX_STATIONS:
                        break
                    print(f"\n[pass {pass_no}][marker {i + 1}/{count}] clicking...")
                    try:
                        fresh = get_marker_locator(page)
                        if not fresh or i >= fresh.count():
                            raise RuntimeError(f"Marker index {i} no longer available (map re-rendered).")
                        station_name, station_index, rows = process_marker(page, fresh.nth(i))
                        if station_index in processed:
                            print(f"    Already processed {station_index}; skipping.")
                            close_station_panel(page)
                            continue
                        processed.add(station_index)
                        path = OUTPUT_DIR / safe_filename(station_index, station_name)
                        added = append_rows(path, rows)
                        print(f"    Rows kept (excluding today): {len(rows)}")
                        print(f"    New rows added: {added}")
                        print(f"    Excel: {path}")
                        successful += 1
                        new_in_pass += 1
                    except Exception as exc:
                        failed += 1
                        print(f"    FAILED: {exc}")
                        save_debug(page, f"ERROR_marker_p{pass_no}_{i:04d}")
                    finally:
                        close_station_panel(page)
                        time.sleep(0.4)
                if new_in_pass == 0:
                    print(f"Pass {pass_no}: no new stations; stopping.")
                    break

            print("\n======================================")
            print("Finished.")
            print(f"Successful stations: {successful}")
            print(f"Failed attempts: {failed}")
            print(f"Unique stations saved: {len(processed)}")
            print(f"Output folder: {OUTPUT_DIR.resolve()}")
            print(f"Network log (API endpoints): {NETWORK_LOG}")
            print("======================================")
        finally:
            try: netlog.close()
            except Exception: pass
            try: page.close()
            except Exception: pass
            try: context.close()
            except Exception: pass
            try: browser.close()
            except Exception: pass


if __name__ == "__main__":
    main()
