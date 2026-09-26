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
            (DEBUG_DIR / f"{safe}_panel.html").write_text(panel.evaluate("el => el.outerHTML"), encoding="utf-8")
            (DEBUG_DIR / f"{safe}_panel.txt").write_text(panel.inner_text(), encoding="utf-8")
        except Exception:
            pass


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
    for selector in ["#map .leaflet-marker-icon", ".leaflet-container .leaflet-marker-icon"]:
        loc = page.locator(selector)
        try:
            if loc.count() > 0:
                return loc
        except Exception:
            pass
    return None


def _candidate_panels(root):
    selectors = [
        "[role='dialog']:visible",
        ".modal.show:visible", ".modal.in:visible",
        ".offcanvas.show:visible",
        "div[class*='station']:visible", "section[class*='station']:visible",
        "div[id*='station']:visible", "section[id*='station']:visible",
        "div[class*='rainfall']:visible", "section[class*='rainfall']:visible",
        "div[class*='observation']:visible", "section[class*='observation']:visible",
        "div[class*='single']:visible", "section[class*='single']:visible",
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
                    key = el.evaluate("el => el").__hash__() if False else None
                    # Exclude anything physically inside the Leaflet map.
                    inside_map = el.evaluate("el => !!el.closest('.leaflet-container')")
                    if inside_map:
                        continue
                    txt = el.inner_text(timeout=1000)
                    low = txt.lower()
                    if "station index" in low and "view" in low:
                        # Deduplicate by a stable DOM path-ish signature.
                        sig = (el.evaluate("el => el.tagName + '|' + (el.id || '') + '|' + (el.className || '')"), len(txt))
                        if sig not in seen:
                            seen.add(sig)
                            result.append(el)
                except Exception:
                    pass
        except Exception:
            pass
    return result


def visible_station_panel(page):
    panels = _candidate_panels(page)
    if not panels:
        return None
    # Prefer the smallest candidate: outer containers often contain the same content.
    ranked = []
    for p in panels:
        try:
            text = p.inner_text()
            ranked.append((len(text), p))
        except Exception:
            pass
    ranked.sort(key=lambda x: x[0])
    return ranked[0][1] if ranked else None


def station_details(panel, sequence):
    text = " ".join(panel.inner_text().split())
    m = re.search(r"Station\s*:\s*(.+?)(?:\s*\|\s*District\s*:|\s+District\s*:|$)", text, re.I)
    station_name = m.group(1).strip() if m else None
    if not station_name:
        m = re.search(r"Station\s*Name\s*[:\-]\s*(.+?)(?:\s+Station\s+Index|$)", text, re.I)
        station_name = m.group(1).strip() if m else None
    idx = re.search(r"Station\s*Index\s*[:\-]?\s*([A-Za-z0-9_-]+)", text, re.I)
    station_index = idx.group(1).strip() if idx else None
    if not station_name:
        raise RuntimeError(f"Station name not found in station interface for marker {sequence}.")
    if not station_index:
        raise RuntimeError(f"Actual Station Index not found in station interface for marker {sequence}.")
    return station_name[:180], station_index


def set_period_daily(panel):
    selects = panel.locator("select")
    for i in range(selects.count()):
        sel = selects.nth(i)
        try:
            opts = sel.locator("option").all_inner_texts()
            if any(o.strip().lower() == "daily" for o in opts):
                try:
                    sel.select_option(label="Daily")
                except Exception:
                    sel.select_option(value="3")
                selected = sel.locator("option:checked").inner_text().strip()
                if selected.lower() != "daily":
                    raise RuntimeError(f"Period did not become Daily; selected={selected!r}")
                return
        except Exception:
            pass
    raise RuntimeError("Daily option was not found in the station interface.")


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


def read_table(panel):
    tables = panel.locator("table")
    for i in range(tables.count()):
        table = tables.nth(i)
        try:
            rows = table.locator("tr")
            if rows.count() == 0:
                continue
            header = [x.strip().lower() for x in rows.nth(0).locator("th,td").all_inner_texts()]
            if "date" not in header or "daily" not in header:
                continue
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
            if out:
                return out
        except Exception:
            pass
    return []


def extract_rainfall(panel):
    set_period_daily(panel)
    click_view(panel)
    for _ in range(40):
        rows = read_table(panel)
        if rows:
            # IMPORTANT: remove only today's row; retain every other returned date.
            return [row for row in rows if row[0] != today_iso()]
        time.sleep(0.5)
    raise RuntimeError("No Daily rainfall table appeared after clicking VIEW.")


def close_station_panel(page):
    selectors = [
        "[aria-label='Close']:visible", "button[title='Close']:visible",
        ".btn-close:visible", ".close:visible",
        "button:has-text('Close'):visible",
    ]
    for selector in selectors:
        try:
            loc = page.locator(selector).last
            if loc.is_visible():
                loc.click(force=True)
                page.wait_for_timeout(250)
                return
        except Exception:
            pass
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(250)
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


def process_marker(page, marker, sequence):
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
        raise RuntimeError("Marker clicked, but no outside-map station interface appeared.")

    station_name, station_index = station_details(panel, sequence)
    print(f"    Station: {station_name}")
    print(f"    Station Index: {station_index}")
    if DEBUG:
        save_debug(page, f"marker_{sequence:04d}_station_panel", panel)

    rows = extract_rainfall(panel)
    if DEBUG:
        save_debug(page, f"marker_{sequence:04d}_after_view", panel)
    return station_name, station_index, rows


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print("======================================")
    print("DHM Daily Rainfall Extractor")
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
        context = browser.new_context(viewport={"width": 1440, "height": 1000}, locale="en-US")
        page = context.new_page()
        try:
            click_24_hour(page)
            marker_locator = None
            for _ in range(40):
                marker_locator = get_marker_locator(page)
                if marker_locator and marker_locator.count() > 0:
                    break
                page.wait_for_timeout(500)
            if not marker_locator:
                raise RuntimeError("No Leaflet station markers were found.")
            total = marker_locator.count()
            if DEBUG:
                total_to_process = min(total, DEBUG_MAX_STATIONS)
            elif MAX_STATIONS:
                total_to_process = min(total, MAX_STATIONS)
            else:
                total_to_process = total
            print(f"Found {total} marker(s); processing {total_to_process}.")

            successful = failed = 0
            for sequence in range(1, total_to_process + 1):
                print(f"\n[{sequence}/{total_to_process}] Clicking marker {sequence}...")
                try:
                    markers = get_marker_locator(page)
                    if not markers or sequence > markers.count():
                        raise RuntimeError(f"Marker {sequence} is no longer available.")
                    station_name, station_index, rows = process_marker(page, markers.nth(sequence - 1), sequence)
                    path = OUTPUT_DIR / safe_filename(station_index, station_name)
                    added = append_rows(path, rows)
                    print(f"    Rows kept (excluding today): {len(rows)}")
                    print(f"    New rows added: {added}")
                    print(f"    Excel: {path}")
                    successful += 1
                except Exception as exc:
                    failed += 1
                    print(f"    FAILED: {exc}")
                    save_debug(page, f"ERROR_marker_{sequence:04d}")
                finally:
                    close_station_panel(page)
                    time.sleep(0.4)

            print("\n======================================")
            print("Finished.")
            print(f"Successful stations: {successful}")
            print(f"Failed stations: {failed}")
            print(f"Output folder: {OUTPUT_DIR.resolve()}")
            print("======================================")
        finally:
            try: page.close()
            except Exception: pass
            try: context.close()
            except Exception: pass
            try: browser.close()
            except Exception: pass


if __name__ == "__main__":
    main()
