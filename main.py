import datetime as dt
import os
import re
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

MAP_URL = "https://dhm.gov.np/hydrology/rainfall-watch-map"
OUTPUT_DIR = Path("output")
DEBUG_DIR = OUTPUT_DIR / "debug"

DEBUG = os.getenv("DEBUG", "1") == "1"
HEADLESS = os.getenv("HEADLESS", "0") != "0"
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
        for fmt in DATE_FORMATS:
            try:
                return dt.datetime.strptime(m.group(0), fmt).strftime("%Y-%m-%d")
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
            (DEBUG_DIR / f"{safe}_panel.txt").write_text(panel.inner_text(), encoding="utf-8")
        except Exception:
            pass
        try:
            (DEBUG_DIR / f"{safe}_panel.html").write_text(panel.evaluate("el => el.outerHTML"), encoding="utf-8")
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
            page.wait_for_timeout(3500)
            return
        except Exception:
            pass
    raise RuntimeError("Could not click DHM '24 Hour'.")


def get_marker_locator(page):
    for selector in ("#map .leaflet-marker-icon", ".leaflet-container .leaflet-marker-icon"):
        loc = page.locator(selector)
        try:
            if loc.count() > 0:
                return loc
        except Exception:
            pass
    return None


def _xpath_for_element_js():
    return """
    el => {
      function visible(e) {
        const s = getComputedStyle(e);
        const r = e.getBoundingClientRect();
        return s.display !== 'none' && s.visibility !== 'hidden' &&
               parseFloat(s.opacity || '1') > 0 && r.width > 0 && r.height > 0;
      }
      function xpath(e) {
        if (e.id) return '//*[@id="' + e.id.replace(/"/g, '\\"') + '"]';
        const parts = [];
        while (e && e.nodeType === 1) {
          let n = 1;
          let sib = e.previousElementSibling;
          while (sib) { if (sib.tagName === e.tagName) n++; sib = sib.previousElementSibling; }
          parts.unshift(e.tagName.toLowerCase() + '[' + n + ']');
          e = e.parentElement;
        }
        return '/' + parts.join('/');
      }
      let best = null;
      for (const e of document.querySelectorAll('body *')) {
        if (!visible(e)) continue;
        if (e.closest('.leaflet-container')) continue;
        const t = (e.innerText || '').replace(/\\s+/g, ' ').trim();
        const low = t.toLowerCase();
        if (!low.includes('station index') || !low.includes('view')) continue;
        if (!low.includes('period') && !low.includes('observations')) continue;
        if (t.length > 6000) continue;
        if (!best || t.length < best.text.length) best = {text:t, xp:xpath(e)};
      }
      return best;
    }
    """


def find_station_panel(page):
    """Find the station interface OUTSIDE the Leaflet map."""
    # First try the main page.
    try:
        found = page.evaluate(_xpath_for_element_js())
        if found:
            return page.locator(f"xpath={found['xp']}")
    except Exception:
        pass

    # Also support a station interface rendered inside an iframe.
    for frame in page.frames:
        if frame == page.main_frame:
            continue
        try:
            found = frame.evaluate(_xpath_for_element_js())
            if found:
                return frame.locator(f"xpath={found['xp']}")
        except Exception:
            pass
    return None


def station_name_and_index(panel, sequence):
    text = re.sub(r"[ \t]+", " ", panel.inner_text()).strip()

    # Prefer the exact station label used by the DHM station interface.
    name = None
    patterns = [
        r"Station\s*:\s*(.+?)(?=\s*\|\s*District\s*:|\n|$)",
        r"Station\s*Name\s*:\s*(.+?)(?=\s*\|\s*District\s*:|\n|$)",
    ]
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            name = m.group(1).strip()
            break

    if not name:
        lines = [re.sub(r"\s+", " ", x).strip() for x in panel.inner_text().splitlines() if x.strip()]
        for i, line in enumerate(lines[:-1]):
            if re.fullmatch(r"Station(?:\s+Name)?", line, re.I):
                name = lines[i + 1]
                break

    idx = None
    m = re.search(r"Station\s*Index\s*[:\-]?\s*([A-Za-z0-9_-]+)", text, re.I)
    if m:
        idx = m.group(1).strip()

    if not idx:
        lines = [re.sub(r"\s+", " ", x).strip() for x in panel.inner_text().splitlines() if x.strip()]
        for i, line in enumerate(lines[:-1]):
            if re.fullmatch(r"Station\s*Index", line, re.I):
                idx = lines[i + 1].strip()
                break

    if not name:
        name = f"Station_{sequence:04d}"
    if not idx:
        raise RuntimeError("Actual Station Index was not found in the station interface.")
    return name[:180], idx


def set_period_daily(panel):
    selects = panel.locator("select")
    for i in range(selects.count()):
        s = selects.nth(i)
        try:
            options = s.locator("option").all_inner_texts()
            if any(x.strip().lower() == "daily" for x in options):
                s.select_option(label="Daily")
                return True
        except Exception:
            pass
    # Some versions may expose the select just outside the smallest panel.
    try:
        select = panel.locator("xpath=ancestor::*[.//select][1]").locator("select").first
        if select.count() and any(x.strip().lower() == "daily" for x in select.locator("option").all_inner_texts()):
            select.select_option(label="Daily")
            return True
    except Exception:
        pass
    return False


def click_view(panel):
    candidates = [
        panel.get_by_role("button", name="VIEW", exact=True).first,
        panel.locator("button:has-text('VIEW')").first,
        panel.locator("input[value='VIEW']").first,
        panel.locator("a:has-text('VIEW')").first,
    ]
    for loc in candidates:
        try:
            loc.wait_for(state="visible", timeout=2500)
            loc.click(force=True)
            return True
        except Exception:
            pass
    return False


def read_table(panel):
    """Read all Daily rows, then remove today's row only."""
    today = dt.date.today().strftime("%Y-%m-%d")
    tables = panel.locator("table")
    for i in range(tables.count()):
        table = tables.nth(i)
        try:
            rows = table.locator("tr")
            header = None
            header_map = {}
            for r in range(rows.count()):
                cells = rows.nth(r).locator("th,td").all_inner_texts()
                norm = [re.sub(r"\s+", " ", x).strip().lower() for x in cells]
                if any(x == "date" for x in norm) and any(x == "daily" for x in norm):
                    header = norm
                    for j, x in enumerate(norm):
                        if x in {"date", "daily", "total"}:
                            header_map[x] = j
                    break
            if not header or "date" not in header_map or "daily" not in header_map:
                continue

            extracted = []
            for r in range(rows.count()):
                values = [x.strip() for x in rows.nth(r).locator("th,td").all_inner_texts()]
                if len(values) <= max(header_map.values()):
                    continue
                d = parse_date(values[header_map["date"]])
                if not d or d == today:
                    continue
                daily = parse_number(values[header_map["daily"]])
                if daily is None:
                    continue
                total = parse_number(values[header_map["total"]]) if "total" in header_map else None
                extracted.append((d, daily, total))
            if extracted:
                return extracted
        except Exception:
            pass
    return []


def wait_for_table(panel):
    for _ in range(40):
        rows = read_table(panel)
        if rows:
            return rows
        time.sleep(0.5)
    return []


def safe_filename(station_index, station_name):
    clean_idx = re.sub(r"[^A-Za-z0-9_-]+", "_", str(station_index)).strip("_")
    clean_name = re.sub(r"[^\w\-]+", "_", station_name).strip("_") or "Unknown_Station"
    return f"{clean_idx}_{clean_name}.xlsx"


def append_rows(path, rows):
    if path.exists():
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Daily Rainfall"
        ws.append(["Date", "Daily Rainfall (mm)", "Running Total on page (mm)"])

    existing = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0] is not None:
            v = row[0]
            if isinstance(v, dt.datetime):
                existing.add(v.strftime("%Y-%m-%d"))
            elif isinstance(v, dt.date):
                existing.add(v.strftime("%Y-%m-%d"))
            else:
                p = parse_date(str(v))
                if p:
                    existing.add(p)

    added = 0
    for d, daily, total in sorted(rows):
        if d in existing:
            continue
        ws.append([d, daily, total])
        ws.cell(ws.max_row, 1).number_format = "yyyy-mm-dd"
        existing.add(d)
        added += 1
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)
    wb.close()
    return added


def close_station_panel(page, panel):
    if panel is None:
        return
    # Try common close controls inside the outside-map interface.
    for loc in [
        panel.locator("button[aria-label='Close']").first,
        panel.locator(".btn-close").first,
        panel.locator("button.close").first,
        panel.get_by_role("button", name=re.compile(r"close", re.I)).first,
    ]:
        try:
            if loc.is_visible():
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


def process_marker(page, marker, sequence):
    marker.scroll_into_view_if_needed(timeout=3000)
    try:
        marker.click(force=True, timeout=5000)
    except Exception:
        marker.dispatch_event("click")

    panel = None
    for _ in range(40):
        panel = find_station_panel(page)
        if panel is not None:
            break
        time.sleep(0.25)

    if panel is None:
        save_debug(page, f"ERROR_marker_{sequence:04d}_no_station_panel")
        raise RuntimeError("Marker clicked, but the OUTSIDE-map station interface was not found.")

    save_debug(page, f"marker_{sequence:04d}_station_panel", panel)

    station_name, station_index = station_name_and_index(panel, sequence)
    print(f"    Station: {station_name}")
    print(f"    Station Index: {station_index}")

    if not set_period_daily(panel):
        raise RuntimeError("Period dropdown was not found or could not be set to Daily.")

    if not click_view(panel):
        raise RuntimeError("VIEW button was not found in the outside-map station interface.")

    rows = wait_for_table(panel)
    if not rows:
        save_debug(page, f"ERROR_marker_{sequence:04d}_no_daily_rows", panel)
        raise RuntimeError("VIEW completed but no Daily rainfall rows were found after removing today's row.")

    return station_name, station_index, rows, panel


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print("DHM Daily Rainfall Extractor")
    print(f"Today: {dt.date.today():%Y-%m-%d}")
    print("Date field is NOT modified.")
    print("Period is set to Daily, then VIEW is clicked.")
    print("All returned rows are saved EXCEPT today's row.")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, args=["--no-sandbox"])
        context = browser.new_context(viewport={"width": 1440, "height": 1000}, locale="en-US")
        page = context.new_page()
        try:
            click_24_hour(page)
            for _ in range(40):
                markers = get_marker_locator(page)
                if markers and markers.count() > 0:
                    break
                page.wait_for_timeout(500)
            markers = get_marker_locator(page)
            if not markers:
                raise RuntimeError("No Leaflet station markers found.")
            total = markers.count()
            limit = DEBUG_MAX_STATIONS if DEBUG else (min(total, MAX_STATIONS) if MAX_STATIONS else total)
            limit = min(total, limit)
            print(f"Found {total} Leaflet marker(s); processing {limit}.")

            success = failed = 0
            for sequence in range(1, limit + 1):
                panel = None
                print(f"\n[{sequence}/{limit}] Clicking Leaflet station marker...")
                try:
                    markers = get_marker_locator(page)
                    if not markers or sequence > markers.count():
                        raise RuntimeError("Leaflet marker count changed while processing.")
                    station_name, station_index, rows, panel = process_marker(page, markers.nth(sequence - 1), sequence)
                    path = OUTPUT_DIR / safe_filename(station_index, station_name)
                    added = append_rows(path, rows)
                    print(f"    Rows saved: {len(rows)}; new rows: {added}")
                    print(f"    Excel: {path}")
                    success += 1
                except Exception as exc:
                    failed += 1
                    print(f"    FAILED: {exc}")
                    save_debug(page, f"ERROR_marker_{sequence:04d}", panel)
                finally:
                    close_station_panel(page, panel)
                    time.sleep(0.5)

            print("\nFinished")
            print(f"Successful: {success}")
            print(f"Failed: {failed}")
            print(f"Output: {OUTPUT_DIR.resolve()}")
        finally:
            page.close()
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
