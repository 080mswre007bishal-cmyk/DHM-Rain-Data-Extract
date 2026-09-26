# DHM Rainfall Watch → Excel (free, automated, 100% browser-based)

Everything below happens in a web browser — GitHub's website. Nothing
is installed or run on your own computer; the actual scraping runs on
GitHub's servers, twice a day, automatically, forever, with no visit
required from you.

## Honesty first
DHM's rainfall page loads the real daily rainfall table with
JavaScript only after you select "Daily" and click "VIEW" — that
isn't visible in the plain page HTML, so a real (headless) browser is
used to click through each station's page the way a person would, then
reads whatever table appears, instead of guessing at a hidden API.
**Before you trust this, run the debug check in Step 4** — it's the
one part that may need adjusting, and it's a single button click.

## How it works
1. `discover_stations.py` — builds `stations.json` (station names,
   ids, coordinates) from DHM's plain HTML pages. Runs once, in the
   cloud, via the "Discover Stations" workflow.
2. `main.py` — for each station listed in `stations.json`, opens its
   own page (`dhm.gov.np/hydrology/rainfallSingle/<id>`), clicks the
   "Daily" tab, clicks VIEW, and reads the resulting table (Date |
   Daily rainfall (mm) | running Total shown on the page). Today's row
   is always skipped, since that figure is usually still accumulating.
   For each station's `output/<station>.xlsx`, any date already saved
   is skipped and only new dates are appended — so running it twice,
   or missing a scheduled run, never creates duplicates, and a missed
   day gets backfilled automatically the next time it runs, since each
   run re-reads the whole visible window (usually the last 8 days).
3. `.github/workflows/rainfall.yml` — runs `main.py` automatically
   twice a day, for free, on GitHub's servers, and attaches every
   run's `output/` folder (the Excel files) to that run as a
   downloadable artifact.
4. `.github/workflows/discover.yml` and `.github/workflows/debug.yml`
   — the two one-off cloud jobs described in Steps 2 and 4 below.

Google Drive upload is **not** currently wired in — output is
available by downloading the run artifact instead (Step 5). The
Drive-upload code from an earlier version can be re-added later if
wanted; ask and it can be dropped back in.

## Step 1 — Create the GitHub repo and add the files
1. On github.com, click **New repository**, give it a name, create it.
2. Use **Add file → Create new file**. For each file, type its exact
   path as the file name (this also creates the folders for you) and
   paste in the contents, then commit:
   - `main.py`
   - `discover_stations.py`
   - `requirements.txt`
   - `.github/workflows/rainfall.yml`
   - `.github/workflows/discover.yml`
   - `.github/workflows/debug.yml`
3. Go to **Settings → Actions → General → Workflow permissions** and
   select **Read and write permissions**, then Save. (This lets the
   discovery workflow commit `stations.json` back into your repo.)

## Step 2 — Build stations.json (one click, ~10-15 min)
Go to the **Actions** tab → **Discover Stations (run once)** → **Run
workflow**. When it finishes, `stations.json` will have been committed
straight into your repo automatically — nothing to download. Re-run
this any time DHM adds new stations you want picked up.

## Step 3 — (Optional) Adjust the schedule
The cron times in `rainfall.yml` are in UTC. Nepal is UTC+5:45, which
doesn't align cleanly to whole-minute UTC cron slots every day of the
year in the same way — the given times (03:15 and 15:15 UTC → 09:00
and 21:00 NPT) are close enough for a twice-daily check. Edit the
`cron:` line on GitHub (pencil icon on the file) if you want different
times.

## Step 4 — Sanity-check the scraper before trusting it (one click)
Go to **Actions** → **Debug Rainfall Extraction** → **Run workflow**
(leave `station_count` at its default of 10). When it finishes, open
the run's logs and check the printed `read N row(s): [...]` line for
each station, then download the **debug-output** artifact zip from the
bottom of the run page. For each station, open its matching `.txt`
file, find the real "Accumulated Rainfall" table text, and confirm the
(date, daily, total) values logged actually match what's really on the
page — right dates, right numbers, nothing missed or duplicated. If
it's off, open `main.py` on GitHub (pencil icon), find
`extract_rainfall_history()` and `table_rows_from_page()`, and adjust
using that `.txt` file as your reference. Commit the change and re-run
this workflow to confirm the fix.

## Step 5 — Go live (already running)
`.github/workflows/rainfall.yml` runs automatically on its schedule —
no further setup needed. You can also trigger it manually any time
from the **Actions** tab (**Run workflow** button) to test it end to
end. After each run, open that run's page and scroll to the bottom —
the **rainfall-output** artifact zip has every station's
`output/<station>.xlsx`, ready to download.

## Cost
Entirely free at this scale — GitHub Actions' free tier easily covers
two short runs a day (public repos get unlimited free minutes; private
repos get 2,000 free minutes/month). No server or hosting needed.
