# DHM Rainfall Watch → Excel → Google Drive (free, automated, 100% browser-based)

Everything below happens in a web browser — GitHub's website, Google
Cloud Console, Google Drive. Nothing is installed or run on your own
computer; the actual scraping runs on GitHub's servers.

## Honesty first
DHM's rainfall map loads the actual 24-hour rainfall number with
JavaScript after you click "VIEW" — that request isn't visible in the
plain page HTML. So this uses a real (headless) browser on GitHub's
servers to click through each station's page the way a person would,
then reads whatever number appears, instead of guessing at a hidden
API. **Before you trust this, run the debug check in Step 4** — it's
the one part that may need adjusting, and it's a single button click.

## How it works
1. `discover_stations.py` — builds `stations.json` (station names,
   ids, coordinates) from DHM's plain HTML pages. Runs once, in the
   cloud, via the "Discover Stations" workflow.
2. `main.py` — for each station, opens its page, selects "Daily",
   clicks VIEW, and reads the last 8 days of rainfall figures shown
   there. It skips any date already saved in `output/<station>.xlsx`
   and appends only the new ones, then uploads/updates that file in a
   Google Drive folder — only when there's something new to add. Since
   each run re-reads the last 8 days, a missed scheduled run (site
   down, GitHub hiccup, etc.) gets backfilled automatically the next
   time it runs successfully.
3. `.github/workflows/rainfall.yml` — runs `main.py` automatically
   twice a day, for free, on GitHub's servers.
4. `.github/workflows/discover.yml` and `.github/workflows/debug.yml`
   — the two one-off cloud jobs described below.

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
straight into your repo automatically — nothing to download.

## Step 3 — Google Drive access (free, all in your browser)
1. In Google Cloud Console, create a project (or use an existing one)
   and enable the **Google Drive API**.
2. Create a **Service Account**, then create a JSON key for it and
   download the file (this is the one file you'll need to copy from —
   you can open it in any text viewer, you don't need to run anything).
3. In Google Drive, create the folder where you want the per-station
   Excel files to live, open its sharing settings, and share it with
   the service account's email address (looks like
   `something@your-project.iam.gserviceaccount.com`) as **Editor**.
4. Copy the folder's ID from its URL:
   `https://drive.google.com/drive/folders/`**`THIS_PART`**

## Step 4 — Sanity-check the scraper before trusting it (one click)
Go to **Actions** → **Debug Rainfall Extraction** → **Run workflow**
(leave station_count at 3). When it finishes, open the run's logs and
check the printed `read N day(s): [...]` line for each station against
the **debug-output** artifact zip (download it from the bottom of the
run page). For each station, open its `.txt` file, find the real
8-day table text, and confirm the (date, value) pairs logged actually
match what's on the page — right dates, right numbers, nothing missed
or duplicated. If it's off, open `main.py` on GitHub (pencil icon),
find `extract_rainfall_history()`, and adjust the date/number parsing
using that `.txt` file as your reference. Commit the change and re-run
the debug workflow to confirm the fix.

## Step 5 — Add secrets and go live
In your repo's **Settings → Secrets and variables → Actions**, add:
- `GOOGLE_SERVICE_ACCOUNT_JSON` — paste the *entire contents* of the
  service account JSON key file from Step 3.
- `DRIVE_FOLDER_ID` — the folder ID from Step 3.

That's it — `.github/workflows/rainfall.yml` will run on its schedule
(twice daily), and you can also trigger it manually from the **Actions**
tab (**Run workflow** button) any time to test it end to end.

## Adjusting the schedule
The cron times in `rainfall.yml` are in UTC. Nepal is UTC+5:45, which
doesn't align cleanly to whole-minute UTC cron slots every day of the
year in the same way — the given times (03:15 and 15:15 UTC → 09:00
and 21:00 NPT) are close enough for a twice-daily check. Edit the
`cron:` line on GitHub (pencil icon on the file) if you want different
times.

## Cost
Entirely free at this scale:
- GitHub Actions: free tier easily covers this (public repos get
  unlimited free minutes; private repos get 2,000 free minutes/month).
- Google Drive API: free.
- No server or hosting needed — GitHub runs everything for you.
