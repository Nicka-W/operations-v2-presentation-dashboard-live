# Operations V2 Dashboard — Setup

Same architecture as your live Operations dashboard: local Excel workbook →
`dashboard_data.json` → GitHub Pages, refreshed every minute by a Windows
scheduled task.

Everything in this folder was generated for you. The steps below are the
parts that have to run on your own machine (I can't push to GitHub,
register a Windows scheduled task, or touch your local Excel file from here).

## 1. Folder layout

Copy this whole folder somewhere under your OneDrive Operations area, e.g.:

```
C:\Users\Dell\OneDrive\002.Operations\Dashboard\Operations-V2\
  code\operations-v2-presentation-dashboard-live\   <- this repo
  Operations Data.xlsx                               <- your branch's workbook
  Backups\
```

## 2. Point the parser at your workbook

Open `scripts/refresh_dashboard_data.py` and update `WORKBOOK_PATH` near the
top to your actual workbook location, e.g.:

```python
WORKBOOK_PATH = r"C:\Users\Dell\OneDrive\002.Operations\Dashboard\Operations-V2\Operations Data.xlsx"
```

**Important — verify the layout before trusting the output.** I built the
column mapping (`WORKBOOK_LAYOUT` in that file) against the exact workbook
you uploaded. If the branch/department workbook you swap in later has any
columns shifted, renamed, or reordered from that layout, update
`WORKBOOK_LAYOUT` to match — don't assume it lines up. Two things already
differed from the original handover doc in your file, so it's worth an
actual check rather than an assumption:
- `sku-monthly` has **two** branch columns (`2026 CPT` and `2026 GEORGE`),
  not one.
- `assembly-backorders` sits at columns `AD:AG`, one column right of what
  the original doc described.

## 3. Create the GitHub repo

1. Create a new GitHub repo, e.g. `operations-v2-presentation-dashboard-live`.
2. Enable GitHub Pages for it (Settings → Pages → deploy from `main`).
3. Push this folder's contents to `main`.

```
cd operations-v2-presentation-dashboard-live
git init
git remote add origin <your-repo-url>
git add .
git commit -m "Initial Operations V2 dashboard"
git branch -M main
git push -u origin main
```

## 4. Test the parser and publisher manually first

```
python scripts\refresh_dashboard_data.py
```

Confirms `dashboard_data.json` builds correctly. Open `index.html` locally
in a browser to sanity-check the KPI cards, table, and chart views before
wiring up auto-publish.

Then test the full publish flow (requires the repo to already be pushed and
clean):

```
python scripts\publish_dashboard_data.py
```

## 5. Register the scheduled task

In an **elevated** PowerShell prompt:

```
cd scripts
.\register_local_autopublish.ps1
```

This creates a task named `Operations V2 PPT Dashboard Auto Publish` that
runs every minute. Check its status any time with:

```
schtasks /Query /TN "Operations V2 PPT Dashboard Auto Publish" /FO LIST /V
```

## 6. Health checks

- **Scheduled task**: `Last Result` should read `0`.
- **Local log**: `scripts\local_autopublish.log`
- **Live JSON**: `https://<your-username>.github.io/<repo-name>/dashboard_data.json`
  — check `generatedAt` and `sourceModifiedAt` are recent.
- **Repo cleanliness**: if the repo has uncommitted changes,
  `publish_dashboard_data.py` will refuse to run (same as the original
  dashboard) — commit or stash before publishing continues automatically.

## Known gotchas (carried over from the original dashboard)

- The **Refresh Now** button on the site only reloads the already-published
  `dashboard_data.json` — it does not rebuild from Excel. If the site looks
  stale, check the scheduled task and log first.
- If the workbook gets renamed, `find_workbook_fallback()` in
  `refresh_dashboard_data.py` will auto-detect a single `.xlsx` in the same
  folder — but only if there's exactly one candidate file there.
- Keep Excel formulas in `V` (share) and the fill-rate column extended to
  row 100 if you add more rows of data, matching `LIVE_DATA_MAX_ROW`.
- The snapshot helper (`save_excel_snapshot.ps1`) always opens its own
  hidden Excel instance rather than attaching to your live session — don't
  change that, it's what avoids the `0x800AC472` busy errors and
  interference with your other open spreadsheets that broke the original
  setup.

## Color rules (currently same as Operations)

| KPI | Red | Orange | Green |
|---|---|---|---|
| Housekeeping | < 50% | 50–79% | ≥ 80% |
| JHB / George / Dispatch Accuracy | ≤ 97% | 97–98% | ≥ 99% |
| Urgent Orders | > 10% | 5–10% | ≤ 5% |
| Assembly Fill Rate | < 70% | 70–84% | ≥ 85% |

You said you'd give me new thresholds later if this branch needs different
cutoffs — just update `KPI_COLOR_RULES` in `refresh_dashboard_data.py` and
the matching logic in `index.html`'s `colorFor()` function when you're
ready.
