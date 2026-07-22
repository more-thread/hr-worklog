# HR Weekly Report Generator

A single-page website that turns a weekly Excel export of **HR-project workload,
tasks, and progress per developer** into a visual, stakeholder-ready report —
with optional AI-written analysis by Claude.

Everything runs **in the browser**. There is no backend, no build step, and no
server to maintain — which is exactly why it can be hosted for free on GitHub
Pages.

## What it does

1. **Upload** a weekly `.xlsx` / `.xls` / `.csv`. It's parsed locally (SheetJS) —
   your data never leaves your machine.
2. **Auto-maps** your columns to the fields it understands; you can correct any
   mapping in the UI, so it works even if your headers differ from the template.
3. **Generates a dashboard**: KPI tiles (total tasks, % Done Development, ongoing,
   not-yet-started, failed-QA rework, team size), an overall-status donut,
   tasks-per-developer (stacked by status), tasks-by-request-type (total vs done),
   a completion-by-developer bar, and a per-developer breakdown table.
4. **Writes the analysis** — a weekly summary, improvement recommendations, and
   risk/blocker callouts. Without a key it uses built-in logic; with your Anthropic
   API key it uses **Claude** for a polished narrative and chart captions.
5. **Print / Save as PDF** for sharing (button top-right of the report).

## Expected Excel template (ONEHRMS weekly monitoring)

The tool is built around the ONEHRMS weekly monitoring template — a task sheet
with these columns:

| No. | Assigned To | Task Name | Module | Request Type | Progress | Started Development | Ended Development | Completed Date | Status |
|-----|-------------|-----------|--------|--------------|----------|---------------------|-------------------|----------------|--------|
| 1 | GRACIAS,ROLAND JOHN,POCSIDIO | 000125-0726-001 Form: EMPLOYEE PROFILE … | ONEHRMS | SYSTEM ENHANCEMENT | FOR QA | 46220 | 46225 | NULL | Done Development |

- **Assigned To** is the only required column. Names in `LAST,FIRST,MIDDLE`
  format are auto-cleaned to readable `First Last`.
- **Status** is normalized to the three ONEHRMS states: *Done Development*,
  *Ongoing*, *Not Yet Started*.
- **Progress** is treated as a workflow stage; `FAILED BY QA` is flagged as
  **rework** and `PAUSE`/`HOLD` as **paused** — both surface in the risk callouts.
- **Request Type** drives the "tasks by type" chart (mirroring your Summary's
  *Completed per Task Type*).
- Date columns accept Excel serials (e.g. `46220`) or dates; `NULL`/blank is fine.

### Multi-sheet workbooks

The real weekly file has four sheets — **Summary**, **Previous Week Progress**,
**Current Week Goals**, **Issues & Concerns**. The tool:

- **auto-selects** the data sheet (*Previous Week Progress*) — you can switch
  sheets with the picker if needed;
- automatically shows a **Goals for the coming week** section from *Current Week
  Goals*, and an **Issues & concerns** section from *Issues & Concerns*, when
  those sheets are present.

The blank template (`qwe.xlsx`) is a single sheet with the columns above and works
the same way. If a future template renames a column, just adjust the dropdowns in
the **Map your columns** step — no code changes needed. Click **Load sample data**
in the app to see the exact shape it expects.

## Example file

`examples/sample-onehrms-week.csv` is a **synthetic** sample (fake names, same
structure) you can drag into the app to see a full report immediately — safe to
commit and publish.

> ⚠️ **Real spreadsheets are not committed.** The `.gitignore` excludes `*.xlsx`
> so real ONEHRMS files (which contain real employee names and internal task
> details) are never pushed to this public repo. Drop your real file in to test
> locally; it stays ignored. Only publish a file you've confirmed is safe to make
> public.

## Reporting period

If the workbook has a **Summary** sheet with `Start Date` / `End Date` cells (as
the ONEHRMS weekly file does), the report header shows the reporting period
automatically (e.g. *Reporting period Jul 21 – Jul 26, 2026*).

## Run it locally

Because it loads local `.js` files, open it through a tiny web server (opening the
file directly with `file://` can block the scripts in some browsers):

```bash
cd hr-weekly-report
python3 -m http.server 8137
# then visit http://localhost:8137
```

## Deploy to GitHub Pages (free)

### Option A — GitHub web UI (no command line)

1. Create a new repository on GitHub, e.g. `hr-weekly-report`.
2. Click **Add file → Upload files**, drag in **all** the files/folders from this
   project (`index.html`, `css/`, `js/`), and commit.
3. Go to **Settings → Pages**.
4. Under **Build and deployment → Source**, choose **Deploy from a branch**.
5. Set branch to **`main`** and folder to **`/ (root)`**, then **Save**.
6. Wait ~1 minute. Your site is live at
   `https://<your-username>.github.io/hr-weekly-report/`.

### Option B — Git command line

```bash
cd hr-weekly-report
git init
git add .
git commit -m "HR weekly report generator"
git branch -M main
git remote add origin https://github.com/<your-username>/hr-weekly-report.git
git push -u origin main
```

Then follow steps 3–6 above (Settings → Pages).

> The `js/xlsx.full.min.js` and `js/chart.umd.min.js` libraries are vendored
> (committed into the repo), so the site works offline and needs no CDN.

## Using Claude for AI analysis

The AI narrative is **optional and opt-in**:

1. In the app, open **Claude AI settings**.
2. Paste your Anthropic API key (get one at
   [console.anthropic.com](https://console.anthropic.com/settings/keys)).
3. Generate a report, then click **✨ Generate with Claude**.

**About the key & security:** GitHub Pages only serves static files, so there's
nowhere to safely store a shared secret key. Instead, each user supplies *their
own* key at runtime. The key is:

- kept **only in your browser** (in memory, or in `localStorage` if you tick
  "Remember key on this device"),
- **never** committed to the repo or sent anywhere except directly to Anthropic's
  API,
- sent from the browser using Anthropic's
  `anthropic-dangerous-direct-browser-access` header (direct browser calls).

If you'd rather users never handle a key, the alternative is a small serverless
proxy (Cloudflare Workers / Netlify Functions) that holds the key server-side —
that requires a second free account and is a larger setup than pure GitHub Pages.

Without a key, the report still generates using the built-in (non-AI) analysis.

## Privacy

- Spreadsheet parsing and all chart rendering happen entirely client-side.
- Only when you explicitly click **Generate with Claude** is an *aggregated*
  summary of the data (per-developer/per-project totals and flagged tasks — not
  the raw file) sent to Anthropic's API to write the narrative.

## Tech

- [SheetJS](https://sheetjs.com) — Excel/CSV parsing
- [Chart.js](https://www.chartjs.org) — charts
- Anthropic Messages API — optional AI analysis
- Vanilla HTML/CSS/JS — no framework, no build
