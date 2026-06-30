# RICE Delivery Tracker

An executive-ready Flask web dashboard for tracking delivery of technology **RICE** objects — **R**eports, **I**ntegrations, **C**onversions, and **E**xtensions — across the full delivery lifecycle: Design → Build → SIT → UAT → Cutover → Go-Live.

The application ingests an Excel RICE tracker, computes derived delivery metrics on the server, and renders an interactive single-page dashboard with real-time filtering, charts, a delivery grid, and a phase-shaded Gantt timeline.

---

## Quick start

```bash
pip install -r requirements.txt
python app.py
```

Then open **http://127.0.0.1:5000** in your browser.

> Requires Python 3.9+.

---

## Data source

On startup the app looks for **`rice_tracker_data.xlsx`** in the application folder.

- **If the file exists**, the dashboard loads immediately.
- **If the file is missing**, you'll see a clean upload screen (drag-and-drop or file picker). Drop your RICE tracker workbook there and the app validates it, saves it as `rice_tracker_data.xlsx`, and reloads.

A working copy of the source workbook is bundled so the app runs out of the box. To use your own data, either replace `rice_tracker_data.xlsx` or upload through the UI.

### Ingestion is fault-tolerant

The backend uses pandas with safe parsing throughout, so malformed data never crashes the app:

- Dates are parsed with `errors="coerce"` and tolerate mixed formats.
- Percent fields like `"33% - In Progress"` are parsed to numeric progress.
- Missing columns are resolved by normalized header matching; absent columns degrade gracefully.
- Null and junk values are cleaned and defaulted.

---

## Features

**Executive summary** — Total / Completed / In Progress / Not Started / Delayed-Blocked cards, each with a RICE-type breakdown.

**RICE type cards** — Per-type totals with a Chart.js donut and percent-complete in the center.

**Raw object status** — True, ungrouped `Object Status` values straight from the source (e.g. FUT-Pending), with type breakdowns.

**Executive status mapping** — Raw statuses mapped to executive buckets (Blocked / Delayed / Completed / In Progress / Not Started / Other) with a transparent legend.

**Sprint summary** — Object counts and status breakdowns per program sprint/phase, assigned by actual → planned delivery date (else Unscheduled).

**Delivery plan grid** — AG Grid with frozen/pinned columns (RICE ID, Object Name, type badge, design & dev sprint breadcrumbs, complexity, build hours, functional owner), sorting, and virtual rendering.

**Gantt timeline** — Diamond = spec complete, bar = build window, dot = delivery, over phase-shaded columns (Sprint 1 green, Sprint 2 blue, Sprint 3 purple, SIT amber, UAT violet, Cutover red) with a yellow current-week overlay.

**Resource capacity heatmap** — Developers needed per week = `CEILING(weekly hours / 45)`, color-graded white → green → yellow → red; click a cell to see the underlying objects.

**Risk panels** — Lean spec risk and build risk lists derived from spec/dev status, dates, and progress.

**Dashboard matrices** — RICE Type × executive status and RICE Type × raw object status, with clickable cells that filter the grid.

**Data quality panel** — Counts of missing sprint, spec date, build hours, delivery date, owners, status, and RICE type.

**Global filters** — Multi-select Accountable Org and Module (Choices.js), single-select In Scope, and free-text search across RICE ID, object name, description, module, systems, and owners. All filtering and aggregation happen live in the browser.

**UX** — Sticky navigation, responsive layout, light (default) / dark theme toggle, and CSV export of the currently filtered set.

---

## Project structure

```
.
├── app.py                  # Flask app: ingestion, derived metrics, JSON/CSV/upload routes
├── requirements.txt        # Python dependencies
├── rice_tracker_data.xlsx  # Bundled source data (replaceable / uploadable)
├── templates/
│   └── index.html          # Upload screen + dashboard shell
└── static/
    ├── styles.css          # Deloitte palette, light/dark themes, all component styling
    └── app.js              # Filtering, aggregation, charts, grid, Gantt, heatmap, export
```

## Routes

| Route | Method | Purpose |
|-------|--------|---------|
| `/` | GET | Dashboard (or upload screen if no data) |
| `/api/data` | GET | Full processed record set + summary, filters, timeline, data quality (JSON) |
| `/api/upload` | POST | Validate and store an uploaded workbook |
| `/api/export` | GET | CSV export of the processed records |
| `/health` | GET | Health/status check |

## Configuration

- **`HOURS_PER_DEV_WEEK`** (default `45`) in `app.py` drives capacity and Gantt fallback duration.
- **`PROGRAM_TIMELINE`** in `app.py` defines the sprint/phase windows used for shading and sprint assignment.

## Tech stack

Flask · pandas · openpyxl on the backend; vanilla JS with Chart.js, AG Grid Community, and Choices.js (loaded via CDN) on the frontend. No build step and no JS framework required.
