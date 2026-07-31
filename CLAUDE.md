# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Flask dashboard that ingests an Excel "RICE Tracker" workbook (Reports, Integrations, Conversions, Extensions) and renders an executive delivery view: summary cards, a delivery grid, a phase-shaded Gantt, a resource-capacity heatmap, risk panels, and a resource-planning simulator. No JS framework, no build step — server renders one HTML shell, a single `static/app.js` drives everything client-side against a JSON payload.

## Running the app

```bash
./run.sh                      # creates/reuses .venv, installs deps, runs on :5000
# or manually:
pip install -r requirements.txt
python app.py                 # http://127.0.0.1:5000
```

There is no test suite, linter, or build step in this repo — validate changes by running the app and hitting the routes/UI directly (e.g. `curl http://127.0.0.1:5000/api/data`, `curl http://127.0.0.1:5000/health`). For JS syntax checks: `node --check static/app.js`. For Python syntax checks: `python3 -c "import ast; ast.parse(open('app.py').read())"`.

## Data source

On startup the app looks for `rice_tracker_data.xlsx` in the repo root. If missing, the UI shows an upload screen; uploading via `/api/upload` validates and overwrites that file. There's no database — the workbook *is* the datastore, re-parsed from scratch on every `/api/data` request.

## Architecture

**Backend (`app.py`, single file) — three concerns, in order:**

1. **Ingestion** (`load_dataframe`, `_resolve_columns`, `process`): reads the workbook via pandas/openpyxl, resolves the source's real column headers to canonical keys via `COLUMN_MAP` using *normalized prefix matching* (`_resolve_columns`) — this is what makes the app tolerant of header drift (renamed columns, extra whitespace, trailing notes) across workbook revisions. All parsing goes through the safe parsers (`parse_date`, `parse_pct`, `parse_num`, `clean_str`), which never raise — bad input degrades to `None`/`""` instead of crashing a request. `process()` is the single ingestion entrypoint; it's re-run on every `/api/data` and `/api/resource-plan` call (no caching), so it must stay fast and side-effect-free.

2. **Derived fields computed per-record inside `process()`**: Gantt bar dates (`gantt_start`/`gantt_delivery`, with fallback duration from `build_hours` when actual dates are missing — see `HOURS_PER_DEV_WEEK`), sprint/phase assignment via `find_phase()` against `PROGRAM_TIMELINE_EXPANDED` (Actual → Planned → Unscheduled), and risk flags (`lean_spec_risk`, `build_risk`) from status keywords and date proximity to "today". **`PROGRAM_TIMELINE`** (near the top of `app.py`) is the hardcoded program calendar (Sprint 1/2/3, SIT 1/2, UAT, Cutover, Post Go-Live) — it drives Gantt phase shading, sprint assignment, and resource-plan deadlines; update it when the program plan changes. `_with_gaps()` auto-inserts synthetic "Gap b/w X and Y" phases between non-contiguous windows so every date lands somewhere.

3. **Resource planning** (`build_resource_plan` and helpers below `_plan_scope`): a separate, heavier pipeline invoked only by `/api/resource-plan`. Scope is fixed (`_plan_scope`: Deloitte-accountable, in-scope, module ≠ ARCS, RICE type ≠ Conversion). Work items get a priority score (`TYPE_PRIORITY` + `COMPLEXITY_PRIORITY`), then `_simulate()` runs a greedy list-scheduler across a synthetic dev roster on a business-day calendar. `_run_scenario()` wraps this in a hire-when-needed loop that adds resources until every deadline is met or hiring stops helping. Resources are **siloed by RICE type** (`_scenario_payload` groups items by `_type_key` and schedules each type against its own pool — devs never cross type pools). Four scenarios are produced (`aggressive`, `optimized`, `conservative`, `two_wave`), each just a different deadline/contingency/onboarding-rate config over the same scheduler.

Routes are thin wrappers at the bottom of the file: `/` (shell), `/api/data`, `/api/resource-plan`, `/api/upload`, `/api/export` (CSV), `/api/saved-filters` (GET/POST/DELETE, persisted to `saved_filters.json`), `/health`.

**Frontend (`static/app.js`, single file, no modules/bundler):**

- `State` (top of file) is the one mutable global: full payload, filtered records, Chart.js/AG Grid instances, active quick-filters, Choices.js instances.
- Boot sequence is the `DOMContentLoaded` handler near the top — `initTheme` → `initUpload` → `initTopbar` → `initFilters` → `initPlanToggle` → `initResourcePlan` → `initModal`, then `loadData()` fetches `/api/data` and calls `applyFilters()`, which is the fan-out point: it recomputes `State.filtered` and re-renders every panel (`renderTypeCards`, `renderGrid`, `renderGantt`, `renderCapacity`, `renderRisk`, `renderMatrix`, `renderDataQuality`, etc.). Any change to filtering logic needs `applyFilters()` to stay the single source of truth for what's currently visible — panels should not filter independently.
- Filtering is entirely client-side over the full record set already in `State.data`; there is no server-side filtering endpoint.
- The resource-planning UI (`initResourcePlan`/`loadResourcePlan`/`renderResourcePlan`/`renderRpChart`/`renderRpCompare`/`renderRpTimeline`) is a separate fetch/render cycle against `/api/resource-plan`, independent of the main filter/grid state.
- Colors/ordering conventions live as module-level constants near the top (`TYPE_ORDER`, `TYPE_COLOR`, `PHASE_FILL`, `statusColor`/`statusRank` keyword heuristics) — reuse these rather than hardcoding colors elsewhere, so the grid, Gantt, charts, and matrices stay visually consistent.
- Third-party libs (Chart.js, Choices.js, AG Grid Community) are loaded via CDN `<script>` tags in `templates/index.html`, not npm — there's no `package.json`.

## Key config knobs (in `app.py`)

- `HOURS_PER_DEV_WEEK` — drives capacity heatmap thresholds and Gantt/resource-plan fallback durations.
- `PROGRAM_TIMELINE` — the sprint/phase calendar; edit dates here when the program schedule shifts.
- `SIT1_START` / `SIT2_START`, `MAX_PLAN_RESOURCES`, `TYPE_PRIORITY`, `COMPLEXITY_PRIORITY` — resource-planning scenario inputs.
- Resource-plan runtime params (team size, ramp weeks/%, contingency %, buffer days) are query-string overrides handled in `_plan_params()`, not hardcoded.
