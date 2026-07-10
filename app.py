"""
RICE Tracker — Executive Delivery Dashboard
Flask backend with robust pandas ingestion of the Overall RICE Tracker workbook.

Run:
    pip install -r requirements.txt
    python app.py
Then open http://127.0.0.1:5000
"""

import os
import io
import re
import json
import math
import datetime as dt

import numpy as np
import pandas as pd
from flask import Flask, jsonify, request, render_template, send_file

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
DATA_FILE = os.path.join(BASE_DIR, "rice_tracker_data.xlsx")
SAVED_FILTERS_FILE = os.path.join(BASE_DIR, "saved_filters.json")
SHEET_CANDIDATES = ["1.4.2 - Overall Rice Tracker_rb"]  # preferred sheet names
ALLOWED_EXT = {".xlsx", ".xlsm", ".xls"}
HOURS_PER_DEV_WEEK = 45.0

os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload cap

# Program timeline (fixed configuration from the program plan).
PROGRAM_TIMELINE = [
    {"name": "Sprint 1", "type": "Sprint", "status": "Completed", "start": "2026-03-23", "end": "2026-05-01"},
    {"name": "Sprint 2", "type": "Sprint", "status": "In Progress", "start": "2026-06-22", "end": "2026-07-17"},
    {"name": "Sprint 3", "type": "Sprint", "status": "Planned", "start": "2026-07-27", "end": "2026-08-21"},
    {"name": "SIT 1", "type": "SIT", "status": "Planned", "start": "2026-09-28", "end": "2026-10-30"},
    {"name": "SIT 2", "type": "SIT", "status": "Planned", "start": "2026-11-09", "end": "2026-12-11"},
    {"name": "UAT", "type": "UAT", "status": "Planned", "start": "2026-12-14", "end": "2027-01-15"},
    {"name": "Cutover", "type": "Cutover", "status": "Planned", "start": "2027-01-18", "end": "2027-02-07"},
    # Open-ended catch-all: anything delivered after Cutover lands here rather
    # than falling through to Unscheduled. "end" is a far-future sentinel, not
    # a real program date — open_ended tells the UI to render/scale it as "onward".
    {"name": "Post Go-Live", "type": "Milestone", "status": "Planned", "start": "2027-02-08", "end": "2099-12-31", "open_ended": True},
]


def _with_gaps(program_timeline):
    """Insert a synthetic "Gap b/w X and Y" phase between any two consecutive
    phases that aren't back-to-back, so a date landing between cycles gets its
    own tile instead of silently rolling into the next phase."""
    expanded = []
    for i, phase in enumerate(program_timeline):
        expanded.append(phase)
        if i + 1 >= len(program_timeline):
            continue
        nxt = program_timeline[i + 1]
        gap_start = pd.Timestamp(phase["end"]) + pd.Timedelta(days=1)
        gap_end = pd.Timestamp(nxt["start"]) - pd.Timedelta(days=1)
        if gap_start <= gap_end:
            expanded.append({
                "name": f"Gap b/w {phase['name']} and {nxt['name']}",
                "type": "Gap",
                "status": "—",
                "start": gap_start.strftime("%Y-%m-%d"),
                "end": gap_end.strftime("%Y-%m-%d"),
            })
    return expanded


PROGRAM_TIMELINE_EXPANDED = _with_gaps(PROGRAM_TIMELINE)

# Column name -> canonical key. Matching is done by normalized prefix so the app
# survives minor header drift (trailing notes, whitespace, case).
COLUMN_MAP = {
    "RICE ID": "rice_id",
    "RICE Type": "rice_type",
    "Pre Fix - Workstream": "prefix_workstream",
    "Sub Entity": "sub_entity",
    "Release": "release",
    "Object Name": "object_name",
    "Object Description": "description",
    "Module": "module",
    "Workstream": "workstream",
    "Object Status": "object_status",
    "Accountable Org": "accountable_org",
    "Operation System": "operation_system",
    "Scope Origin": "scope_origin",
    "In Scope": "in_scope",
    "Method": "method",
    "Interface Direction": "interface_direction",
    "Source System": "source_system",
    "Target System": "target_system",
    "Design Sprint": "design_sprint",
    "Dev - Sprint": "dev_sprint",
    "Tech Spec Owner": "tech_spec_owner",
    "RICE Owner": "rice_owner",
    "RICE Status": "rice_status",
    "Complexity": "complexity",
    "Spec Completion Date - Planned": "spec_planned",
    "Spec Completion Date - Revised": "spec_revised",
    "Spec Start Date - Actual": "spec_start_actual",
    "Spec Completion Date - Actual": "spec_actual",
    "Spec - Approval Date": "spec_approval",
    "Functional Owner": "functional_owner",
    "Spec % Complete": "spec_pct_raw",
    "Functional Spec Status": "fspec_status",
    "Technical Owner": "technical_owner",
    "Dev Start Date - Planned": "dev_start_planned",
    "Dev Start Date - Actual": "dev_start_actual",
    # Build/dev-complete milestone. Older workbooks label it "Build + UT
    # Completion Date"; newer ones renamed it to "Dev + UT Completion Date"
    # (with inconsistent spacing between the two columns). These drive the
    # delivery date and sprint assignment — a mismatch dumps every object into
    # "Unscheduled" — so we map both spellings.
    "Build + UT Completion Date - P": "build_planned",
    "Build + UT Completion Date - A": "build_actual",
    "Dev+ UT Completion Date - Planned": "build_planned",
    "Dev + UT Completion Date - Actual": "build_actual",
    "Dev %": "dev_pct_raw",
    "Dev Status": "dev_status",
    "FUT Start Date - Actual": "fut_start_actual",
    "FUT Completion Date - Planned": "fut_planned",
    "FUT Completion Date - Actual": "fut_actual",
    "FUT Status": "fut_status",
    "FUT Required?": "fut_required",
    "Frequency": "frequency",
    "Comments": "comments",
    "Build hours": "build_hours",
    "Build Status": "build_status",
}


def _norm(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower()


# --------------------------------------------------------------------------- #
# Safe parsers
# --------------------------------------------------------------------------- #
def parse_date(value):
    """Return ISO date string or None. Never raises."""
    if value is None:
        return None
    try:
        ts = pd.to_datetime(value, errors="coerce")
    except Exception:
        return None
    if ts is None or pd.isna(ts):
        return None
    try:
        return ts.strftime("%Y-%m-%d")
    except Exception:
        return None


def parse_pct(value):
    """Extract a 0-100 percentage from values like '33% - In Progress'. None if absent."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        v = float(value)
        if v <= 1.0:
            v *= 100.0
        return max(0.0, min(100.0, v))
    s = str(value)
    m = re.search(r"(-?\d+(?:\.\d+)?)\s*%", s)
    if not m:
        m = re.search(r"(-?\d+(?:\.\d+)?)", s)
    if not m:
        return None
    try:
        v = float(m.group(1))
    except ValueError:
        return None
    if v <= 1.0 and "%" not in s:
        v *= 100.0
    return max(0.0, min(100.0, v))


def parse_num(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    try:
        v = float(value)
        if math.isnan(v):
            return None
        return v
    except (ValueError, TypeError):
        m = re.search(r"-?\d+(?:\.\d+)?", str(value))
        return float(m.group(0)) if m else None


def clean_str(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    s = str(value).replace("\r", " ").strip()
    if s.lower() in ("nan", "nat", "none"):
        return ""
    return s


# --------------------------------------------------------------------------- #
# Workbook loading
# --------------------------------------------------------------------------- #
def _pick_sheet(xls):
    for name in SHEET_CANDIDATES:
        if name in xls.sheet_names:
            return name
    # fallback: sheet containing a "RICE ID" header
    for name in xls.sheet_names:
        try:
            head = pd.read_excel(xls, sheet_name=name, nrows=1)
        except Exception:
            continue
        if any(_norm(c).startswith("rice id") for c in head.columns):
            return name
    return xls.sheet_names[0]


def _resolve_columns(df):
    """Build {canonical_key: actual_column} using normalized prefix matching."""
    resolved = {}
    norm_targets = {_norm(k): v for k, v in COLUMN_MAP.items()}
    for col in df.columns:
        nc = _norm(col)
        if nc in norm_targets:
            resolved.setdefault(norm_targets[nc], col)
            continue
        for nk, key in norm_targets.items():
            if nc.startswith(nk[:18]) and key not in resolved:
                resolved[key] = col
                break
    return resolved


def load_dataframe(path):
    """Load workbook, return (df, sheet_name). Raises informative ValueError."""
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    xls = pd.ExcelFile(path)
    sheet = _pick_sheet(xls)
    df = pd.read_excel(xls, sheet_name=sheet, dtype=object)
    df = df.dropna(how="all")
    return df, sheet


def process(path):
    """Full ingestion pipeline -> dict ready to serialize as JSON."""
    df, sheet = load_dataframe(path)
    cols = _resolve_columns(df)

    if "rice_id" not in cols:
        raise ValueError("Could not locate a 'RICE ID' column in the workbook.")

    df = df[df[cols["rice_id"]].notna()].copy()

    def g(row, key):
        c = cols.get(key)
        return row[c] if c is not None and c in row else None

    today = pd.Timestamp(dt.date.today())
    timeline = [
        {**t, "start_ts": pd.Timestamp(t["start"]), "end_ts": pd.Timestamp(t["end"])}
        for t in PROGRAM_TIMELINE_EXPANDED
    ]

    def find_phase(ts):
        # The expanded timeline fills every gap between phases with its own
        # "Gap b/w X and Y" entry, so it tiles the calendar with no holes
        # (aside from before the very first phase). Matching the earliest
        # entry whose end hasn't passed therefore lands a date on its real
        # phase, its gap tile, or (before day one) the first phase.
        if ts is None or pd.isna(ts):
            return None
        for t in timeline:
            if ts <= t["end_ts"]:
                return t["name"]
        return None

    records = []
    for _, row in df.iterrows():
        rec = {}
        # raw text fields
        for key in ("rice_id", "rice_type", "object_name", "description", "module",
                    "workstream", "accountable_org", "in_scope", "source_system",
                    "target_system", "design_sprint", "dev_sprint", "complexity",
                    "object_status", "functional_owner", "technical_owner",
                    "tech_spec_owner", "rice_owner", "rice_status", "fspec_status",
                    "dev_status", "fut_status", "build_status", "fut_required",
                    "method", "interface_direction", "frequency", "sub_entity",
                    "release", "comments"):
            rec[key] = clean_str(g(row, key))

        if not rec["rice_type"]:
            rec["rice_type"] = "Unspecified"

        # dates
        for key in ("spec_planned", "spec_revised", "spec_start_actual", "spec_actual",
                    "spec_approval", "dev_start_planned", "dev_start_actual",
                    "build_planned", "build_actual", "fut_start_actual",
                    "fut_planned", "fut_actual"):
            rec[key] = parse_date(g(row, key))

        # numbers / percents
        rec["build_hours"] = parse_num(g(row, "build_hours"))
        rec["dev_pct"] = parse_pct(g(row, "dev_pct_raw"))
        rec["spec_pct"] = parse_pct(g(row, "spec_pct_raw"))
        rec["dev_pct_raw"] = clean_str(g(row, "dev_pct_raw"))
        rec["spec_pct_raw"] = clean_str(g(row, "spec_pct_raw"))

        # hours calc
        bh = rec["build_hours"]
        dp = rec["dev_pct"]
        if bh is not None:
            pct = (dp if dp is not None else 0.0) / 100.0
            rec["hours_consumed"] = round(bh * pct, 1)
            rec["hours_left"] = round(bh - rec["hours_consumed"], 1)
            if dp is not None and dp >= 100:
                rec["hours_left"] = 0.0
        else:
            rec["hours_consumed"] = None
            rec["hours_left"] = None

        # ---- Gantt logic ----
        spec_eff = rec["spec_revised"] or rec["spec_planned"]
        rec["spec_effective"] = spec_eff
        # diamond: spec complete marker (actual if available else effective)
        rec["gantt_spec"] = rec["spec_actual"] or spec_eff

        # objects with no dev work ("Not Applicable") get no build bar; the
        # spec diamond above still renders.
        dev_not_applicable = (rec["dev_status"] or "").strip().lower() == "not applicable"

        # build bar spans Dev Start Date -> Dev + UT Completion Date.
        # start: actual dev-start else planned dev-start (spec dates are shown
        # separately as the diamond marker and must not drive the bar start).
        build_start = rec["dev_start_actual"] or rec["dev_start_planned"]
        gantt_start = pd.Timestamp(build_start) if build_start and not dev_not_applicable else None

        # delivery (Dev + UT Completion): actual else planned
        delivery = rec["build_actual"] or rec["build_planned"]
        rec["delivery_date"] = delivery
        gantt_delivery = pd.Timestamp(delivery) if delivery and not dev_not_applicable else None

        # fallback duration from build hours when delivery missing
        if gantt_start is not None and gantt_delivery is None and bh:
            weeks = max(bh / HOURS_PER_DEV_WEEK, 0.2)
            gantt_delivery = gantt_start + pd.Timedelta(weeks=weeks)
            rec["gantt_delivery_estimated"] = True
        else:
            rec["gantt_delivery_estimated"] = False

        # if we have delivery but no start, give a short lead-in bar
        if gantt_start is None and gantt_delivery is not None and bh:
            weeks = max(bh / HOURS_PER_DEV_WEEK, 0.2)
            gantt_start = gantt_delivery - pd.Timedelta(weeks=weeks)

        # guard against reversed bars from inconsistent source dates
        if gantt_start is not None and gantt_delivery is not None and gantt_start > gantt_delivery:
            gantt_start = gantt_delivery - pd.Timedelta(weeks=1)

        rec["gantt_start"] = gantt_start.strftime("%Y-%m-%d") if gantt_start is not None else None
        rec["gantt_delivery"] = gantt_delivery.strftime("%Y-%m-%d") if gantt_delivery is not None else None

        # ---- Sprint assignment (Actual -> Planned -> Unscheduled) ----
        deliver_ts = None
        if rec["build_actual"]:
            deliver_ts = pd.Timestamp(rec["build_actual"])
        elif rec["build_planned"]:
            deliver_ts = pd.Timestamp(rec["build_planned"])
        phase = find_phase(deliver_ts)
        if phase:
            rec["assigned_sprint"] = phase
        elif dev_not_applicable:
            rec["assigned_sprint"] = ""
        else:
            rec["assigned_sprint"] = "Unscheduled"

        # ---- Risk flags ----
        spec_eff_ts = pd.Timestamp(spec_eff) if spec_eff else None
        obj_stat = (rec["object_status"] or "").lower()
        completed = "complete" in obj_stat or "done" in obj_stat
        lean = False
        if rec["fspec_status"] and "delay" in rec["fspec_status"].lower():
            lean = True
        if "delay" in obj_stat:
            lean = True
        if spec_eff_ts is not None and not completed and not rec["spec_actual"]:
            if spec_eff_ts < today:
                lean = True
            elif spec_eff_ts <= today + pd.Timedelta(days=14):
                lean = True
        rec["lean_spec_risk"] = lean

        build = False
        ds = (rec["dev_status"] or "").lower()
        if "delay" in ds or "block" in ds:
            build = True
        if gantt_delivery is not None and (dp is None or dp < 100):
            if gantt_delivery <= today + pd.Timedelta(days=14):
                build = True
            if gantt_delivery < today:
                build = True
        rec["build_risk"] = build

        records.append(rec)

    payload = {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "source_sheet": sheet,
        "record_count": len(records),
        "timeline": PROGRAM_TIMELINE_EXPANDED,
        "hours_per_dev_week": HOURS_PER_DEV_WEEK,
        "records": records,
        "filters": _filter_options(records),
        "data_quality": _data_quality(records),
        "summary": _summary(records),
    }
    return payload


def _uniq_sorted(values):
    seen = sorted({v for v in values if v})
    return seen


def _filter_options(records):
    return {
        "accountable_org": _uniq_sorted(r["accountable_org"] for r in records),
        "module": _uniq_sorted(r["module"] for r in records),
        "release": _uniq_sorted(r["release"] for r in records),
        "sub_entity": _uniq_sorted(r["sub_entity"] for r in records),
        "in_scope": _uniq_sorted(r["in_scope"] for r in records),
        "rice_type": _uniq_sorted(r["rice_type"] for r in records),
        "object_status": _uniq_sorted(r["object_status"] for r in records),
        "design_sprint": _uniq_sorted(r["design_sprint"] for r in records),
        "dev_sprint": _uniq_sorted(r["dev_sprint"] for r in records),
        "workstream": _uniq_sorted(r["workstream"] for r in records),
        "functional_owner": _uniq_sorted(r["functional_owner"] for r in records),
        "technical_owner": _uniq_sorted(r["technical_owner"] for r in records),
        "complexity": _uniq_sorted(r["complexity"] for r in records),
        "assigned_sprint": _uniq_sorted(r["assigned_sprint"] for r in records),
    }


def _data_quality(records):
    total = len(records) or 1
    miss = {
        "Sprint": 0, "Spec Date": 0, "Build Hours": 0, "Delivery Date": 0,
        "Functional Owner": 0, "Object Status": 0, "RICE Type": 0, "Module": 0,
    }
    for r in records:
        if not r["design_sprint"] and not r["dev_sprint"]:
            miss["Sprint"] += 1
        if not (r["spec_planned"] or r["spec_revised"] or r["spec_actual"]):
            miss["Spec Date"] += 1
        if r["build_hours"] is None:
            miss["Build Hours"] += 1
        if not (r["build_planned"] or r["build_actual"]):
            miss["Delivery Date"] += 1
        if not r["functional_owner"]:
            miss["Functional Owner"] += 1
        if not r["object_status"]:
            miss["Object Status"] += 1
        if not r["rice_type"] or r["rice_type"] == "Unspecified":
            miss["RICE Type"] += 1
        if not r["module"]:
            miss["Module"] += 1
    return [
        {"field": k, "missing": v, "pct": round(100.0 * v / total, 1)}
        for k, v in miss.items()
    ]


def _summary(records):
    in_scope = [r for r in records if r["in_scope"] == "Yes"]
    return {
        "total_all": len(records),
        "total_in_scope": len(in_scope),
        "total_build_hours": round(sum(r["build_hours"] or 0 for r in records), 1),
        "in_scope_build_hours": round(sum(r["build_hours"] or 0 for r in in_scope), 1),
    }


# --------------------------------------------------------------------------- #
# Resource planning
# --------------------------------------------------------------------------- #
# Fixed planning scope: Deloitte-accountable, in-scope objects, excluding the
# ARCS module (the "ARCS (ADP)" module is a different module and stays in) and
# excluding Conversions. Blocked/delayed objects are schedulable like any other.
SIT1_START = pd.Timestamp("2026-09-28")
SIT2_START = pd.Timestamp("2026-11-09")
MAX_PLAN_RESOURCES = 40

# Queue priority: complexity dominates (start complex work first to absorb
# delays), integrations get a strong boost, hours break ties.
TYPE_PRIORITY = {"integration": 30, "extension": 15, "report": 0}
COMPLEXITY_PRIORITY = {"very complex": 100, "complex": 80, "medium": 50, "simple": 20}


def _plan_scope(records):
    out = []
    for r in records:
        if (r["accountable_org"] or "").strip().lower() != "deloitte":
            continue
        if (r["in_scope"] or "").strip().lower() != "yes":
            continue
        if (r["module"] or "").strip().upper() == "ARCS":
            continue
        if (r["rice_type"] or "").strip().lower() == "conversion":
            continue
        out.append(r)
    return out


def _remaining_hours(r):
    """Remaining build effort. hours_left already credits Dev % progress."""
    if r["hours_left"] is not None:
        return r["hours_left"]
    return r["build_hours"]


def _plan_items(scope):
    """Split scope into schedulable items / done / not-applicable. Objects with
    no build-hours estimate get the median hours of their complexity tier."""
    # medians for defaulting, computed from the scoped population
    by_cx = {}
    all_hours = []
    for r in scope:
        if r["build_hours"]:
            by_cx.setdefault((r["complexity"] or "").strip().lower(), []).append(r["build_hours"])
            all_hours.append(r["build_hours"])
    med = {k: float(np.median(v)) for k, v in by_cx.items()}
    overall_med = float(np.median(all_hours)) if all_hours else HOURS_PER_DEV_WEEK * 2

    items, done, not_applicable = [], [], []
    for r in scope:
        if (r["dev_status"] or "").strip().lower() == "not applicable":
            not_applicable.append(r)
            continue
        if (r["build_status"] or "").strip().lower() == "completed":
            done.append(r)
            continue
        rem = _remaining_hours(r)
        estimated = False
        if not rem or rem <= 0:
            if r["build_hours"] is None:
                cx = (r["complexity"] or "").strip().lower()
                rem = med.get(cx, overall_med)
                estimated = True
            else:
                done.append(r)  # effort fully consumed
                continue
        cx = (r["complexity"] or "").strip().lower()
        priority = (COMPLEXITY_PRIORITY.get(cx, 50)
                    + TYPE_PRIORITY.get((r["rice_type"] or "").strip().lower(), 0))
        items.append({
            "rice_id": r["rice_id"],
            "object_name": r["object_name"],
            "rice_type": r["rice_type"],
            "module": r["module"],
            "complexity": r["complexity"] or "Unrated",
            "hours": float(rem),
            "estimated": estimated,
            "priority": priority,
        })
    return items, done, not_applicable


def _simulate(items, roster, cal, hpw, ramp_pct):
    """Greedy list scheduler. items are priority-ordered; each resource pulls
    the next queue item when free. roster is [(onboard_day, ramp_days)] —
    existing team members carry ramp_days=0. Day indices are business days."""
    daily = hpw / 5.0
    pending = list(items)
    res = [{"onboard": ob, "ramp": rd, "cur": None, "assignments": []}
           for ob, rd in sorted(roster)]
    finish = {}
    weeks = len(cal) // 5 + 1
    wk_capacity = [0.0] * weeks
    wk_scheduled = [0.0] * weeks
    last_day = 0
    for d in range(len(cal)):
        if not pending and all(rs["cur"] is None for rs in res):
            break
        for rs in res:
            if d < rs["onboard"]:
                continue
            if rs["cur"] is None and not pending:
                continue
            rate = daily * (ramp_pct if d - rs["onboard"] < rs["ramp"] else 1.0)
            wk_capacity[d // 5] += rate
            avail = rate
            while avail > 1e-9:
                if rs["cur"] is None:
                    if not pending:
                        break
                    it = pending.pop(0)
                    rs["cur"] = {"item": it, "rem": it["hours"], "start": d}
                take = min(avail, rs["cur"]["rem"])
                rs["cur"]["rem"] -= take
                avail -= take
                if rs["cur"]["rem"] <= 1e-9:
                    a = rs["cur"]
                    rs["assignments"].append({"item": a["item"], "start": a["start"], "end": d})
                    finish[a["item"]["rice_id"]] = d
                    last_day = max(last_day, d)
                    rs["cur"] = None
            wk_scheduled[d // 5] += rate - max(avail, 0.0)
    unfinished = [rs["cur"]["item"]["rice_id"] for rs in res if rs["cur"]] + \
                 [it["rice_id"] for it in pending]
    return {"resources": res, "finish": finish, "unfinished": unfinished,
            "last_day": last_day, "wk_capacity": wk_capacity, "wk_scheduled": wk_scheduled}


def _score(sim, items, deadlines):
    """(deadline misses, total business days late, overall finish) — lower is better."""
    miss, late = 0, 0
    for it in items:
        fin = sim["finish"].get(it["rice_id"])
        dl = deadlines[it["rice_id"]]
        if fin is None:
            miss, late = miss + 1, late + 120
        elif fin > dl:
            miss, late = miss + 1, late + (fin - dl)
    return (miss, late, sim["last_day"])


def _run_scenario(items, deadlines, cal, cfg, hpw, ramp_days, ramp_pct, current_team):
    """Hire-when-needed planning: onboard at most cap resources per week,
    starting as early as the math demands, stopping once every deadline holds
    — or once further hires stop reducing lateness (deadlines unreachable)."""
    cap = cfg["max_onboards_per_week"]
    base = [(0, 0)] * current_team          # existing team: no ramp-up
    hires, week, stale = [], 0, 0
    sim = _simulate(items, base, cal, hpw, ramp_pct)
    score = _score(sim, items, deadlines)
    best = (list(hires), sim, score)
    while score[0] > 0 and len(hires) < MAX_PLAN_RESOURCES and week < 52 and stale < 4:
        if sum(1 for h in hires if h[0] // 5 == week) < cap:
            hires.append((week * 5, ramp_days))
            sim = _simulate(items, base + hires, cal, hpw, ramp_pct)
            score = _score(sim, items, deadlines)
            if score < best[2]:
                best, stale = (list(hires), sim, score), 0
            else:
                stale += 1
        else:
            week += 1
    return best[0], best[1]


def _iso(cal, d):
    d = min(max(d, 0), len(cal) - 1)
    return cal[d].strftime("%Y-%m-%d")


def _type_key(t):
    """Canonical RICE-type label used to silo resource pools."""
    return (t or "").strip() or "Unspecified"


def _split_team(total, weights, types):
    """Apportion `total` existing devs across `types` by `weights` (largest
    remainder). Falls back to an equal split when no weight is available."""
    types = list(types)
    if total <= 0 or not types:
        return {t: 0 for t in types}
    wsum = sum(max(weights.get(t, 0), 0) for t in types)
    if wsum <= 0:
        base, extra = divmod(total, len(types))
        return {t: base + (1 if i < extra else 0) for i, t in enumerate(types)}
    quota = {t: total * max(weights.get(t, 0), 0) / wsum for t in types}
    out = {t: int(quota[t]) for t in types}
    rem = total - sum(out.values())
    for t in sorted(types, key=lambda t: quota[t] - int(quota[t]), reverse=True)[:rem]:
        out[t] += 1
    return out


def _scenario_payload(label, description, items, deadlines, cal, cfg, hpw, ramp_days, ramp_pct, team_by_type):
    """Resources are siloed by RICE type: each type is scheduled against its
    own dev pool (existing team of that type + its own hire track) and never
    shares developers across types. Results are merged for reporting."""
    daily = hpw / 5.0

    groups = {}
    for it in items:
        groups.setdefault(_type_key(it["rice_type"]), []).append(it)

    resources = []
    finish = {}
    unfinished = []
    hires_total = 0
    current_team_total = 0
    onboards_by_week = {}
    spans_by_type = {}                       # type -> [(onboard_day, last_end_day)]
    wk_capacity, wk_scheduled = {}, {}       # week index -> hours
    dev_seq = 0

    for ty in sorted(groups):
        its = groups[ty]
        ct = int(team_by_type.get(ty, 0))
        current_team_total += ct
        hires, sim = _run_scenario(its, deadlines, cal, cfg, hpw, ramp_days, ramp_pct, ct)
        hires_total += len(hires)
        finish.update(sim["finish"])
        unfinished += sim["unfinished"]
        for h in hires:
            w = h[0] // 5
            onboards_by_week[w] = onboards_by_week.get(w, 0) + 1
        for w in range(len(sim["wk_capacity"])):
            wk_capacity[w] = wk_capacity.get(w, 0.0) + sim["wk_capacity"][w]
            wk_scheduled[w] = wk_scheduled.get(w, 0.0) + sim["wk_scheduled"][w]

        spans = spans_by_type.setdefault(ty, [])
        for idx, rs in enumerate(sim["resources"]):
            if not rs["assignments"] and rs["cur"] is None:
                continue
            existing = idx < ct               # base team is appended first in _run_scenario
            dev_seq += 1
            ends = [a["end"] for a in rs["assignments"]]
            rolloff = max(ends) if ends else rs["onboard"]
            cap_hours = sum(daily * (ramp_pct if d - rs["onboard"] < rs["ramp"] else 1.0)
                            for d in range(rs["onboard"], rolloff + 1))
            done_hours = sum(a["item"]["hours"] for a in rs["assignments"])
            resources.append({
                "name": f"{ty[:3].upper()} Dev {dev_seq}" + (" (current)" if existing else ""),
                "rice_type": ty,
                "existing": existing,
                "onboard": _iso(cal, rs["onboard"]),
                "rolloff": _iso(cal, rolloff),
                "hours": round(done_hours, 1),
                "utilization": round(100.0 * done_hours / cap_hours, 1) if cap_hours else 0.0,
                "assignments": [{
                    "rice_id": a["item"]["rice_id"],
                    "object_name": a["item"]["object_name"],
                    "rice_type": a["item"]["rice_type"],
                    "complexity": a["item"]["complexity"],
                    "hours": round(a["item"]["hours"], 1),
                    "estimated": a["item"]["estimated"],
                    "start": _iso(cal, a["start"]),
                    "end": _iso(cal, a["end"]),
                    "deadline": _iso(cal, deadlines[a["item"]["rice_id"]]),
                    "late": a["end"] > deadlines[a["item"]["rice_id"]],
                } for a in rs["assignments"]],
            })
            if ends:
                spans.append((rs["onboard"], max(ends)))

    misses = []
    for it in items:
        fin = finish.get(it["rice_id"])
        dl = deadlines[it["rice_id"]]
        if fin is None or fin > dl:
            misses.append({
                "rice_id": it["rice_id"], "object_name": it["object_name"],
                "rice_type": it["rice_type"], "hours": round(it["hours"], 1),
                "finish": _iso(cal, fin) if fin is not None else None,
                "deadline": _iso(cal, dl),
            })

    # weekly ramp series with per-type active-developer breakdown (for the
    # stacked chart). A resource is active from its onboard week to roll-off.
    all_ends = [e for spans in spans_by_type.values() for _, e in spans]
    n_weeks = (max(all_ends, default=0) // 5) + 1
    types_present = sorted(spans_by_type)
    weekly = []
    peak_week = None
    for w in range(n_weeks):
        d0, d4 = w * 5, w * 5 + 4
        by_type = {}
        for ty in types_present:
            c = sum(1 for s, e in spans_by_type[ty] if s <= d4 and e >= d0)
            if c:
                by_type[ty] = c
        active = sum(by_type.values())
        row = {
            "week": _iso(cal, w * 5),
            "active": active,
            "by_type": by_type,
            "onboards": onboards_by_week.get(w, 0),
            "capacity": round(wk_capacity.get(w, 0.0), 1),
            "scheduled": round(wk_scheduled.get(w, 0.0), 1),
        }
        weekly.append(row)
        if peak_week is None or active > peak_week["active"]:
            peak_week = row

    finish_day = max(finish.values()) if finish else 0
    utils = [r["utilization"] for r in resources]
    return {
        "label": label,
        "description": description,
        "max_onboards_per_week": cfg["max_onboards_per_week"],
        "contingency_pct": cfg.get("contingency_pct", 0),
        "types": types_present,
        "kpis": {
            "peak_team": peak_week["active"] if peak_week else 0,
            "peak_by_type": peak_week["by_type"] if peak_week else {},
            "onboarded": hires_total,
            "existing_team": current_team_total,
            "existing_by_type": {ty: int(team_by_type.get(ty, 0)) for ty in sorted(groups)},
            "finish": _iso(cal, finish_day) if not unfinished else None,
            "buffer_bdays": int(np.busday_count(cal[min(finish_day, len(cal) - 1)].date(),
                                                SIT1_START.date())) if not unfinished else None,
            "hours": round(sum(it["hours"] for it in items), 1),
            "objects": len(items),
            "feasible": len(misses) == 0,
            "misses": len(misses),
            "late_hours": round(sum(m["hours"] for m in misses), 1),
            "avg_utilization": round(float(np.mean(utils)), 1) if utils else 0.0,
        },
        "weekly": weekly,
        "resources": resources,
        "misses": misses,
    }


def build_resource_plan(records, params):
    scope = _plan_scope(records)
    items, done, not_applicable = _plan_items(scope)
    items.sort(key=lambda it: (-it["priority"], -it["hours"]))

    hpw = params["hours_per_week"]
    ramp_days = int(params["ramp_weeks"] * 5)
    ramp_pct = params["ramp_pct"] / 100.0
    contingency = params["contingency_pct"] / 100.0
    buffer_days = int(params["buffer_days"])

    today = pd.Timestamp(dt.date.today())
    start = today if np.is_busday(today.date()) else pd.Timestamp(np.busday_offset(today.date(), 0, roll="forward"))
    cal = pd.bdate_range(start, periods=560)
    sit1_dl = max(int(cal.searchsorted(SIT1_START)) - 1, 0)   # last bday before SIT 1
    sit2_dl = max(int(cal.searchsorted(SIT2_START)) - 1, 0)

    def all_by(day):
        return {it["rice_id"]: day for it in items}

    def inflated(factor):
        return [{**it, "hours": it["hours"] * (1.0 + factor)} for it in items]

    # Two-wave: integrations, extensions and complex work must land before
    # SIT 1; simple/medium reports may land in the SIT1→SIT2 window.
    def wave_of(it):
        if (it["rice_type"] or "").lower() in ("integration", "extension"):
            return 1
        if (it["complexity"] or "").lower() in ("complex", "very complex"):
            return 1
        return 2

    wave_items = sorted(items, key=lambda it: (wave_of(it), -it["priority"], -it["hours"]))
    wave_deadlines = {it["rice_id"]: (sit1_dl if wave_of(it) == 1 else sit2_dl) for it in items}

    # Existing team, siloed by RICE type. Default (auto): one dev per open
    # object with active dev work, split across types by the in-flight work mix
    # (objects currently In Progress / Delayed). An explicit team size is
    # apportioned across the types that still have open work by the same mix.
    wip_by_type = {}
    for r in scope:
        if (r["build_status"] or "").strip().lower() == "completed":
            continue
        if (r["dev_status"] or "").strip().lower() in ("in progress", "delayed"):
            k = _type_key(r["rice_type"])
            wip_by_type[k] = wip_by_type.get(k, 0) + 1
    item_types = sorted({_type_key(it["rice_type"]) for it in items})
    if params["current_team"] >= 0:
        team_by_type = _split_team(int(params["current_team"]), wip_by_type, item_types)
    else:
        # auto: staff each type to its in-flight count; guarantee at least one
        # dev overall so a fresh plan still has a starting pool.
        team_by_type = {t: wip_by_type.get(t, 0) for t in item_types}
        if item_types and sum(team_by_type.values()) == 0:
            team_by_type[item_types[0]] = 1
    current_team = sum(team_by_type.values())
    params = {**params, "current_team": current_team}

    scenarios = {
        "aggressive": _scenario_payload(
            "Aggressive", f"Everything done {buffer_days} working days before SIT 1 · up to 2 onboards/week per type",
            items, all_by(max(sit1_dl - buffer_days, 0)), cal,
            {"max_onboards_per_week": 2}, hpw, ramp_days, ramp_pct, team_by_type),
        "optimized": _scenario_payload(
            "Optimized", "Everything done just-in-time for SIT 1 · max 1 onboard/week per type · smallest team",
            items, all_by(sit1_dl), cal,
            {"max_onboards_per_week": 1}, hpw, ramp_days, ramp_pct, team_by_type),
        "conservative": _scenario_payload(
            "Conservative", f"SIT 1 deadline with +{params['contingency_pct']:.0f}% effort contingency baked into every estimate",
            inflated(contingency), all_by(sit1_dl), cal,
            {"max_onboards_per_week": 1, "contingency_pct": params["contingency_pct"]}, hpw, ramp_days, ramp_pct, team_by_type),
        "two_wave": _scenario_payload(
            "Two-Wave", "Integrations, extensions & complex work before SIT 1 · simple/medium reports may land before SIT 2",
            wave_items, wave_deadlines, cal,
            {"max_onboards_per_week": 1}, hpw, ramp_days, ramp_pct, team_by_type),
    }

    est_count = sum(1 for it in items if it["estimated"])
    return {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "params": params,
        "plan_start": start.strftime("%Y-%m-%d"),
        "sit1_start": SIT1_START.strftime("%Y-%m-%d"),
        "sit2_start": SIT2_START.strftime("%Y-%m-%d"),
        "scope": {
            "description": "Deloitte · In Scope · Module ≠ ARCS · no Conversions",
            "objects": len(scope),
            "open": len(items),
            "completed": len(done),
            "not_applicable": len(not_applicable),
            "planned_hours": round(sum(it["hours"] for it in items), 1),
            "estimated_objects": est_count,
        },
        "scenarios": scenarios,
    }


def _plan_params():
    def num(name, default, lo, hi):
        try:
            v = float(request.args.get(name, default))
        except (TypeError, ValueError):
            v = default
        return max(lo, min(hi, v))
    return {
        "hours_per_week": num("hours_per_week", HOURS_PER_DEV_WEEK, 10, 80),
        "ramp_weeks": num("ramp_weeks", 2, 0, 8),
        "ramp_pct": num("ramp_pct", 50, 10, 100),
        "contingency_pct": num("contingency_pct", 15, 0, 100),
        "buffer_days": num("buffer_days", 10, 0, 40),
        # -1 = auto: derive from objects with dev work currently in flight
        "current_team": num("current_team", -1, -1, MAX_PLAN_RESOURCES),
    }


# --------------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------------- #
@app.route("/")
def index():
    return render_template("index.html", has_data=os.path.exists(DATA_FILE))


@app.route("/api/data")
def api_data():
    if not os.path.exists(DATA_FILE):
        return jsonify({"error": "no_data", "message": "No data file found. Please upload a workbook."}), 404
    try:
        payload = process(DATA_FILE)
    except Exception as exc:  # never crash — report cleanly
        return jsonify({"error": "processing_error", "message": str(exc)}), 500
    return jsonify(payload)


@app.route("/api/resource-plan")
def api_resource_plan():
    if not os.path.exists(DATA_FILE):
        return jsonify({"error": "no_data", "message": "No data file found. Please upload a workbook."}), 404
    try:
        payload = process(DATA_FILE)
        plan = build_resource_plan(payload["records"], _plan_params())
    except Exception as exc:
        return jsonify({"error": "processing_error", "message": str(exc)}), 500
    return jsonify(plan)


@app.route("/api/upload", methods=["POST"])
def api_upload():
    if "file" not in request.files:
        return jsonify({"error": "no_file", "message": "No file part in request."}), 400
    f = request.files["file"]
    if not f or f.filename == "":
        return jsonify({"error": "no_file", "message": "No file selected."}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return jsonify({"error": "bad_type", "message": f"Unsupported file type '{ext}'. Use .xlsx/.xls/.xlsm."}), 400

    raw = f.read()
    # validate it parses before committing
    try:
        tmp = os.path.join(UPLOAD_DIR, "_validate" + ext)
        with open(tmp, "wb") as out:
            out.write(raw)
        process(tmp)
    except Exception as exc:
        return jsonify({"error": "parse_failed", "message": f"Could not read workbook: {exc}"}), 400
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)

    with open(DATA_FILE, "wb") as out:
        out.write(raw)
    return jsonify({"ok": True, "message": "File uploaded and validated."})


@app.route("/api/export")
def api_export():
    if not os.path.exists(DATA_FILE):
        return jsonify({"error": "no_data"}), 404
    payload = process(DATA_FILE)
    df = pd.DataFrame(payload["records"])
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    mem = io.BytesIO(buf.getvalue().encode("utf-8"))
    mem.seek(0)
    return send_file(mem, mimetype="text/csv", as_attachment=True,
                     download_name="rice_tracker_export.csv")


def _load_saved_filters():
    if not os.path.exists(SAVED_FILTERS_FILE):
        return {}
    try:
        with open(SAVED_FILTERS_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _store_saved_filters(views):
    tmp = SAVED_FILTERS_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(views, fh, indent=2)
    os.replace(tmp, SAVED_FILTERS_FILE)


@app.route("/api/saved-filters", methods=["GET"])
def api_saved_filters_get():
    return jsonify(_load_saved_filters())


@app.route("/api/saved-filters", methods=["POST"])
def api_saved_filters_save():
    body = request.get_json(silent=True) or {}
    name = (body.get("name") or "").strip()
    view = body.get("view")
    if not name:
        return jsonify({"error": "bad_request", "message": "A filter name is required."}), 400
    if not isinstance(view, dict):
        return jsonify({"error": "bad_request", "message": "A filter payload is required."}), 400
    views = _load_saved_filters()
    views[name] = view
    _store_saved_filters(views)
    return jsonify({"ok": True, "filters": views})


@app.route("/api/saved-filters/<name>", methods=["DELETE"])
def api_saved_filters_delete(name):
    views = _load_saved_filters()
    if name in views:
        del views[name]
        _store_saved_filters(views)
    return jsonify({"ok": True, "filters": views})


@app.route("/health")
def health():
    return jsonify({"ok": True, "has_data": os.path.exists(DATA_FILE)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f" * RICE Tracker dashboard starting on http://127.0.0.1:{port}")
    app.run(host="127.0.0.1", port=port, debug=True)
