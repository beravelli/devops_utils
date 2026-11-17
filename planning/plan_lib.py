"""Shared helpers for the FY planning workbook.

Two things live here because both the workbook generator (build_plan.py) and the
story-breakdown skill (.claude/skills/story-breakdown/stories.py) need them:

* read_model()          - recompute the whole capacity model in Python
* add_dashboard_chart() - (re)create the Dashboard bar chart

Why recompute rather than read the sheet's own numbers: openpyxl writes formulas
as strings with no cached values, so until Excel, Sheets or LibreOffice has opened
and saved the file, every formula cell reads back as None. Recomputing from the
input cells is deterministic and gives the same answer the formulas will.

Why rebuild the chart: openpyxl drops charts and images on a load/save round-trip.
Any script that edits this workbook in place must put the chart back.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

from openpyxl.chart import BarChart, Reference

FONT = "Arial"

# ---- sheet geometry -------------------------------------------------------
SHEETS = dict(asm="Assumptions", hol="Holidays", cal="Sprint Calendar",
              cap="Team Capacity", sp="Sprint Plan", ap="Annual Plan", db="Dashboard")
DATA_R0 = 4            # first data row on Sprint Calendar / Sprint Plan / Annual Plan
AP_R1 = 23             # last Annual Plan row (16 examples + 4 blank template rows)
CAP_R0, CAP_R1 = 4, 12  # Team Capacity roster rows
HOL_R0, HOL_R1 = 5, 25  # Holidays rows


def _d(v):
    """Coerce a cell value to a date."""
    if isinstance(v, dt.datetime):
        return v.date()
    return v


def read_model(path: str | Path) -> dict:
    """Recompute the capacity model and the roadmap from the workbook's input cells."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    asm, hol, sp, ap, cap = (wb[SHEETS[k]] for k in ("asm", "hol", "sp", "ap", "cap"))

    weeks = int(asm["B6"].value)
    ns = int(asm["B7"].value)
    start = _d(asm["B8"].value)
    weekdays = weeks * 5                      # B9 is a formula: sprint length x 5
    focus = float(asm["B10"].value)
    thru = float(asm["B11"].value)
    resv = float(asm["B12"].value)
    oncall = float(asm["B13"].value)
    conf_default = float(asm["B14"].value)
    step, endoff = 7 * weeks, 7 * weeks - 3

    efte = 0.0
    roster = []
    for r in range(CAP_R0, CAP_R1 + 1):
        name, fte = cap.cell(row=r, column=1).value, cap.cell(row=r, column=3).value
        if not name or fte is None:
            continue
        alloc = cap.cell(row=r, column=4).value or 0
        ramp = cap.cell(row=r, column=5).value or 0
        eff = float(fte) * float(alloc) * float(ramp)
        efte += eff
        roster.append(dict(name=name, role=cap.cell(row=r, column=2).value, effective_fte=round(eff, 4),
                           pto_days=cap.cell(row=r, column=6).value))

    holidays = [_d(hol.cell(row=r, column=1).value) for r in range(HOL_R0, HOL_R1 + 1)
                if hol.cell(row=r, column=1).value]

    # ---- epics ----
    epics = []
    for r in range(DATA_R0, AP_R1 + 1):
        eid, est = ap.cell(row=r, column=1).value, ap.cell(row=r, column=6).value
        if not eid or est in (None, ""):
            continue
        conf = ap.cell(row=r, column=7).value or conf_default
        j, k = ap.cell(row=r, column=10).value, ap.cell(row=r, column=11).value
        spans = (int(k) - int(j) + 1) if (j and k) else 0
        ra = float(est) / float(conf) if conf else float(est)
        epics.append(dict(
            row=r, id=eid, name=ap.cell(row=r, column=2).value,
            workstream=ap.cell(row=r, column=3).value, priority=ap.cell(row=r, column=4).value,
            owner=ap.cell(row=r, column=5).value, estimate=float(est), confidence=float(conf),
            risk_adjusted=round(ra, 2), quarter=ap.cell(row=r, column=9).value,
            start_sprint=int(j) if j else None, end_sprint=int(k) if k else None,
            spans=spans, points_per_sprint=round(ra / spans, 2) if spans else 0.0,
            status=ap.cell(row=r, column=14).value))

    # ---- sprints ----
    sprints = []
    for i in range(ns):
        st = start + dt.timedelta(days=step * i)
        en = st + dt.timedelta(days=endoff)
        nh = sum(1 for h in holidays if st <= h <= en)
        wd = weekdays - nh
        pto = sp.cell(row=DATA_R0 + i, column=6).value or 0
        gross = efte * wd
        net = gross - float(pto) - oncall * wd
        base = net * focus * thru
        comm = base * (1 - resv)
        n = i + 1
        planned = sum(e["points_per_sprint"] for e in epics
                      if e["start_sprint"] and e["start_sprint"] <= n <= e["end_sprint"])
        sprints.append(dict(
            n=n, label=f"S{n:02d}", quarter=f"Q{(st.month - 1) // 3 + 1}",
            start=st.isoformat(), end=en.isoformat(), holidays=nh, working_days=wd,
            gross_dev_days=round(gross, 2), pto_days=float(pto),
            oncall_days=round(oncall * wd, 2), net_dev_days=round(net, 2),
            baseline_pts=round(base, 2), commitable_pts=round(comm, 2),
            planned_from_epics=round(planned, 2),
            load=round(planned / comm, 4) if comm else None))

    return dict(
        workbook=str(path), source="recomputed from input cells",
        cadence=dict(sprint_weeks=weeks, weekdays_per_sprint=weekdays, sprints=ns,
                     sprint1_start=start.isoformat(), last_sprint_end=sprints[-1]["end"] if sprints else None),
        levers=dict(focus_factor=focus, throughput_pts_per_focused_dev_day=thru,
                    unplanned_reserve=resv, oncall_people=oncall, default_confidence=conf_default),
        effective_fte=round(efte, 4), roster=roster, holidays=[h.isoformat() for h in holidays],
        totals=dict(baseline_pts=round(sum(s["baseline_pts"] for s in sprints), 2),
                    commitable_pts=round(sum(s["commitable_pts"] for s in sprints), 2),
                    demand_risk_adjusted_pts=round(sum(e["risk_adjusted"] for e in epics), 2),
                    demand_estimate_pts=round(sum(e["estimate"] for e in epics), 2)),
        epics=epics, sprints=sprints)


def add_dashboard_chart(ws, min_row=16, max_row=20, cat_row0=17):
    """(Re)create the quarterly commitable-vs-planned bar chart on the Dashboard."""
    ch = BarChart()
    ch.type = "col"
    ch.style = 2
    ch.title = "Commitable capacity vs planned demand, by quarter (story points)"
    ch.y_axis.title = "Story points"
    ch.x_axis.title = "Quarter"
    ch.roundedCorners = False
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.title.overlay = False
    ch.y_axis.title.overlay = False
    ch.x_axis.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.add_data(Reference(ws, min_col=6, max_col=7, min_row=min_row, max_row=max_row), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=cat_row0, max_row=max_row))
    ch.height = 8.5
    ch.width = 17
    ws.add_chart(ch, "K16")
    return ch
