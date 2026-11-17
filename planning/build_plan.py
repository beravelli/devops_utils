#!/usr/bin/env python3
"""Build a yearly (FY2027) sprint + story-point planning workbook, baseline example.

Cadence is parameterised: set WEEKS / NS and every sheet, range and formula follows.
"""
import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from plan_lib import add_dashboard_chart

OUT = str(Path(__file__).resolve().parent / "FY2027-sprint-plan-baseline.xlsx")
FONT = "Arial"

# ---------------- cadence ----------------
WEEKS = 3                      # sprint length in weeks
NS = 17                        # sprints in the plan (17 x 3 weeks = 357 days, 4 Jan -> 24 Dec 2027)
WD = 5 * WEEKS                 # weekdays per sprint, before holidays
STEP = 7 * WEEKS               # days between sprint starts
ENDOFF = STEP - 3              # Monday start -> Friday of the final week
R0 = 4                         # first data row on Sprint Calendar / Sprint Plan
R1 = R0 + NS - 1               # last data row
RT = R1 + 1                    # totals row
FY26_VEL = 57.0                # illustrative FY2026 actual velocity, per 3-week sprint
FY26_DAYS = 63.0               # illustrative FY2026 net dev-days, per 3-week sprint

# ---------------- style helpers ----------------
NAVY = "1F3864"; BAND = "DCE6F1"; YELLOW = "FFFF00"; GREY = "F2F2F2"
BLUE_TXT = "0000FF"; GREEN_TXT = "008000"

def F(sz=10, b=False, color="000000", it=False):
    return Font(name=FONT, size=sz, bold=b, color=color, italic=it)

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOPLINE = Border(top=Side(style="medium", color=NAVY))

NUM_PTS = '#,##0.0;(#,##0.0);-'
NUM_DAY = '#,##0.0;(#,##0.0);-'
NUM_INT = '#,##0;(#,##0);-'
NUM_PCT = '0.0%;(0.0%);-'
NUM_FTE = '0.00'
NUM_DATE = 'ddd dd-mmm-yyyy'

def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = F(14, True, NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = F(9, color="595959", it=True)

def header_row(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = F(10, True, "FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[row].height = 32

def section(ws, row, text):
    c = ws.cell(row=row, column=1, value=text)
    c.font = F(11, True, NAVY)
    c.border = TOPLINE

def inp(cell, key=False):
    cell.font = F(10, color=BLUE_TXT)
    if key:
        cell.fill = PatternFill("solid", fgColor=YELLOW)
    cell.border = BOX

def link(cell):
    cell.font = F(10, color=GREEN_TXT)
    cell.border = BOX

def calc(cell, bold=False):
    cell.font = F(10, b=bold)
    cell.border = BOX

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w

wb = Workbook()

# =====================================================================
# 1. READ ME
# =====================================================================
ws = wb.active
ws.title = "Read Me"
ws.sheet_properties.tabColor = NAVY
title(ws, "FY2027 Yearly Sprint & Story-Point Plan — Baseline Example",
      f"Example workbook on a {WEEKS}-week cadence ({NS} sprints). Every number here is illustrative: replace the "
      "inputs with your own and the whole plan recalculates.")
widths(ws, {"A": 34, "B": 86})

rows = [
    ("SECTION", "How this plan is built"),
    ("", "Capacity, not wishful thinking, sets the baseline. The chain is: team roster -> effective FTE -> dev-days per "
         "sprint -> minus holidays, PTO and on-call -> times a focus factor -> times a throughput rate (points per "
         "focused dev-day) = baseline velocity per sprint. A reserve for unplanned work is held back, and what is left "
         "is the commitable story points you plan roadmap work against."),
    ("", "Demand is the other half: epics on 'Annual Plan' carry a point estimate and a confidence, which gives a "
         "risk-adjusted size. Spread over the sprints an epic spans, that becomes the planned points per sprint. "
         "'Dashboard' compares demand against supply, by quarter and for the year."),
    ("SECTION", f"Cadence: {WEEKS}-week sprints"),
    ("", f"{NS} sprints of {WEEKS} weeks run Monday 4 January to Friday 24 December 2027 ({NS} x {STEP} = "
         f"{NS*STEP} days), leaving the last week of December outside the plan as shutdown and slack. Each sprint "
         f"carries {WD} weekdays before company holidays. Quarters get an uneven 5/4/4/4 split, because 21 days do "
         f"not divide a quarter."),
    ("", "A longer sprint does not create capacity — annual dev-days barely move. What changes is granularity: fewer, "
         "larger commitments, a coarser feedback loop, and less whipsaw from a single holiday or a week of leave, "
         "since each absence is a smaller share of a longer sprint."),
    ("", "To change cadence: set 'Sprint length' on 'Assumptions' (weekdays per sprint follows it automatically), set "
         "'Sprints in plan', then add or delete rows on 'Sprint Calendar' and 'Sprint Plan' so both hold that many "
         "rows, and re-point the start/end sprint numbers on 'Annual Plan'."),
    ("SECTION", "Tabs, in reading order"),
    ("Assumptions", "Every planning lever in one place: sprint cadence, focus factor, throughput baseline, reserve, "
                    "on-call cover, target investment mix. Start here."),
    ("Holidays", "Company holiday dates. The sprint calendar nets these out of working days automatically."),
    ("Sprint Calendar", f"{NS} {WEEKS}-week sprints with start/end dates, quarter, and net working days per sprint."),
    ("Team Capacity", "Roster: FTE, share of time allocated to this team, ramp factor, planned PTO. Produces effective FTE."),
    ("Sprint Plan", "The baseline, sprint by sprint: dev-days -> baseline points -> commitable points -> planned -> "
                    "load % -> actual completed -> variance -> rolling velocity. This is the sheet you live in."),
    ("Annual Plan", "The roadmap: epics with estimate, confidence, risk-adjusted size, target quarter, and the sprint "
                    "range they span."),
    ("Dashboard", "Supply vs demand for the year and per quarter, investment mix vs target, and the baseline "
                  "reconciliation against last year's actual velocity."),
    ("SECTION", "Colour legend"),
    ("Blue text on yellow", "Key input / baseline lever. These move the whole plan — change them deliberately."),
    ("Blue text", "Ordinary input. Type over it (PTO days, epic estimates, actual points completed, notes)."),
    ("Black text", "Formula. Do not overwrite; it will stop recalculating."),
    ("Green text", "Formula that pulls a value from another tab."),
    ("Grey fill", "Totals and derived roll-ups."),
    ("SECTION", "How to use it"),
    ("1. Set the calendar", "On 'Assumptions', set the planning year and the Sprint 1 start date (a Monday). Replace "
                            "the holiday dates on 'Holidays' with your company's."),
    ("2. Load the roster", "On 'Team Capacity', replace the example people with yours: FTE, allocation %, ramp factor "
                           "for anyone new, and planned PTO days for the year."),
    ("3. Calibrate the baseline", "On 'Assumptions', set the focus factor and the throughput baseline. Sanity-check "
                                  "both against last year on the 'Dashboard' baseline reconciliation block before you "
                                  "trust the plan. Throughput is per dev-day, so it is cadence-independent; velocity "
                                  "per sprint is not."),
    ("4. Spread PTO", f"On 'Sprint Plan', column F, put planned leave days in the sprints where they fall. "
                      f"'Team Capacity' rows {RT+2}-{RT+4} check the total against the roster."),
    ("5. Load the roadmap", "On 'Annual Plan', enter epics with estimates, confidence and a start/end sprint. Watch "
                            "Load % on 'Sprint Plan' and per quarter on 'Dashboard'; anything over 100% is overcommitted."),
    ("6. Track actuals", "Each sprint, put completed points in 'Sprint Plan' column N. Variance and rolling 3-sprint "
                         "velocity update, and after 3-4 sprints you re-baseline throughput from real data."),
    ("SECTION", "Example row — what a filled-in 'Annual Plan' row looks like"),
    ("", "ID: INI-05  |  Initiative: Checkout redesign (phase 2)  |  Workstream: Product / Roadmap  |  Priority: P0  |  "
         "Owner: B. Chen  |  Estimate: 48 pts  |  Confidence: 75%  |  Risk-adjusted: 64.0 pts (= 48 / 0.75)  |  "
         "Target quarter: Q2  |  Start sprint: 6  |  End sprint: 8  |  Spans 3 sprints  |  21.3 pts per sprint  |  "
         "Status: Committed"),
    ("SECTION", "Assumptions and hardcoded numbers in this example"),
    ("Dates", f"Sprint 1 starts Monday 4 January 2027; {NS} {WEEKS}-week sprints run to Friday 24 December 2027. "
              "Holidays are the 2027 US federal holidays plus two winter-shutdown days — an example, not your policy. "
              "New Year's Day and 31 December fall outside the sprint dates, so they cost the plan nothing."),
    ("Team", "7 example people, 5.90 effective FTE after allocation and ramp, 125 PTO days across the year. Illustrative."),
    ("Focus factor 70%", "Example value. Typical observed range is 60-75% of a working day on sprint backlog. "
                         "Measure your own rather than inheriting this one."),
    ("Throughput 1.25 pts/day", "Example value, reverse-engineered from the illustrative FY2026 actuals on 'Dashboard' "
                                f"({FY26_VEL:.1f} pts per {WEEKS}-week sprint at {FY26_DAYS:.1f} net dev-days and a 70% "
                                "focus factor = 1.29). Story points are team-relative, so this number is meaningless "
                                "across teams — derive yours from your own history."),
    ("Reserve 15%", "Example value for unplanned work. Teams with heavy escalation load run 20-30%."),
    ("FY2026 actuals", f"{FY26_VEL:.1f} pts per sprint and {FY26_DAYS:.1f} net dev-days per sprint, on 'Dashboard'. "
                       f"Placeholders stated on this same {WEEKS}-week cadence — if your history is on a different "
                       "sprint length, convert it per sprint before comparing, or compare dev-days instead."),
    ("Roadmap epics", "16 invented initiatives sized so annual demand lands at 87.3% of commitable capacity, which is "
                      "roughly where a year plan should start, with every quarter between 86% and 90%."),
    ("Over-loaded sprints", "Six sprints — S03, S07, S08, S11, S12, S16 — show Load % over 100% while every quarter "
                            "sits at 86-90%. That is not a mistake: epic points spread evenly across the sprints an "
                            "epic spans, but capacity does not, and each of those six loses 1-2 working days to "
                            "holidays. Move scope into S05, S09, S13, S14 and S17, which sit at 49-71% because the "
                            "next quarter's epics have not started yet. Levelling that is the actual exercise of "
                            "yearly planning."),
    ("Annual load 87.3%", "698.3 risk-adjusted points of demand against 799.9 commitable. On the earlier 2-week "
                          "version the same roadmap ran 92.1%, and three epic estimates (INI-06, INI-09, INI-10) were "
                          "trimmed when the cadence changed: 17 x 15 weekdays is 255 working days against 260 for 26 "
                          "x 10, and on-call cover scales with weekdays, so a 3-week year holds slightly less "
                          "commitable capacity than a 2-week one."),
]
r = 4
OVERLOAD_ROW = None
for label, text in rows:
    if label == "SECTION":
        section(ws, r, text)
        r += 1
        continue
    if label == "OVERLOAD":
        OVERLOAD_ROW = r
        r += 1
        continue
    c = ws.cell(row=r, column=1, value=label)
    c.font = F(10, True)
    c.alignment = Alignment(vertical="top", wrap_text=True)
    t = ws.cell(row=r, column=2, value=text)
    t.font = F(10)
    t.alignment = Alignment(vertical="top", wrap_text=True)
    ws.row_dimensions[r].height = max(14, 13 * (len(text) // 105 + 1))
    r += 1

# =====================================================================
# 2. ASSUMPTIONS
# =====================================================================
asm = wb.create_sheet("Assumptions")
asm.sheet_properties.tabColor = "2F5597"
title(asm, "Planning Assumptions — FY2027 Baseline",
      "Blue on yellow = key lever. Blue = input. Black = formula. Green = pulled from another tab.")
widths(asm, {"A": 42, "B": 14, "C": 24, "D": 78})
header_row(asm, 4, ["Assumption", "Value", "Unit", "Notes / source"])

A = [
    ("Planning year", 2027, "year", "Calendar-year plan. Change this and the Sprint 1 start date together.", NUM_INT, "key"),
    ("Sprint length", WEEKS, "weeks", f"{WEEKS}-week cadence. Drives weekdays per sprint below and every date on "
                                      "'Sprint Calendar'. Change it and you must also set 'Sprints in plan' and add or "
                                      "delete rows on 'Sprint Calendar' and 'Sprint Plan'.", NUM_INT, "key"),
    ("Sprints in plan", NS, "sprints", f"{NS} x {WEEKS} weeks = {NS*STEP} days: Mon 4 Jan to Fri 24 Dec 2027, leaving "
                                       "the last week of December as shutdown and slack. Must match the number of rows "
                                       "on 'Sprint Calendar' and 'Sprint Plan'.", NUM_INT, "key"),
    ("Sprint 1 start date", dt.date(2027, 1, 4), "date", "First Monday of FY2027. Every other sprint date derives from it.", NUM_DATE, "key"),
    ("Weekdays per sprint", "=$B$6*5", "days", "Formula: sprint length x 5. Before company holidays, which "
                                               "'Sprint Calendar' nets out per sprint.", NUM_INT, "calc"),
    ("Focus factor", 0.70, "share of a dev-day", "Share of a working day left for sprint backlog after ceremonies, reviews, interviews, admin. Example value; typical observed range 60-75%. Measure your own.", NUM_PCT, "key"),
    ("Throughput baseline", 1.25, "points per focused dev-day", "KEY LEVER. Per dev-day, so it does not change when the sprint length does. Example value derived from the illustrative FY2026 actuals in the 'Dashboard' baseline reconciliation. Story points are team-relative: derive this from your own history, never from another team's.", '0.00', "key"),
    ("Unplanned-work reserve", 0.15, "share of capacity", "Held back for escalations, production bugs and scope found mid-sprint. Example value; 20-30% if your escalation load is heavy.", NUM_PCT, "key"),
    ("Support / on-call cover", 1.0, "people per sprint", "Engineers off the sprint board on rotation for the whole sprint.", '0.0', "plain"),
    ("Default estimate confidence", 0.70, "share", "Used to risk-adjust epic estimates on 'Annual Plan' (risk-adjusted = estimate / confidence).", NUM_PCT, "plain"),
]
r = 5
for label, val, unit, note, fmt, kind in A:
    asm.cell(row=r, column=1, value=label).font = F(10, True)
    asm.cell(row=r, column=1).border = BOX
    c = asm.cell(row=r, column=2, value=val); c.number_format = fmt
    if kind == "calc":
        calc(c, True)
        c.fill = PatternFill("solid", fgColor=GREY)
    else:
        inp(c, kind == "key")
    c.alignment = Alignment(horizontal="center")
    u = asm.cell(row=r, column=3, value=unit); u.font = F(10, color="595959"); u.border = BOX
    n = asm.cell(row=r, column=4, value=note); n.font = F(9, color="404040")
    n.alignment = Alignment(vertical="top", wrap_text=True); n.border = BOX
    asm.row_dimensions[r].height = max(15, 12 * (len(note) // 95 + 1))
    r += 1
WEEKSC, SPRINTS, S1, WKDAYS = "$B$6", "$B$7", "$B$8", "$B$9"
FOCUS, THRU, RESV, ONCALL, CONF = "$B$10", "$B$11", "$B$12", "$B$13", "$B$14"
asm["B10"].comment = Comment("Example value, not a measurement of your team. Run a two-sprint time study (calendar + "
                             "interrupt log) and replace it.", "Baseline lever")
asm["B11"].comment = Comment(f"Points per focused dev-day. Derived here from FY2026 example actuals on 'Dashboard' "
                             f"({FY26_VEL:.1f} pts / ({FY26_DAYS:.1f} dev-days x 70%) = 1.29) and rounded down to 1.25 "
                             "to stay conservative. Being per-dev-day, it survives a change of sprint length.",
                             "Baseline lever")
asm["B12"].comment = Comment("Capacity deliberately left unplanned. If sprints routinely end with carry-over, raise "
                             "this before raising estimates.", "Baseline lever")

section(asm, 16, "Target investment mix")
header_row(asm, 17, ["Workstream", "Target share", "", "Notes / source"])
mix = [("Product / Roadmap", 0.55), ("Platform / Scalability", 0.20),
       ("Tech Debt / Quality", 0.15), ("KTLO / Support", 0.10)]
r = 18
for name, share in mix:
    asm.cell(row=r, column=1, value=name).font = F(10, True)
    asm.cell(row=r, column=1).border = BOX
    c = asm.cell(row=r, column=2, value=share); c.number_format = NUM_PCT; inp(c, True)
    c.alignment = Alignment(horizontal="center")
    asm.cell(row=r, column=3).border = BOX
    asm.cell(row=r, column=4).border = BOX
    r += 1
asm.cell(row=22, column=1, value="Total").font = F(10, True)
t = asm.cell(row=22, column=2, value="=SUM(B18:B21)"); t.number_format = NUM_PCT
calc(t, True); t.fill = PatternFill("solid", fgColor=GREY)
t.alignment = Alignment(horizontal="center")
asm.cell(row=22, column=4, value="Must equal 100%. Example split; set it to whatever your leadership has agreed.").font = F(9, color="404040")
asm.cell(row=22, column=1).border = BOX

section(asm, 24, "Derived baseline (formulas — do not overwrite)")
header_row(asm, 25, ["Metric", "Value", "Unit", "Formula source"])
D = [
    ("Effective team FTE", "='Team Capacity'!$G$13", NUM_FTE, "FTE", "Roster total: FTE x allocation x ramp"),
    ("Gross dev-days per sprint", "=$B$26*$B$9", NUM_DAY, "dev-days", "Effective FTE x weekdays per sprint"),
    ("Net dev-days per sprint (plan average)", f"=AVERAGE('Sprint Plan'!$H${R0}:$H${R1})", NUM_DAY, "dev-days", "After holidays, PTO and on-call"),
    ("Baseline velocity per sprint (average)", f"=AVERAGE('Sprint Plan'!$I${R0}:$I${R1})", NUM_PTS, "story points", "Net dev-days x focus factor x throughput"),
    ("Commitable velocity per sprint (average)", f"=AVERAGE('Sprint Plan'!$K${R0}:$K${R1})", NUM_PTS, "story points", "Baseline less the unplanned-work reserve"),
    ("Baseline annual capacity", f"=SUM('Sprint Plan'!$I${R0}:$I${R1})", NUM_PTS, "story points", f"Sum of all {NS} sprints"),
    ("Commitable annual capacity", f"=SUM('Sprint Plan'!$K${R0}:$K${R1})", NUM_PTS, "story points", "What the roadmap may be planned against"),
    ("Roadmap demand (risk-adjusted)", "=SUM('Annual Plan'!$H$4:$H$23)", NUM_PTS, "story points", "Sum of risk-adjusted epic sizes"),
    ("Annual load", "=IF($B$32=0,\"\",$B$33/$B$32)", NUM_PCT, "% of commitable", "Demand / commitable capacity"),
]
r = 26
for label, f, fmt, unit, note in D:
    asm.cell(row=r, column=1, value=label).font = F(10, True)
    asm.cell(row=r, column=1).border = BOX
    c = asm.cell(row=r, column=2, value=f); c.number_format = fmt
    link(c); c.alignment = Alignment(horizontal="center")
    c.fill = PatternFill("solid", fgColor=GREY)
    u = asm.cell(row=r, column=3, value=unit); u.font = F(10, color="595959"); u.border = BOX
    n = asm.cell(row=r, column=4, value=note); n.font = F(9, color="404040"); n.border = BOX
    r += 1
asm.freeze_panes = "A5"

# =====================================================================
# 3. HOLIDAYS
# =====================================================================
hol = wb.create_sheet("Holidays")
hol.sheet_properties.tabColor = "8EAADB"
title(hol, "Company Holidays — FY2027",
      "Example list: 2027 US federal holidays plus two winter-shutdown days. Replace with your company's calendar; "
      "'Sprint Calendar' counts whatever is in rows 5-25.")
widths(hol, {"A": 20, "B": 8, "C": 34, "D": 56})
header_row(hol, 4, ["Date", "Day", "Holiday", "Notes"])
H = [
    (dt.date(2027, 1, 1), "New Year's Day", "Falls before Sprint 1 starts (4 Jan), so it costs the plan nothing."),
    (dt.date(2027, 1, 18), "Martin Luther King Jr. Day", "3rd Monday in January."),
    (dt.date(2027, 2, 15), "Presidents' Day", "3rd Monday in February."),
    (dt.date(2027, 5, 31), "Memorial Day", "Last Monday in May."),
    (dt.date(2027, 6, 18), "Juneteenth (observed)", "19 June 2027 is a Saturday."),
    (dt.date(2027, 7, 5), "Independence Day (observed)", "4 July 2027 is a Sunday."),
    (dt.date(2027, 9, 6), "Labor Day", "1st Monday in September."),
    (dt.date(2027, 11, 25), "Thanksgiving", "4th Thursday in November."),
    (dt.date(2027, 11, 26), "Day after Thanksgiving", "Company holiday, not a federal one."),
    (dt.date(2027, 12, 23), "Winter shutdown", "Example company day. Delete if you do not close."),
    (dt.date(2027, 12, 24), "Christmas Day (observed)", "25 December 2027 is a Saturday. Last working day of the plan."),
    (dt.date(2027, 12, 31), "New Year's Day 2028 (observed)",
     f"Falls after Sprint {NS} ends (24 Dec), in the leftover week, so it costs the plan nothing."),
]
r = 5
for d, name, note in H:
    c = hol.cell(row=r, column=1, value=d); c.number_format = NUM_DATE; inp(c)
    dy = hol.cell(row=r, column=2, value=f'=TEXT(A{r},"ddd")'); calc(dy)
    dy.alignment = Alignment(horizontal="center")
    n = hol.cell(row=r, column=3, value=name); inp(n)
    nt = hol.cell(row=r, column=4, value=note); nt.font = F(9, color="404040"); nt.border = BOX
    nt.alignment = Alignment(vertical="top", wrap_text=True)
    r += 1
for rr in range(r, 26):
    for cc in range(1, 5):
        hol.cell(row=rr, column=cc).border = BOX
    hol.cell(row=rr, column=1).number_format = NUM_DATE
    hol.cell(row=rr, column=2, value=f'=IF($A{rr}="","",TEXT($A{rr},"ddd"))').font = F(10)
hol.cell(row=27, column=1, value="Holidays falling inside the sprint dates:").font = F(10, True)
cnt = hol.cell(row=27, column=4, value=f"=SUM('Sprint Calendar'!$G${R0}:$G${R1})")
calc(cnt, True); cnt.number_format = NUM_INT; cnt.fill = PatternFill("solid", fgColor=GREY)
hol.freeze_panes = "A5"
HOLRANGE = "'Holidays'!$A$5:$A$25"

# =====================================================================
# 4. SPRINT CALENDAR
# =====================================================================
cal = wb.create_sheet("Sprint Calendar")
cal.sheet_properties.tabColor = "8EAADB"
title(cal, f"Sprint Calendar — FY2027 ({NS} {WEEKS}-week sprints)",
      "Dates derive from the sprint length and Sprint 1 start date on 'Assumptions'. Only the Theme / Milestone "
      "column is for typing in.")
widths(cal, {"A": 9, "B": 9, "C": 19, "D": 19, "E": 9, "F": 11, "G": 11, "H": 13, "I": 48})
header_row(cal, 3, ["Sprint #", "Sprint", "Start (Mon)", "End (Fri)", "Quarter", "Weekdays",
                    "Company Holidays", "Net Working Days", "Theme / Milestone"])
themes = {1: "Plan kickoff; checkout redesign discovery", 2: "Billing service extraction begins",
          4: "Checkout phase 1 ships", 5: "Q1 review; quarter boundary falls mid-sprint",
          6: "Q2 planning; checkout phase 2 begins", 9: "Mid-year review; H2 re-baseline (boundary mid-sprint)",
          10: "Data platform migration starts", 13: "Enterprise SSO GA target (boundary mid-sprint)",
          14: "Q4 planning; code-freeze dates agreed", 16: "Pre-freeze hardening; Thanksgiving week",
          17: f"Winter shutdown; plan FY2028 (ends 24 Dec)"}
for i in range(NS):
    r = R0 + i
    a = cal.cell(row=r, column=1, value=(1 if i == 0 else f"=A{r-1}+1")); calc(a)
    a.number_format = NUM_INT; a.alignment = Alignment(horizontal="center")
    b = cal.cell(row=r, column=2, value=f'="S"&TEXT($A{r},"00")'); calc(b, True)
    b.alignment = Alignment(horizontal="center")
    if i == 0:
        c = cal.cell(row=r, column=3, value=f"='Assumptions'!{S1}"); link(c)
    else:
        c = cal.cell(row=r, column=3, value=f"=C{r-1}+'Assumptions'!{WEEKSC}*7"); calc(c)
    c.number_format = NUM_DATE
    d = cal.cell(row=r, column=4, value=f"=C{r}+'Assumptions'!{WEEKSC}*7-3"); calc(d); d.number_format = NUM_DATE
    e = cal.cell(row=r, column=5, value=f'="Q"&ROUNDUP(MONTH($C{r})/3,0)'); calc(e)
    e.alignment = Alignment(horizontal="center")
    f = cal.cell(row=r, column=6, value=f"='Assumptions'!{WKDAYS}"); link(f)
    f.number_format = NUM_INT; f.alignment = Alignment(horizontal="center")
    g = cal.cell(row=r, column=7, value=f'=COUNTIFS({HOLRANGE},">="&$C{r},{HOLRANGE},"<="&$D{r})')
    calc(g); g.number_format = NUM_INT; g.alignment = Alignment(horizontal="center")
    h = cal.cell(row=r, column=8, value=f"=$F{r}-$G{r}"); calc(h, True)
    h.number_format = NUM_INT; h.alignment = Alignment(horizontal="center")
    t = cal.cell(row=r, column=9, value=themes.get(i + 1, "")); inp(t)
    t.font = F(9, color=BLUE_TXT)
    if i % 2 == 1:
        for cc in range(1, 10):
            cal.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=BAND)
cal.cell(row=RT, column=2, value="TOTAL").font = F(10, True)
for col in ("F", "G", "H"):
    c = cal[f"{col}{RT}"]; c.value = f"=SUM({col}{R0}:{col}{R1})"; calc(c, True); c.number_format = NUM_INT
    c.fill = PatternFill("solid", fgColor=GREY); c.alignment = Alignment(horizontal="center")
cal.cell(row=RT, column=9, value="Net working days drive every capacity number on 'Sprint Plan'.").font = F(9, color="404040")
cal.freeze_panes = "C4"

# =====================================================================
# 5. TEAM CAPACITY
# =====================================================================
cap = wb.create_sheet("Team Capacity")
cap.sheet_properties.tabColor = "538135"
title(cap, "Team Capacity — Roster and Effective FTE",
      "Example roster. Replace the people, FTE, allocation, ramp and PTO with yours; everything downstream recalculates.")
widths(cap, {"A": 20, "B": 20, "C": 8, "D": 15, "E": 12, "F": 17, "G": 14, "H": 18, "I": 44})
header_row(cap, 3, ["Member", "Role", "FTE", "Team Allocation %", "Ramp Factor",
                    "Planned PTO (days/yr)", "Effective FTE", "Avg PTO Days / Sprint", "Notes"])
ROSTER = [
    ("A. Rivera", "Senior Engineer", 1.0, 1.00, 1.00, 20, ""),
    ("B. Chen", "Senior Engineer", 1.0, 1.00, 1.00, 18, ""),
    ("C. Okafor", "Engineer II", 1.0, 1.00, 1.00, 15, ""),
    ("D. Nowak", "Engineer II", 1.0, 1.00, 1.00, 20, ""),
    ("E. Haddad", "Engineer I", 1.0, 1.00, 0.60, 12, "Joined Nov 2026; ramp factor 60% for the first half, review at Q2."),
    ("F. Silva", "QA Engineer", 1.0, 0.80, 1.00, 18, "20% of time on a second team's release testing."),
    ("G. Muller", "Tech Lead", 1.0, 0.50, 1.00, 22, "Half the week on architecture, hiring and cross-team work."),
]
r = 4
for name, role, fte, alloc, ramp, pto, note in ROSTER:
    cap.cell(row=r, column=1, value=name).font = F(10, True); cap.cell(row=r, column=1).border = BOX
    x = cap.cell(row=r, column=2, value=role); inp(x)
    for col, val, fmt in ((3, fte, NUM_FTE), (4, alloc, NUM_PCT), (5, ramp, NUM_PCT), (6, pto, NUM_INT)):
        c = cap.cell(row=r, column=col, value=val); c.number_format = fmt; inp(c)
        c.alignment = Alignment(horizontal="center")
    g = cap.cell(row=r, column=7, value=f"=$C{r}*$D{r}*$E{r}"); calc(g, True)
    g.number_format = NUM_FTE; g.alignment = Alignment(horizontal="center")
    h = cap.cell(row=r, column=8, value=f"=IF('Assumptions'!{SPRINTS}=0,0,$F{r}/'Assumptions'!{SPRINTS})")
    calc(h); h.number_format = NUM_DAY; h.alignment = Alignment(horizontal="center")
    n = cap.cell(row=r, column=9, value=note); n.font = F(9, color="404040"); n.border = BOX
    n.alignment = Alignment(vertical="top", wrap_text=True)
    r += 1
for rr in range(11, 13):
    for cc in range(1, 10):
        cap.cell(row=rr, column=cc).border = BOX
    cap.cell(row=rr, column=7, value=f"=$C{rr}*$D{rr}*$E{rr}").number_format = NUM_FTE
    cap.cell(row=rr, column=8, value=f"=IF('Assumptions'!{SPRINTS}=0,0,$F{rr}/'Assumptions'!{SPRINTS})").number_format = NUM_DAY
    for cc in (3, 4, 5, 6):
        inp(cap.cell(row=rr, column=cc))
    cap.cell(row=rr, column=3).number_format = NUM_FTE
    cap.cell(row=rr, column=4).number_format = NUM_PCT
    cap.cell(row=rr, column=5).number_format = NUM_PCT
cap.cell(row=13, column=1, value="TOTAL").font = F(10, True)
for col, fmt in (("C", NUM_FTE), ("F", NUM_INT), ("G", NUM_FTE), ("H", NUM_DAY)):
    c = cap[f"{col}13"]; c.value = f"=SUM({col}4:{col}12)"; calc(c, True)
    c.number_format = fmt; c.fill = PatternFill("solid", fgColor=GREY)
    c.alignment = Alignment(horizontal="center")
cap.cell(row=13, column=9, value="Effective FTE is the number 'Sprint Plan' multiplies by working days.").font = F(9, color="404040")

section(cap, 15, "PTO reconciliation")
checks = [
    ("Leave days entered on 'Sprint Plan' (col F)", f"=SUM('Sprint Plan'!$F${R0}:$F${R1})"),
    ("Roster PTO total (col F above)", "=$F$13"),
    ("Difference — should be zero", "=$D$16-$D$17"),
]
r = 16
for label, f in checks:
    cap.cell(row=r, column=1, value=label).font = F(10, True if r == 18 else False)
    c = cap.cell(row=r, column=4, value=f); calc(c, r == 18); c.number_format = NUM_DAY
    c.alignment = Alignment(horizontal="center")
    if r == 18:
        c.fill = PatternFill("solid", fgColor=GREY)
    r += 1
cap.cell(row=18, column=9, value="Non-zero means the sprint-by-sprint spread no longer matches the roster.").font = F(9, color="404040")
cap.freeze_panes = "A4"
EFTE = "'Team Capacity'!$G$13"

# =====================================================================
# 6. SPRINT PLAN
# =====================================================================
sp = wb.create_sheet("Sprint Plan")
sp.sheet_properties.tabColor = "C00000"
title(sp, f"Sprint Plan — Baseline Capacity, Commitment and Actuals (FY2027, {WEEKS}-week sprints)",
      "Baseline velocity = net dev-days x focus factor x throughput. Commitable = baseline less the unplanned-work "
      "reserve. Type only in the blue columns: PTO / Leave (F), Completed (N), Notes (Q).")
widths(sp, {"A": 9, "B": 9, "C": 9, "D": 11, "E": 12, "F": 12, "G": 13, "H": 11, "I": 13,
            "J": 12, "K": 12, "L": 13, "M": 10, "N": 12, "O": 12, "P": 14, "Q": 40})
header_row(sp, 3, ["Sprint #", "Sprint", "Quarter", "Working Days", "Gross Dev-Days", "PTO / Leave Days",
                   "Support & On-Call Days", "Net Dev-Days", "Baseline Velocity (pts)",
                   "Unplanned Reserve (pts)", "Commitable (pts)", "Planned from Roadmap (pts)", "Load %",
                   "Completed (pts)", "Variance vs Planned (pts)", "Rolling 3-Sprint Velocity (pts)", "Notes"])
# 17 three-week sprints; leave peaks in July/August and December
PTO = [5, 6, 7, 6, 6,  7, 8, 7, 7,  10, 11, 9, 7,  6, 6, 7, 10]
assert len(PTO) == NS and sum(PTO) == 125, (len(PTO), sum(PTO))
EXAMPLE = {1: (52, "EXAMPLE actuals — clear column N before you use this plan."),
           2: (45, "EXAMPLE actuals — two-day incident ate the reserve.")}
for i in range(NS):
    r = R0 + i
    cr = R0 + i
    for col, f in ((1, f"='Sprint Calendar'!$A${cr}"), (2, f"='Sprint Calendar'!$B${cr}"),
                   (3, f"='Sprint Calendar'!$E${cr}"), (4, f"='Sprint Calendar'!$H${cr}")):
        c = sp.cell(row=r, column=col, value=f); link(c)
        c.alignment = Alignment(horizontal="center")
        c.number_format = NUM_INT if col in (1, 4) else 'General'
    sp.cell(row=r, column=2).font = F(10, True, GREEN_TXT)
    e = sp.cell(row=r, column=5, value=f"={EFTE}*$D{r}"); calc(e); e.number_format = NUM_DAY
    f_ = sp.cell(row=r, column=6, value=PTO[i]); inp(f_); f_.number_format = NUM_DAY
    f_.alignment = Alignment(horizontal="center")
    g = sp.cell(row=r, column=7, value=f"='Assumptions'!{ONCALL}*$D{r}"); calc(g); g.number_format = NUM_DAY
    h = sp.cell(row=r, column=8, value=f"=$E{r}-$F{r}-$G{r}"); calc(h, True); h.number_format = NUM_DAY
    ii = sp.cell(row=r, column=9, value=f"=$H{r}*'Assumptions'!{FOCUS}*'Assumptions'!{THRU}")
    calc(ii, True); ii.number_format = NUM_PTS
    j = sp.cell(row=r, column=10, value=f"=$I{r}*'Assumptions'!{RESV}"); calc(j); j.number_format = NUM_PTS
    k = sp.cell(row=r, column=11, value=f"=$I{r}-$J{r}"); calc(k, True); k.number_format = NUM_PTS
    k.fill = PatternFill("solid", fgColor="E2EFDA")
    l = sp.cell(row=r, column=12,
                value=f"=SUMPRODUCT(('Annual Plan'!$J$4:$J$23<=$A{r})*('Annual Plan'!$K$4:$K$23>=$A{r})"
                      f"*'Annual Plan'!$M$4:$M$23)")
    calc(l); l.number_format = NUM_PTS
    m = sp.cell(row=r, column=13, value=f'=IF($K{r}=0,"",$L{r}/$K{r})'); calc(m, True)
    m.number_format = NUM_PCT; m.alignment = Alignment(horizontal="center")
    n = sp.cell(row=r, column=14, value=EXAMPLE.get(i + 1, (None,))[0]); inp(n); n.number_format = NUM_PTS
    o = sp.cell(row=r, column=15, value=f'=IF($N{r}="","",$N{r}-$L{r})'); calc(o); o.number_format = NUM_PTS
    if r >= R0 + 2:
        p = sp.cell(row=r, column=16, value=f'=IF(COUNT($N{r-2}:$N{r})=3,AVERAGE($N{r-2}:$N{r}),"")')
    else:
        p = sp.cell(row=r, column=16, value="")
    calc(p); p.number_format = NUM_PTS
    q = sp.cell(row=r, column=17, value=EXAMPLE.get(i + 1, (None, ""))[1]); inp(q)
    q.font = F(9, color=BLUE_TXT)
    if i % 2 == 1:
        for cc in list(range(1, 11)) + list(range(12, 18)):
            if sp.cell(row=r, column=cc).fill.fgColor.rgb in (None, "00000000"):
                sp.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=BAND)
sp.cell(row=RT, column=2, value="TOTAL").font = F(10, True)
for col, fmt in (("D", NUM_INT), ("E", NUM_DAY), ("F", NUM_DAY), ("G", NUM_DAY), ("H", NUM_DAY),
                 ("I", NUM_PTS), ("J", NUM_PTS), ("K", NUM_PTS), ("L", NUM_PTS), ("N", NUM_PTS), ("O", NUM_PTS)):
    c = sp[f"{col}{RT}"]; c.value = f"=SUM({col}{R0}:{col}{R1})"; calc(c, True); c.number_format = fmt
    c.fill = PatternFill("solid", fgColor=GREY)
m = sp[f"M{RT}"]; m.value = f'=IF($K{RT}=0,"",$L{RT}/$K{RT})'; calc(m, True); m.number_format = NUM_PCT
m.fill = PatternFill("solid", fgColor=GREY); m.alignment = Alignment(horizontal="center")
p = sp[f"P{RT}"]; p.value = f'=IF(COUNT($N${R0}:$N${R1})=0,"",AVERAGE($N${R0}:$N${R1}))'; calc(p, True)
p.number_format = NUM_PTS; p.fill = PatternFill("solid", fgColor=GREY)
sp.cell(row=RT, column=17, value=f"Row {RT} is the annual baseline: total commitable (K) vs total planned (L).").font = F(9, color="404040")
sp.cell(row=RT+2, column=2, value="Load % colour: green under 90%, amber 90-100%, red over 100% (overcommitted).").font = F(9, color="404040", it=True)
OVER_SP_ROW = RT + 3
sp.cell(row=OVER_SP_ROW, column=2, value="The red sprints are deliberate. S03, S07, S08, S11, S12 and S16 each lose 1-2 "
        "working days to holidays while the roadmap spreads its points evenly, so they tip over 100%. Level the plan by "
        "moving scope into S05, S09, S13, S14 and S17 — the quarter-boundary sprints sitting at 49-71%, because the next "
        "quarter's epics have not started yet. Every quarter is green at 86-90%; individual sprints are not.").font = F(9, color="404040", it=True)
sp.cell(row=RT+4, column=2, value=f"A {WEEKS}-week cadence absorbs a single holiday or a week of leave better than a "
        "2-week one, because each absence is a smaller share of the sprint — but each commitment is larger, so a "
        "mis-estimate costs more before you find out.").font = F(9, color="404040", it=True)

sp.conditional_formatting.add(f"M{R0}:M{R1}", CellIsRule(operator="greaterThan", formula=["1"],
                              fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, size=10, bold=True, color="9C0006")))
sp.conditional_formatting.add(f"M{R0}:M{R1}", CellIsRule(operator="between", formula=["0.9", "1"],
                              fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(name=FONT, size=10, color="9C6500")))
sp.conditional_formatting.add(f"M{R0}:M{R1}", CellIsRule(operator="lessThan", formula=["0.9"],
                              fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONT, size=10, color="006100")))
sp.conditional_formatting.add(f"O{R0}:O{R1}", CellIsRule(operator="lessThan", formula=["0"],
                              font=Font(name=FONT, size=10, color="9C0006")))
sp.freeze_panes = "D4"

# =====================================================================
# 7. ANNUAL PLAN
# =====================================================================
ap = wb.create_sheet("Annual Plan")
ap.sheet_properties.tabColor = "BF8F00"
title(ap, "Annual Plan — Roadmap Demand in Story Points (FY2027)",
      f"Risk-adjusted points = estimate / confidence. Points per sprint = risk-adjusted / sprints spanned, and that is "
      f"what lands in 'Sprint Plan' column L. Sprint numbers run 1-{NS} on the {WEEKS}-week calendar: "
      f"Q1 = S01-S05, Q2 = S06-S09, Q3 = S10-S13, Q4 = S14-S{NS}. 16 example initiatives; rows 20-23 are blank for yours.")
widths(ap, {"A": 9, "B": 38, "C": 22, "D": 9, "E": 13, "F": 13, "G": 12, "H": 15, "I": 10,
            "J": 12, "K": 12, "L": 13, "M": 14, "N": 13})
header_row(ap, 3, ["ID", "Initiative / Epic", "Workstream", "Priority", "Owner", "Estimate (pts)",
                   "Confidence", "Risk-Adjusted (pts)", "Target Quarter", "Start Sprint #", "End Sprint #",
                   "Sprints Spanned", "Points per Sprint", "Status"])
PROD, PLAT, DEBT, KTLO = "Product / Roadmap", "Platform / Scalability", "Tech Debt / Quality", "KTLO / Support"
INIT = [
    ("INI-01", "Checkout redesign (phase 1)", PROD, "P0", "B. Chen", 70, 0.75, "Q1", 1, 4, "Committed"),
    ("INI-02", "Billing service extraction", PLAT, "P0", "A. Rivera", 40, 0.70, "Q1", 1, 5, "Committed"),
    ("INI-03", "Flaky test remediation", DEBT, "P1", "C. Okafor", 20, 0.80, "Q1", 2, 5, "Committed"),
    ("INI-04", "KTLO, bug SLA and on-call load", KTLO, "P1", "G. Muller", 30, 0.90, "Q1", 1, 5, "Committed"),
    ("INI-05", "Checkout redesign (phase 2)", PROD, "P0", "B. Chen", 48, 0.75, "Q2", 6, 8, "Committed"),
    ("INI-06", "Mobile feature parity", PROD, "P1", "D. Nowak", 27, 0.65, "Q2", 6, 9, "Committed"),
    ("INI-07", "Observability and SLO upgrade", PLAT, "P1", "A. Rivera", 25, 0.80, "Q2", 7, 9, "Committed"),
    ("INI-08", "KTLO, bug SLA and on-call load", KTLO, "P1", "G. Muller", 25, 0.90, "Q2", 6, 9, "Committed"),
    ("INI-09", "Enterprise SSO and SCIM", PROD, "P0", "C. Okafor", 40, 0.70, "Q3", 10, 12, "Planned"),
    ("INI-10", "Data platform migration", PLAT, "P0", "A. Rivera", 35, 0.60, "Q3", 10, 13, "Planned"),
    ("INI-11", "Accessibility compliance (WCAG 2.2)", DEBT, "P1", "F. Silva", 18, 0.85, "Q3", 11, 13, "Planned"),
    ("INI-12", "KTLO, bug SLA and on-call load", KTLO, "P1", "G. Muller", 28, 0.90, "Q3", 10, 13, "Planned"),
    ("INI-13", "Pricing and packaging experiments", PROD, "P1", "D. Nowak", 40, 0.70, "Q4", 14, 16, "Planned"),
    ("INI-14", "Infrastructure cost optimisation", PLAT, "P2", "A. Rivera", 25, 0.75, "Q4", 14, 17, "Stretch"),
    ("INI-15", "Legacy API v1 deprecation", DEBT, "P2", "C. Okafor", 30, 0.75, "Q4", 15, 17, "Planned"),
    ("INI-16", "KTLO, bug SLA and on-call load", KTLO, "P1", "G. Muller", 24, 0.90, "Q4", 14, 17, "Planned"),
]
for i, (iid, name, wsname, pri, owner, est, conf, q, s0, s1_, status) in enumerate(INIT):
    r = 4 + i
    for col, val in ((1, iid), (2, name), (3, wsname), (4, pri), (5, owner)):
        c = ap.cell(row=r, column=col, value=val); inp(c)
        if col == 2:
            c.font = F(10, True, BLUE_TXT)
    fc = ap.cell(row=r, column=6, value=est); inp(fc); fc.number_format = NUM_PTS
    fc.alignment = Alignment(horizontal="center")
    gc = ap.cell(row=r, column=7, value=conf); inp(gc, True); gc.number_format = NUM_PCT
    gc.alignment = Alignment(horizontal="center")
    hc = ap.cell(row=r, column=8, value=f'=IF($F{r}="","",IF($G{r}=0,$F{r},$F{r}/$G{r}))')
    calc(hc, True); hc.number_format = NUM_PTS
    ic = ap.cell(row=r, column=9, value=q); inp(ic); ic.alignment = Alignment(horizontal="center")
    jc = ap.cell(row=r, column=10, value=s0); inp(jc); jc.number_format = NUM_INT
    jc.alignment = Alignment(horizontal="center")
    kc = ap.cell(row=r, column=11, value=s1_); inp(kc); kc.number_format = NUM_INT
    kc.alignment = Alignment(horizontal="center")
    lc = ap.cell(row=r, column=12, value=f'=IF(OR($J{r}="",$K{r}=""),0,$K{r}-$J{r}+1)')
    calc(lc); lc.number_format = NUM_INT; lc.alignment = Alignment(horizontal="center")
    mc = ap.cell(row=r, column=13, value=f'=IF(OR($L{r}=0,$H{r}=""),0,$H{r}/$L{r})')
    calc(mc, True); mc.number_format = NUM_PTS
    nc = ap.cell(row=r, column=14, value=status); inp(nc); nc.alignment = Alignment(horizontal="center")
    if i % 2 == 1:
        for cc in range(1, 15):
            if ap.cell(row=r, column=cc).fill.fgColor.rgb in (None, "00000000"):
                ap.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=BAND)
for r in range(20, 24):
    for cc in (1, 2, 3, 4, 5, 9, 14):
        inp(ap.cell(row=r, column=cc))
    for cc, fmt in ((6, NUM_PTS), (7, NUM_PCT), (10, NUM_INT), (11, NUM_INT)):
        c = ap.cell(row=r, column=cc); inp(c); c.number_format = fmt
        c.alignment = Alignment(horizontal="center")
    hc = ap.cell(row=r, column=8, value=f'=IF($F{r}="","",IF($G{r}=0,$F{r},$F{r}/$G{r}))')
    calc(hc); hc.number_format = NUM_PTS
    lc = ap.cell(row=r, column=12, value=f'=IF(OR($J{r}="",$K{r}=""),0,$K{r}-$J{r}+1)')
    calc(lc); lc.number_format = NUM_INT
    mc = ap.cell(row=r, column=13, value=f'=IF(OR($L{r}=0,$H{r}=""),0,$H{r}/$L{r})')
    calc(mc); mc.number_format = NUM_PTS
ap.cell(row=24, column=2, value="TOTAL — roadmap demand").font = F(10, True)
for col, fmt in (("F", NUM_PTS), ("H", NUM_PTS), ("M", NUM_PTS)):
    c = ap[f"{col}24"]; c.value = f"=SUM({col}4:{col}23)"; calc(c, True); c.number_format = fmt
    c.fill = PatternFill("solid", fgColor=GREY)
ap.cell(row=25, column=2, value="Commitable capacity for the year (from 'Sprint Plan')").font = F(10, True)
c = ap["H25"]; c.value = f"=SUM('Sprint Plan'!$K${R0}:$K${R1})"; link(c); c.number_format = NUM_PTS
c.fill = PatternFill("solid", fgColor=GREY)
ap.cell(row=26, column=2, value="Annual load (demand / commitable)").font = F(10, True)
c = ap["H26"]; c.value = '=IF($H$25=0,"",$H$24/$H$25)'; calc(c, True); c.number_format = NUM_PCT
c.fill = PatternFill("solid", fgColor=GREY)
ap.cell(row=27, column=2, value="Rule of thumb: start the year near 85-90%. At 100% the first surprise breaks the plan.").font = F(9, color="404040", it=True)
ap.cell(row=28, column=2, value=f"Points per sprint rose with the cadence: the same epic spread over {WEEKS}-week "
        "sprints carries roughly half again as many points per sprint as it did over 2-week ones.").font = F(9, color="404040", it=True)

dv_ws = DataValidation(type="list", formula1=f'"{PROD},{PLAT},{DEBT},{KTLO}"', allow_blank=True)
dv_pr = DataValidation(type="list", formula1='"P0,P1,P2,P3"', allow_blank=True)
dv_q = DataValidation(type="list", formula1='"Q1,Q2,Q3,Q4"', allow_blank=True)
dv_st = DataValidation(type="list", formula1='"Committed,Planned,Stretch,Parked,Done"', allow_blank=True)
for dv, rng in ((dv_ws, "C4:C23"), (dv_pr, "D4:D23"), (dv_q, "I4:I23"), (dv_st, "N4:N23")):
    ap.add_data_validation(dv); dv.add(rng)
ap.auto_filter.ref = "A3:N23"
ap.freeze_panes = "C4"

# =====================================================================
# 8. DASHBOARD
# =====================================================================
db = wb.create_sheet("Dashboard")
db.sheet_properties.tabColor = "7030A0"
title(db, f"Dashboard — Supply vs Demand, FY2027 ({WEEKS}-week sprints)",
      "Everything here is a formula. The only inputs are the two FY2026 actuals in the baseline reconciliation block.")
widths(db, {"A": 42, "B": 15, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 11, "I": 15, "J": 3})

section(db, 3, "Annual roll-up")
ANN = [
    ("Sprints in plan", f"='Assumptions'!{SPRINTS}", NUM_INT),
    ("Net working days", f"=SUM('Sprint Calendar'!$H${R0}:$H${R1})", NUM_INT),
    ("Net dev-days available", f"=SUM('Sprint Plan'!$H${R0}:$H${R1})", NUM_DAY),
    ("Baseline capacity (pts)", f"=SUM('Sprint Plan'!$I${R0}:$I${R1})", NUM_PTS),
    ("Unplanned-work reserve (pts)", f"=SUM('Sprint Plan'!$J${R0}:$J${R1})", NUM_PTS),
    ("Commitable capacity (pts)", f"=SUM('Sprint Plan'!$K${R0}:$K${R1})", NUM_PTS),
    ("Roadmap demand, risk-adjusted (pts)", "=SUM('Annual Plan'!$H$4:$H$23)", NUM_PTS),
    ("Headroom (pts)", "=$B$9-$B$10", NUM_PTS),
    ("Annual load %", '=IF($B$9=0,"",$B$10/$B$9)', NUM_PCT),
    ("Average baseline velocity per sprint (pts)", '=IF($B$4=0,"",$B$7/$B$4)', NUM_PTS),
]
r = 4
for label, f, fmt in ANN:
    c0 = db.cell(row=r, column=1, value=label); c0.font = F(10, True); c0.border = BOX
    c = db.cell(row=r, column=2, value=f); link(c); c.number_format = fmt
    c.alignment = Alignment(horizontal="center")
    c.fill = PatternFill("solid", fgColor=GREY)
    r += 1
db["B11"].font = F(10, True, GREEN_TXT)
db["B12"].font = F(11, True, GREEN_TXT)

section(db, 15, "By quarter")
header_row(db, 16, ["Quarter", "Sprints", "Net Dev-Days", "Baseline (pts)", "Reserve (pts)",
                    "Commitable (pts)", "Planned (pts)", "Load %", "Headroom (pts)"])
for i, q in enumerate(["Q1", "Q2", "Q3", "Q4"]):
    r = 17 + i
    c0 = db.cell(row=r, column=1, value=q); c0.font = F(10, True); c0.border = BOX
    c0.alignment = Alignment(horizontal="center")
    specs = [
        (2, f"=COUNTIFS('Sprint Plan'!$C${R0}:$C${R1},$A{r})", NUM_INT),
        (3, f"=SUMIFS('Sprint Plan'!$H${R0}:$H${R1},'Sprint Plan'!$C${R0}:$C${R1},$A{r})", NUM_DAY),
        (4, f"=SUMIFS('Sprint Plan'!$I${R0}:$I${R1},'Sprint Plan'!$C${R0}:$C${R1},$A{r})", NUM_PTS),
        (5, f"=SUMIFS('Sprint Plan'!$J${R0}:$J${R1},'Sprint Plan'!$C${R0}:$C${R1},$A{r})", NUM_PTS),
        (6, f"=SUMIFS('Sprint Plan'!$K${R0}:$K${R1},'Sprint Plan'!$C${R0}:$C${R1},$A{r})", NUM_PTS),
        (7, f"=SUMIFS('Sprint Plan'!$L${R0}:$L${R1},'Sprint Plan'!$C${R0}:$C${R1},$A{r})", NUM_PTS),
        (8, f'=IF($F{r}=0,"",$G{r}/$F{r})', NUM_PCT),
        (9, f"=$F{r}-$G{r}", NUM_PTS),
    ]
    for col, f, fmt in specs:
        c = db.cell(row=r, column=col, value=f); calc(c, col in (6, 7, 8)); c.number_format = fmt
        c.alignment = Alignment(horizontal="center")
db.cell(row=21, column=1, value="TOTAL").font = F(10, True)
db.cell(row=21, column=1).border = BOX
for col, fmt in (("B", NUM_INT), ("C", NUM_DAY), ("D", NUM_PTS), ("E", NUM_PTS),
                 ("F", NUM_PTS), ("G", NUM_PTS), ("I", NUM_PTS)):
    c = db[f"{col}21"]; c.value = f"=SUM({col}17:{col}20)"; calc(c, True); c.number_format = fmt
    c.fill = PatternFill("solid", fgColor=GREY); c.alignment = Alignment(horizontal="center")
c = db["H21"]; c.value = '=IF($F$21=0,"",$G$21/$F$21)'; calc(c, True); c.number_format = NUM_PCT
c.fill = PatternFill("solid", fgColor=GREY); c.alignment = Alignment(horizontal="center")
for rng in ("H17:H21",):
    db.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["1"],
                                  fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, size=10, bold=True, color="9C0006")))
    db.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["0.9", "1"],
                                  fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(name=FONT, size=10, color="9C6500")))
    db.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0.9"],
                                  fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONT, size=10, color="006100")))
db.cell(row=22, column=1, value=f"Sprint counts per quarter are uneven on a {WEEKS}-week cadence (5/4/4/4), because "
        "21 days do not divide a quarter.").font = F(9, color="404040", it=True)

section(db, 24, "Investment mix vs target")
header_row(db, 25, ["Workstream", "Planned (pts)", "Share of plan", "Target share", "Delta", "", "", "", ""])
for i, (name, asm_row) in enumerate([(PROD, 18), (PLAT, 19), (DEBT, 20), (KTLO, 21)]):
    r = 26 + i
    c0 = db.cell(row=r, column=1, value=name); c0.font = F(10, True); c0.border = BOX
    b = db.cell(row=r, column=2,
                value=f"=SUMIFS('Annual Plan'!$H$4:$H$23,'Annual Plan'!$C$4:$C$23,$A{r})")
    calc(b, True); b.number_format = NUM_PTS; b.alignment = Alignment(horizontal="center")
    c = db.cell(row=r, column=3, value=f'=IF($B$30=0,"",$B{r}/$B$30)'); calc(c)
    c.number_format = NUM_PCT; c.alignment = Alignment(horizontal="center")
    d = db.cell(row=r, column=4, value=f"='Assumptions'!$B${asm_row}"); link(d)
    d.number_format = NUM_PCT; d.alignment = Alignment(horizontal="center")
    e = db.cell(row=r, column=5, value=f'=IF($C{r}="","",$C{r}-$D{r})'); calc(e)
    e.number_format = NUM_PCT; e.alignment = Alignment(horizontal="center")
db.cell(row=30, column=1, value="TOTAL").font = F(10, True); db.cell(row=30, column=1).border = BOX
for col, fmt in (("B", NUM_PTS), ("C", NUM_PCT), ("D", NUM_PCT)):
    c = db[f"{col}30"]; c.value = f"=SUM({col}26:{col}29)"; calc(c, True); c.number_format = fmt
    c.fill = PatternFill("solid", fgColor=GREY); c.alignment = Alignment(horizontal="center")
db.cell(row=31, column=1, value="A workstream not matching its target is a conversation, not an error — but it should be a deliberate one.").font = F(9, color="404040", it=True)

section(db, 33, "Baseline reconciliation — capacity model vs last year's actuals")
REC = [
    (f"FY2026 actual average velocity (pts / {WEEKS}-week sprint)", FY26_VEL, NUM_PTS, True,
     f"INPUT. Placeholder from an illustrative FY2026, restated per {WEEKS}-week sprint. Replace with the mean of your "
     "own last 6-8 sprints — and if those were a different length, convert them first or this comparison is wrong."),
    (f"FY2026 average net dev-days per {WEEKS}-week sprint", FY26_DAYS, NUM_DAY, True,
     "INPUT. Placeholder. Same sprints as the line above, or the comparison is meaningless."),
    ("Implied FY2026 throughput (pts per focused dev-day)",
     '=IF(OR($B$34=0,\'Assumptions\'!$B$10=0),"",$B$34/($B$35*\'Assumptions\'!$B$10))', '0.00', False,
     "Last year's actual velocity divided by (dev-days x focus factor). Per dev-day, so cadence cancels out — this is "
     "the figure to compare across sprint lengths."),
    ("Throughput baseline used in this plan", "='Assumptions'!$B$11", '0.00', False,
     "The lever on 'Assumptions'. Below the implied figure = a conservative plan."),
    ("Delta vs history", '=IF($B$36="","",$B$37/$B$36-1)', NUM_PCT, False,
     "Negative is conservative. More than +10% needs a reason: new tooling, a smaller codebase, fewer dependencies."),
    ("Capacity-model baseline velocity (pts / sprint)", "='Assumptions'!$B$29", NUM_PTS, False,
     f"What the roster and calendar predict, averaged over the {NS} sprints."),
    ("Delta vs FY2026 actual velocity", '=IF($B$34=0,"",$B$39/$B$34-1)', NUM_PCT, False,
     "Two independent routes to the same number. Within about 10% of each other means the baseline holds up."),
]
r = 34
for label, val, fmt, is_input, note in REC:
    c0 = db.cell(row=r, column=1, value=label); c0.font = F(10, True); c0.border = BOX
    c = db.cell(row=r, column=2, value=val); c.number_format = fmt
    c.alignment = Alignment(horizontal="center")
    if is_input:
        inp(c, True)
    else:
        link(c) if isinstance(val, str) and "Assumptions" in val else calc(c, True)
        c.fill = PatternFill("solid", fgColor=GREY)
    n = db.cell(row=r, column=3, value=note); n.font = F(9, color="404040")
    n.alignment = Alignment(vertical="top", wrap_text=True)
    db.merge_cells(start_row=r, start_column=3, end_row=r, end_column=9)
    db.row_dimensions[r].height = max(15, 12 * (len(note) // 90 + 1))
    r += 1

add_dashboard_chart(db)

for s in wb.worksheets:
    s.sheet_view.showGridLines = False
wb.properties.title = f"FY2027 Yearly Sprint & Story-Point Plan ({WEEKS}-week sprints, baseline example)"
wb.properties.creator = "Claude Code"

wb.save(OUT)
print(f"saved {OUT} | {NS} x {WEEKS}-week sprints")
