"""Audit the workbook: structural formula checks + an independent recomputation of the model.

Run after any change to build_plan.py. Exits non-zero if a formula is malformed, references an
unknown sheet, uses a function LibreOffice cannot evaluate, or if the model stops reconciling.
"""
import datetime as dt, re, sys
from openpyxl import load_workbook

from pathlib import Path
P = str(Path(__file__).resolve().parent / "FY2027-sprint-plan-baseline.xlsx")
wb = load_workbook(P)
names = set(wb.sheetnames)
print("sheets:", wb.sheetnames)

BANNED_RE = re.compile(r"(?<![A-Z_.])(XLOOKUP|XMATCH|SORT|FILTER|UNIQUE|SEQUENCE|TEXTJOIN|IFS|SWITCH|MAXIFS|MINIFS|CONCAT)\s*\(")
problems, formula_count = [], 0
ref_re = re.compile(r"'([^']+)'!")
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            formula_count += 1
            if v.count("(") != v.count(")"):
                problems.append(f"{ws.title}!{c.coordinate}: unbalanced parens -> {v}")
            if v.count('"') % 2:
                problems.append(f"{ws.title}!{c.coordinate}: odd number of quotes -> {v}")
            m = BANNED_RE.search(v.upper())
            if m:
                problems.append(f"{ws.title}!{c.coordinate}: banned/prefix-needed function {m.group(1)} -> {v}")
            for sheet in ref_re.findall(v):
                if sheet not in names:
                    problems.append(f"{ws.title}!{c.coordinate}: unknown sheet '{sheet}' -> {v}")
            # bare (unquoted) cross-sheet ref to a name containing a space
            for sheet in names:
                if " " in sheet and (sheet + "!") in v and ("'" + sheet + "'!") not in v:
                    problems.append(f"{ws.title}!{c.coordinate}: unquoted sheet name -> {v}")
print(f"formulas: {formula_count}")

# ---------- independent recomputation ----------
asm, cal, sp, ap, db, cap = (wb[n] for n in ("Assumptions","Sprint Calendar","Sprint Plan","Annual Plan","Dashboard","Team Capacity"))
focus, thru, resv, oncall = asm["B10"].value, asm["B11"].value, asm["B12"].value, asm["B13"].value
WEEKS = asm["B6"].value
N = asm["B7"].value
wkdays, s1 = WEEKS*5, asm["B8"].value
STEP, ENDOFF = 7*WEEKS, 7*WEEKS-3
efte = sum((cap.cell(row=r, column=3).value or 0)*(cap.cell(row=r, column=4).value or 0)*(cap.cell(row=r, column=5).value or 0) for r in range(4, 13))
hols = [cal_ for cal_ in (wb["Holidays"].cell(row=r, column=1).value for r in range(5, 26)) if cal_]
hols = [h.date() if isinstance(h, dt.datetime) else h for h in hols]
pto = [sp.cell(row=4+i, column=6).value for i in range(N)]
inits = [(ap.cell(row=r, column=3).value, ap.cell(row=r, column=6).value, ap.cell(row=r, column=7).value,
          ap.cell(row=r, column=10).value, ap.cell(row=r, column=11).value) for r in range(4, 24)]
inits = [(w, e, c, j, k) for (w, e, c, j, k) in inits if e]

print(f"cadence: {WEEKS}-week x {N} sprints, {wkdays} weekdays each")
print(f"effective FTE: {efte:.2f}  focus {focus}  throughput {thru}  reserve {resv}  on-call {oncall}  PTO total {sum(pto)}")
start = s1.date() if isinstance(s1, dt.datetime) else s1
assert start.weekday() == 0, f"sprint 1 does not start on a Monday: {start}"

rows, qtot = [], {}
for i in range(N):
    st = start + dt.timedelta(days=STEP*i)
    en = st + dt.timedelta(days=ENDOFF)
    assert en.weekday() == 4, f"sprint {i+1} does not end on a Friday"
    nh = sum(1 for h in hols if st <= h <= en)
    wd = wkdays - nh
    gross = efte*wd
    net = gross - pto[i] - oncall*wd
    base = net*focus*thru
    comm = base*(1-resv)
    planned = 0.0
    for w, est, conf, j, k in inits:
        if j <= i+1 <= k:
            planned += (est/conf)/(k-j+1)
    q = f"Q{(st.month-1)//3+1}"
    rows.append((i+1, st, en, q, nh, wd, net, base, comm, planned))
    a = qtot.setdefault(q, [0, 0.0, 0.0, 0.0, 0.0])
    a[0] += 1; a[1] += net; a[2] += base; a[3] += comm; a[4] += planned

print(f"\nlast sprint ends {rows[-1][2]}")
print(f"\n{'S':>3} {'start':>10} {'end':>10} {'Q':>3} {'hol':>3} {'wd':>3} {'net':>6} {'base':>6} {'comm':>6} {'plan':>6} {'load':>6}")
for n, st, en, q, nh, wd, net, base, comm, planned in rows:
    print(f"{n:>3} {st} {en} {q:>3} {nh:>3} {wd:>3} {net:>6.1f} {base:>6.1f} {comm:>6.1f} {planned:>6.1f} {planned/comm:>6.1%}")

tb, tc, tp = sum(r[7] for r in rows), sum(r[8] for r in rows), sum(r[9] for r in rows)
demand = sum(est/conf for _, est, conf, _, _ in inits)
print(f"\nannual: baseline {tb:.1f}  commitable {tc:.1f}  demand(from epics) {demand:.1f}  "
      f"planned(spread over sprints) {tp:.1f}  load {demand/tc:.1%}")
print(f"avg baseline velocity/sprint {tb/N:.1f}  avg net dev-days/sprint {sum(r[6] for r in rows)/N:.1f}")
if abs(tp - demand) > 0.01:
    problems.append(f"spread planned points {tp:.2f} != epic demand {demand:.2f} (an epic falls outside sprints 1-26)")
for q, a in sorted(qtot.items()):
    print(f"{q}: sprints {a[0]}  net dev-days {a[1]:.1f}  baseline {a[2]:.1f}  commitable {a[3]:.1f}  planned {a[4]:.1f}  load {a[4]/a[3]:.1%}")
mix = {}
for w, est, conf, _, _ in inits:
    mix[w] = mix.get(w, 0) + est/conf
for w, v in mix.items():
    print(f"mix {w}: {v:.1f} pts = {v/demand:.1%}")
fy26v, fy26d = db["B34"].value, db["B35"].value
implied = fy26v/(fy26d*focus)
print(f"\nbaseline reconciliation: implied FY26 throughput {implied:.2f} vs plan {thru} "
      f"(delta {thru/implied-1:+.1%}); capacity model {tb/N:.1f} vs FY26 actual {fy26v} (delta {(tb/N)/fy26v-1:+.1%})")
assert abs(sum(asm.cell(row=r, column=2).value for r in range(18, 22)) - 1.0) < 1e-9, "investment mix != 100%"

print("\n--- problems ---")
print("\n".join(problems) if problems else "none")
sys.exit(1 if problems else 0)
