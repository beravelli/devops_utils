# Yearly sprint and story-point plan — baseline example

`FY2027-sprint-plan-baseline.xlsx` is a worked example of a yearly plan built on
capacity rather than wishful thinking. Every figure in it — people, holidays,
estimates, velocity — is an invented placeholder. Replace the inputs and the whole
workbook recalculates.

**Cadence: 17 three-week sprints**, Monday 4 January to Friday 24 December 2027
(17 × 21 = 357 days), leaving the last week of December outside the plan as
shutdown and slack. Each sprint carries 15 weekdays before company holidays.

## The baseline chain

```
roster -> effective FTE
       x working days (company holidays netted out)
       - PTO - on-call cover            = net dev-days
       x focus factor x throughput      = baseline velocity
       - unplanned-work reserve         = commitable story points
```

With the example roster (7 people, 5.90 effective FTE, 70% focus factor,
1.25 points per focused dev-day, 15% reserve): **55.4 points per sprint baseline,
799.9 commitable points across 17 sprints**, against 698.3 points of risk-adjusted
roadmap demand — an **87.3% annual load**, with every quarter between 86% and 90%.

## Tabs

| Tab | What it holds |
|---|---|
| Read Me | Colour legend, how to use it, and every assumption with its source |
| Assumptions | All planning levers: cadence, focus factor, throughput, reserve, target investment mix |
| Holidays | Company holiday dates, netted out of working days automatically |
| Sprint Calendar | 17 three-week sprints, dates, quarter, net working days |
| Team Capacity | Roster with FTE, allocation, ramp factor, PTO, plus a PTO reconciliation check |
| Sprint Plan | The baseline sprint by sprint: dev-days, commitable points, planned, load %, actuals, rolling velocity |
| Annual Plan | Roadmap epics: estimate / confidence = risk-adjusted size, spread over a sprint range |
| Stories | Fine-grained stories per epic, with reconciliation against epic estimates and sprint capacity |
| Dashboard | Supply vs demand by quarter, investment mix vs target, baseline reconciliation |

## Files, and the order to run them

| File | What it is |
|---|---|
| `FY2027-sprint-plan-baseline.xlsx` | The workbook. Generated, but meant to be filled in by hand afterwards. |
| `build_plan.py` | Rebuilds the workbook from scratch. Cadence, roster, holidays and epics are all in here. |
| `verify_plan.py` | Audits every formula and independently recomputes the model. Run after any change to `build_plan.py`. |
| `plan_lib.py` | Shared model recomputation and chart rebuild, used by the generator and by the story skill. |
| `stories-input.json` | The story breakdown. Source of truth for the `Stories` tab. |

**`build_plan.py` overwrites the workbook, which drops the `Stories` tab.** The order is:

```bash
python planning/build_plan.py                                                    # rebuild the plan
python planning/verify_plan.py                                                   # audit it
python .claude/skills/story-breakdown/stories.py write --input planning/stories-input.json
python .claude/skills/story-breakdown/stories.py check                           # reconcile
```

Hand edits made in Excel — PTO tweaks, actual points completed, notes — live only in
the `.xlsx`, so `build_plan.py` discards them. Either keep such edits out of the
generator's way by not regenerating, or fold them back into `build_plan.py` first.
`stories.py write` is safe: it replaces only the `Stories` tab and leaves the rest
of the workbook, including hand edits, alone.

## Breaking epics into stories

`.claude/skills/story-breakdown` is a Claude Code skill: ask for a story breakdown
in a session opened on this repo and it reads the epics, splits them into 1/2/3/5/8
point stories with Given/When/Then acceptance criteria, assigns them to sprints
inside each epic's range, and reconciles the totals.

The Q1 epics are broken down as a worked example — 35 stories, 160 points, each epic
reconciling exactly to its estimate. The other 12 epics are left for the skill to do.

Story points reconcile against an epic's **estimate**, not its risk-adjusted size,
so assigned story points per sprint sit well below the plan's own Load %: the gap is
the risk adjustment, which is deliberately not broken down.

## Changing the cadence

Sprint length drives the calendar. Set **Sprint length** on `Assumptions` (weekdays
per sprint is a formula off it, and every date on `Sprint Calendar` follows), set
**Sprints in plan**, then add or delete rows on `Sprint Calendar` and `Sprint Plan`
so both hold that many rows, and re-point the start/end sprint numbers on
`Annual Plan`.

Throughput is expressed **per dev-day**, so it survives a change of cadence.
Velocity per sprint does not — it scales with sprint length, which is why the
baseline reconciliation states last year's actuals per sprint on the same cadence.

A longer sprint creates no capacity. Annual dev-days barely move; what changes is
granularity: fewer, larger commitments, a coarser feedback loop, and less whipsaw
from a single holiday, since each absence is a smaller share of a longer sprint —
against which each commitment is larger, so a mis-estimate costs more before you
find out.

Two side effects of three-week sprints worth knowing:

- **Quarters split 5/4/4/4**, because 21 days do not divide a quarter. Q1 gets an
  extra sprint and roughly 57 more commitable points than the other three.
- **Annual commitable capacity drops slightly** — 17 × 15 = 255 working days
  against 260 for 26 × 10, and on-call cover scales with weekdays. The same
  roadmap that ran at 92.1% on a 2-week cadence needed three estimates trimmed
  (INI-06, INI-09, INI-10) to come back to 87.3%.

## Two things to know

**Six sprints show Load % over 100%, deliberately.** Every quarter sits at 86–90%,
but epic points spread evenly while capacity does not: S03, S07, S08, S11, S12 and
S16 each lose one or two working days to holidays. Level the plan by moving scope
into S05, S09, S13, S14 and S17, the quarter-boundary sprints sitting at 49–71%
because the next quarter's epics have not started. Levelling that is the actual
exercise of yearly planning, so the example leaves it visible.

**The Dashboard reconciles two independent routes to the same baseline** — the
capacity model (55.4 pts/sprint) against last year's actual velocity (57.0
placeholder), a −2.9% delta. If those disagree by much more than 10%, the focus
factor or the throughput rate is wrong. That check is the part that keeps a
baseline honest.

Story points are team-relative, so the 1.25 points per focused dev-day is
meaningless until it is re-derived from your own last 6–8 sprints.

## Conventions inside the workbook

Blue text on yellow = key baseline lever. Blue text = ordinary input. Black =
formula, do not overwrite. Green = formula pulling from another tab. Grey = totals.

The file holds live formulas with no cached values, so Excel, Google Sheets,
LibreOffice or Numbers will compute them on open; a quick-preview tool may show
blank cells until then.
