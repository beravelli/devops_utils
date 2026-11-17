# Yearly sprint and story-point plan — baseline example

`FY2027-sprint-plan-baseline.xlsx` is a worked example of a yearly plan built on
capacity rather than wishful thinking. Every figure in it — people, holidays,
estimates, velocity — is an invented placeholder. Replace the inputs and the whole
workbook recalculates.

## The baseline chain

```
roster -> effective FTE
       x working days (company holidays netted out)
       - PTO - on-call cover            = net dev-days
       x focus factor x throughput      = baseline velocity
       - unplanned-work reserve         = commitable story points
```

With the example roster (7 people, 5.90 effective FTE, 70% focus factor,
1.25 points per focused dev-day, 15% reserve): **36.9 points per sprint baseline,
814.5 commitable points across 26 sprints**, against 750.2 points of risk-adjusted
roadmap demand — a 92% annual load.

## Tabs

| Tab | What it holds |
|---|---|
| Read Me | Colour legend, how to use it, and every assumption with its source |
| Assumptions | All planning levers: cadence, focus factor, throughput, reserve, target investment mix |
| Holidays | Company holiday dates, netted out of working days automatically |
| Sprint Calendar | 26 two-week sprints, dates, quarter, net working days |
| Team Capacity | Roster with FTE, allocation, ramp factor, PTO, plus a PTO reconciliation check |
| Sprint Plan | The baseline sprint by sprint: dev-days, commitable points, planned, load %, actuals, rolling velocity |
| Annual Plan | Roadmap epics: estimate / confidence = risk-adjusted size, spread over a sprint range |
| Dashboard | Supply vs demand by quarter, investment mix vs target, baseline reconciliation |

## Two things to know

**Eight sprints show Load % over 100%, deliberately.** Every quarter sits at
89–94%, but epic points spread evenly while capacity does not: holiday-short and
PTO-heavy sprints go over, while quarter-boundary sprints (S07, S13, S20, S21) sit
under 65% because the next quarter's epics have not started. Levelling that is the
actual exercise of yearly planning, so the example leaves it visible.

**The Dashboard reconciles two independent routes to the same baseline** — the
capacity model (36.9 pts/sprint) against last year's actual velocity (38.0
placeholder), a −3% delta. If those disagree by much more than 10%, the focus
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
