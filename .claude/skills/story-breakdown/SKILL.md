---
name: story-breakdown
description: Break the Annual Plan's epics into fine-grained, estimated user stories inside planning/FY2027-sprint-plan-baseline.xlsx, then reconcile them against the epic estimates and per-sprint capacity. Use when asked to create stories, split or break down epics, generate a backlog, groom or refine the roadmap into stories, write acceptance criteria for planned work, or size and assign stories to sprints.
---

# Story breakdown

Turns described epics into stories you could actually pull into a sprint, and proves
they add up. The workbook stays the single source of truth: epics on `Annual Plan`,
stories on a generated `Stories` tab, with live formulas tying one to the other.

**The JSON input file is the source of truth for stories, not the tab.** `write`
deletes and recreates `Stories`, so edits typed into that tab are lost on the next
run. Point edits belong in the JSON; capacity, PTO and actuals belong in the other
tabs, which `write` never touches.

## Prerequisites

Epics on `Annual Plan` need at least: ID, name, workstream, estimate, confidence,
and a start/end sprint. Epics missing a sprint range cannot have stories assigned.

Python needs `openpyxl`. If it is missing, run `bootstrap.sh` from this directory —
it builds a `.venv` at the repo root and falls back to fetching wheels with `curl`
when `pip` cannot resolve `files.pythonhosted.org`. Use `.venv/bin/python` for every
command below if the system Python lacks openpyxl.

## Workflow

### 1. Read the context

```
python .claude/skills/story-breakdown/stories.py epics
```

Returns every epic (estimate, confidence, risk-adjusted size, sprint range, points
per sprint) and every sprint (quarter, working days, commitable points, what the
epic spread already implies). Work from these numbers — never from memory of the
plan, and never from the tab's formula cells, which hold no cached values until
Excel or Sheets opens the file.

### 2. Split each epic into stories

Every story is **1, 2, 3, 5 or 8 points**. Nothing above 8: that is an epic wearing
a story's name. Story count therefore follows the epic's size — a 20-point epic is
about 4 stories, a 70-point one about 13. Past roughly 15 stories, say so: the epic
is really two or three, and splitting it on `Annual Plan` will plan better than a
15-story block spanning a whole quarter.

Split **vertically**, so each story delivers something observable. "Build the API
endpoint" and "build the UI for it" are two halves of one story, not two stories.

Patterns worth reaching for, in rough order of usefulness:

| Pattern | Split on | Example |
|---|---|---|
| Workflow steps | stages of the user's journey | browse → add to cart → pay → confirm |
| Business rules | one rule per story | flat-rate shipping, then tiered, then free over £50 |
| Data variations | the shapes of input | single-currency, then multi-currency |
| Interfaces | each surface | web checkout, then mobile, then partner API |
| Happy / unhappy path | success, then each failure mode | payment accepted, declined, timed out |
| CRUD | the verbs | create, then edit, then delete, then list |
| Spike | the unknown, time-boxed | 3 points to prove the migration path, then re-estimate |
| Operational slice | the non-functional cut | works at 10 rps, then at 1000 |

Check each story against INVEST: Independent, Negotiable, Valuable, Estimable,
Small, Testable. A story failing *Valuable* is usually a task; fold it into one.

### 3. Write the story and its acceptance criteria

`story`: "As a <role>, I want <capability> so that <outcome>." Name a real role from
the domain, not "user", when the epic implies one.

`acceptance`: one or more Given / When / Then clauses, separated by a blank line.
These are the test cases; if you cannot write them, the story is not understood yet.
Include the unhappy path where it exists.

### 4. Assign sprints

Stay inside the epic's `start_sprint`–`end_sprint` range. Load each sprint toward,
but not past, its `commitable_pts`, and front-load the stories other work depends
on. Sprints running light in the plan (the quarter-boundary ones) are where extra
stories belong.

### 5. Write the JSON and build the tab

```json
{
  "stories": [
    {
      "epic": "INI-01",
      "story": "As a returning shopper, I want my saved address pre-filled at checkout so that I can pay without retyping it.",
      "acceptance": "Given a signed-in shopper with a saved address\nWhen they open checkout\nThen the address fields are pre-filled and editable\n\nGiven a shopper with no saved address\nWhen they open checkout\nThen the fields are empty and no error is shown",
      "points": 3,
      "sprint": 1,
      "status": "Ready",
      "notes": ""
    }
  ]
}
```

Story IDs are generated (`INI-01-S01`, `-S02`, …) in Annual Plan order, then sprint
order — do not supply them. Keep the file at `planning/stories-input.json` unless
asked otherwise, then:

```
python .claude/skills/story-breakdown/stories.py write --input planning/stories-input.json
```

This rebuilds `Stories` and re-creates the Dashboard chart, which openpyxl drops on
any load/save round-trip.

### 6. Reconcile

```
python .claude/skills/story-breakdown/stories.py check
```

Story points reconcile against the epic's **estimate**, not its risk-adjusted size:
risk adjustment covers what a breakdown cannot see, so counting it twice would
inflate the backlog. Tolerance is ±10%.

Expect assigned story points per sprint to sit **well under** the plan's own Load %
— stories total the epic *estimates*, while `Sprint Plan` column L spreads the
*risk-adjusted* figures. A sprint at 95% on the plan and 65% in stories is not a
contradiction: the gap is the risk adjustment, which is deliberately not broken down.

`check` fails on: a story over 8 points, a story with no points, a sprint outside
its epic's range, an unknown epic. It warns on: variance outside tolerance, missing
acceptance criteria, an unassigned story, an epic with no stories, a sprint whose
assigned stories exceed its commitable points.

When an epic's stories come out **more than 10% above the estimate**, the breakdown
found scope the estimate missed. Do not quietly shrink the stories to fit. Report it
and say what it costs: a bigger estimate on `Annual Plan` raises that quarter's load,
and the Dashboard will show whether it still fits. Changing an estimate is the
user's call, not the skill's.

## Reporting back

Say how many stories per epic, the variance per epic, which sprints are now at or
over capacity, and anything `check` flagged. Name the epics whose estimates look
wrong in light of the breakdown — that is the most valuable thing this exercise
produces.
