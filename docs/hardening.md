# Hardening: where a wrong value could get through, and what checks it

A review of the whole workflow (October 2026) by four reviewers, one per stage: the report's facts, the overlay's
rebuild and discounting, the roll-forward and its bridge, and the roles, gates, chart and page. Each looked for ways a
wrong number, or the wrong basis, reaches this year's value without a check catching it. Findings were reproduced on
synthetic workbooks before anything was changed. Below: what is now checked, then what's left, ranked, for a person to
pick.

A **hold** keeps this year's value back until it's settled. A person can acknowledge a hold with a reason: the value
then goes through, the check stays listed with the reason, and the workpaper records it. If the figures it found
change, it holds again. A **point to check** is listed but doesn't keep the value back.

## Now checked

### This year's cash flows against last year's (cashflows.py)
Every discounting under the value, period by period, on last year's model, rolled to the new date, and on this year's
model at last year's rate.

| Check | When | |
|---|---|---|
| Stale cash flows | this year's discounted cash flows are last year's in every period both have, while the client rows behind them were revised | hold |
| Unchanged model | the same, with the client rows unchanged too | point to check |
| Zero roll exactly 1 | this year's model at last year's date gives last year's value to 4 decimals while the rows were revised | hold |
| The new-forecast step, split | revisions to the periods both years have, periods added and dropped, the terminal value, and what moves outside the discountings; more than 0.5% of last year's value outside: point to check; more than 2%: hold | hold / check |
| Step against the cash flows | the step goes the other way to the cash flows' change | point to check |
| Terminal value nil | the terminal value reads nil this year | hold |
| A longer forecast | this year's model forecasts past the overlay's last period, where last year's didn't | hold |
| A shorter forecast | this year's forecast ends before the overlay's last column | point to check |
| Stand-ins | last year's figures stand in for periods this year's model doesn't have | hold |

### How each discounting is built
Each discounting's function (XNPV, NPV, SUMPRODUCT, present values summed), convention, and whether the app cuts it
off and time-checks it on the roll, on the Model assumptions card and the cash-flow card.

| Check | When | |
|---|---|---|
| Time | the roll moves each discounting on by its own rate: now also those whose factors don't fit a convention the app recomputes (typed period counters were a year out with nothing flagged) | hold beyond 1.5% a year |
| Uncut past | a discounting the app doesn't cut off still discounts a period ending on or before the new date | hold |
| XNPV's date | an XNPV counting from another date than this year's | hold |
| XNPV on the date | an XNPV taking the cash flow on the date in whole | point to check |
| Not time-checked | a discounting whose time can't be measured | point to check |
| A period's rate | quarterly or half-yearly factors: the period's rate is annualised before its cell is looked for (a quarter's rate was taken as a year's, and this year's annual rate written onto it) | fixed |
| This year's rate | not applied where last year's at that end doesn't check out against the report | refused |

### The interest valued, the basis, the equity cell
| Check | When | |
|---|---|---|
| Interest | the report's interest valued (a new fact) against the share the overlay applies: a share row in a SUMPRODUCT, or an interest cell multiplying after the discounting; different, or applied where the report doesn't say | hold |
| Interest unapplied | the report states a share below 100% and none is applied | point to check |
| A partial formula | a discounting that's only part of its cell's formula (×share) is recomputed with the rest of the formula (the methods gave the 100% figure) | fixed |
| Basis | the overlay cell's own basis, from its label or formula (a distribution deducted: ex; added: cum), against the report's; the report silent and the cell cum | hold |
| Basis wording | before / after, including / excluding, pre- / post- the distribution | read |
| Cum this year | a cum value keeps the period ending on this year's date, with its own bridge step | default |
| The equity cell | rows labelled as an enterprise value, a sensitivity, a scenario or a case aren't the equity value; a row needs its label or its formula to say it is; two rows as likely go to a pick, not the listing order; a low above the high isn't a pair; a range printed high first is read the right way round | fixed |

### The gates
| Check | When | |
|---|---|---|
| The tie | the overlay's value doesn't round to the report's (it was a point to check, with the difference called rounding) | hold |
| The rebuild | last year rebuilt in Python against what Excel saved: over 0.1% apart | hold (point to check above 0.0001%) |
| Assumed roll | the months to roll assumed, not read from the dates | hold |
| A check that crashed | the terms comparison failing; the new-lines comparison failing | hold / point to check |
| Another linked workbook | this year's value reads cells of a workbook linked other than the client model, at last year's values | hold |
| Your roles | roles a person confirmed now pass the same checks as the orchestrator's | hold |
| INDIRECT | a cell building its address from text is never held at Excel's value as "safe" | fixed |
| A typed terminal base | the terminal value grows from a figure typed in the overlay | point to check |
| A renamed sheet | two of this year's sheets as like last year's: neither taken; the choice no longer depends on the run | fixed |
| An earlier result | the result shown while the bridge runs again, or after a run that failed, is marked; the workpaper waits | fixed |
| Rules' version | results worked out by older rules are worked out again | fixed |

### The page
- The cash-flow chart's years are four digits and in order (a forecast past 2049 put its 2050s first).
- The Prefer button: a tooltip, a confirmation naming the move, a "recomputing" state, who preferred it and when,
  back to the default or the previous choice; the method step sits before this year's total everywhere.
- The models side by side: one table of each model's specifications (the report, the overlay as saved, the rebuild,
  both client models, this year), last year's flagged where they don't agree.

### Row finding (measured)
On synthetic pairs of models with known answers (`tests/variants.py`: revised, renamed, a downside case inserted, the
sheet renamed, a prior-forecast block, reordered, a row dropped, and combinations), 44 rows: before, 35 right and 9
wrong (every row of an inserted downside case, confidently; the prior-forecast copy taken by its numbers); now 44
right, none wrong. The tools: address, label, block and heading, lineage two steps out, kind, numbers as a band; the
row agents by meaning, then the models with the row's card, lineage, role in the valuation, trail and notes; pick
cards found again in a corrected model; a person's picks logged, anonymised, for calibration.

## Left to do, ranked

High: a wrong value can still reach the result unflagged.

1. ~~Row agents settle "by the numbers" on the candidate nearest last year's~~ — done: by meaning, numbers a band.
2. ~~Same-label matching when a block is inserted, or a case column~~ — done: block copies by heading; case columns.
3. ~~Row picks keyed by address, not by file~~ — done: each pick keeps a card and is found again (held inputs still by
   address).
4. **A mid-period valuation date** keeps the whole straddling period (half a year's cash flow already earned). Point
   to check with the elapsed fraction; a proration method.
5. **A distribution declared at this year's date in the client model** that the overlay doesn't deduct.
6. **Fixed or rolling decided for all sheets at once**: a minority sheet reads the wrong dates.

Medium: flagged weakly, or only by the reviewer.

7. A report figure printed to few digits ties several cells (A$2.3bn ties anything from 2,250 to 2,349): require a
   unique tie, and the more precise statement elsewhere in the report.
8. A range in the report with only the midpoint in the overlay becomes one figure.
9. gpt-sol's and a person's equity-cell picks are barely checked, and outlive a replaced overlay.
10. The image check confirms a figure without comparing its column (a cum, 100% or mid column).
11. The valuation date falls back to the overlay's without a check, and takes a letter's date.
12. Circular references aren't iterated on this year's feed.
13. Factor conventions widened (whole-period counters, 30/360, (1+r/k)^(k·t)) so fewer discountings are "loose".
14. The methods' alternative conventions are wrong for first periods, stubs, quarters and a terminal value in the row.
15. Roles: auto-confirmation counts restated checks; upload order picks this year's model between undated versions;
    identical file names defeat the second opinion; old confirmations survive a blocked rerun.
16. The facts gate passes with nothing approved; the review can't block, and its input is cut mid-JSON; a crash in the
    scenario or reconciliation check raises no need; editing an approved fact doesn't rerun the checks on it.

Low.

17. The chart shows this year's series while the value is held; a terminal value inside the last cash flow inflates
    the last bar.
18. NPV over a range with blanks counts the blank columns; a row of cash flows in billions can be read as factors.
19. Typed per-period overlay rows inside the cash flows slide a period on a rolling horizon.
