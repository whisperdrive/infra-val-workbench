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
| Rules' version | results worked out by older rules are worked out again; a newer row finder makes the row agents look again (their older picks were set aside with nothing rerunning them, so the rows waited for a person) | fixed |
| Your held figures | a figure set for a held input, where the overlay changed under it: found again by its label and last year's figure, or set aside, never applied to another input | point to check |

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
2. ~~Same-label matching when a block is inserted~~ — done (S1 below): a label in more places this year than last, or
   in copies of a block, is settled only by a heading naming last year's case, or by what the model's own valuation
   reads; else it's left in doubt. Case columns: done.
3. ~~Row picks keyed by address, not by file~~ — done: where its label is in copies, a card also needs its place in its
   block and its figures to match. Held figures: done.
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

## Second review, 3 October 2026

Five reviewers, one per part (intake and roles; last year's rebuild and inputs; the roll-forward and row finding; the
result and its outputs; cross-cutting robustness), each reproducing on synthetic workbooks. Ranked: a wrong value
through on ordinary inputs, then after an internal failure, then the rest. "Re-run" marks the ones re-run outside the
reviewer's own session.

### A wrong value through on ordinary inputs, unflagged

- ~~**S1. A copy of a block settled on the wrong copy**~~ — fixed, but for one case. A label in more places this year
  than last (or in copies of a block) is settled by a heading naming last year's case, surely and well ahead (case
  words decide; words every copy shares never do; a heading naming another case rules a copy out), else by the copy
  whose rows reach the row of the model's own valuation last year's row reached (the equity value, the NPV), else it's
  left in doubt for the models or a person; the agents' first pass leaves it too; a pick card in copies needs its place
  and figures. Measured: the review's copy cases right; copies nothing tells apart left open; the pack's 134 rows
  unchanged; the reproduction (a downside 0.93× above the base) now held, not 10.6% low. **Left** (in
  `variants.KNOWN_WRONG`): copies too small to be seen as copies of a block (fewer than three rows, a line and its
  total), the same number of them both years, swapped this year, with no heading and no valuation row of the model's
  own: still matched by their order, whatever their figures. Closing it means doubting every label repeated as often
  both years where nothing tells the repeats apart, which in real models could leave many rows to pick. Was (re-run): A downside or P90 case inserted above the base, the
  base's heading renamed ("Management forecast" shares "case" with "Downside case"), no headings, or the base moved to
  another sheet: every row settled on the copy at confidence 1.0, identity, place and role all agreeing; this year's
  value 10.6% low, the zero roll inside the silent band. `structure.which_copy` takes any heading sharing a word;
  `rowfind._copies` lets lineage rescue the first occurrence (identical in every copy); `in_place` is set on a row
  whose block has copies; a pick card's `matches()` ignores figures and position. `tests/variants.py` had none of
  these cases.
- ~~**S2. A balance at the valuation date on a fixed horizon**~~ — fixed: a client row the value reads in one column
  only, last year's valuation date's, is a balance at the valuation date (`overlay.balance_cells`); this year it's read
  at this year's date, whether the overlay reads it plainly (the app moves it) or by its own date; the split counts the
  move as the balances' (the right roll is no longer held, the wrong one no longer passes); a point to check lists
  each move (a point to check where the app moved it, a note where the periods or the overlay's own date did, as they
  did before); no column at this year's date holds (so a three-month roll on an annual model now holds, where it
  deducted last year's balance). Measured: 110.7, both ways of reading it. Since: a balance read up to a year before
  last year's date (the latest actuals) is read as far before this year's, and one after it is a point to check; a
  cash-flow row read once more at the date by a cell of its own has only that read moved (the distribution declared at
  the date); a balance labelled as at a fixed date (a financial close, a completion) stays at it, a note; and a person
  can keep any balance at its own date, or read a kept one at this year's (the cash-flow card, `balances.json`). Was
  (re-run): Net debt read as `=Client!D11` (last year's
  date's column) still deducts last year's 200, not this year's 180: 18% off, no hold. The same overlay written with
  INDEX/MATCH rolls right and is held by `cf-split`, which measures movement: the stale read passes, the right one is
  held.
- ~~**S3. Equity ends typed in**~~ — fixed: the pairing takes only cells that work the value out; where the report's
  figures are found only typed in, the value is held saying so (it's worked out from other cells, the typed ones named);
  for an engagement paired before (its rebuild doesn't pair again on an update), a typed end holds; an end left at
  last year's to the cent while this year's model changed holds; the overlay's mid no longer the ends' midpoint where
  it was last year holds; the equity cells aren't listed as inputs held at last year's. Was (re-run): A low and high pasted from a sensitivity run (or data-table cells, which load
  as constants) tie to the report and pass as reliable; this year's value is last year's to the cent, with only two
  "held at last year's" notes, which also push the real held inputs off the list.
- ~~**S4. The doctor's holds applied without a person**~~ — fixed: a formula returning a reference (a top-level
  OFFSET, an IF returning one) has its target recorded as read, so the doctor follows it; what changes between the
  years now includes every cell the roll sets (the dates it moves, this year's rate, the inputs a person set, the
  forecast's end, the cut-off periods); holding the cells it finds safe is a person's decision on the Rebuild page
  (a point to check says so), never the orchestrator's; and a held cell whose inputs move on this year's feed holds
  the value. Was (code re-read; reviewer reproduced): A top-level OFFSET, or an
  IF returning one, records only its arguments as reads; the doctor calls it safe and the orchestrator holds it at
  Excel's value: a terminal base frozen at last year's, this year's value about 20% low, no hold.
- ~~**S5. A shared workbook's date set in another engagement**~~ — fixed: a workbook's valuation date and who
  confirmed it are each engagement's own (`file_dates.json`, through the store, seeded from the library's record on
  first read so an update changes nothing); a person's correction or the agents' confirmation goes to that engagement
  only, and the agents' earlier check of the same file and last year's date is reused, per engagement. The library
  keeps only what the file says. Was (re-run): Workbooks are shared by content; a date
  corrected in one engagement re-rolls every other engagement using the file: 3,384.7 became 3,542.4 with no hold,
  its date still shown as confirmed by the agents.
- ~~**S6. Decisions lost or corrupted**~~ — fixed: every decision file through `store.py` (one writer at a time, written
  whole, a damaged one moved aside and the value held until acknowledged); the agents keep their earlier picks.
  Was (re-run): The decision files (acks, terms, rate, method, row picks, held) are
  written in place without locks and read as empty when damaged: eight acknowledgements at once lost 151 of 160; a
  damaged rate file drops this year's rate with no flag; two quick row picks while the overlay thread is busy keep
  one; the rows job drops the agents' earlier picks on every rerun.
- ~~**S7. Decisions outlive their inputs**~~ — fixed: every check's key covers the engagement's files (its roles'
  workbooks and report), so an acknowledgement made before a client model was replaced lapses, said beside the check
  with its reason; a confirmed term keeps the file it was confirmed on and its label, and applies only while both
  still hold. Was: Acknowledgements of `cf-stale`, `cf-exact`, `cf-tv-nil`, `other-link` and
  `roll-assumed` are keyed without this year's figures, and survive a replaced client model; a confirmed term is keyed
  by row address, so another line at that row enters the value "confirmed by you".
- ~~**S8. A person's edits to approved facts**~~ — fixed: the rebuild's and the result's fingerprints carry a person's
  edit (its range's ends, value, unit, basis), and only where there is one, so an update reruns nothing for unedited
  facts; the roles' key is the figures they place files by (valuation, date, conclusion), by value, not every fact's
  approval or page, and an update recognises the older roles key without running anything. Was (code re-read): The rebuild's fingerprint carries only an edit's value
  text; the result's none of it: a growth rate or a rate range edited reruns nothing, and the page calls the result
  current.
- ~~**S9. The rebuilt-model fallback**~~ — fixed: where the row it would take has its label in more than one place
  this year (copies: a low case beside a central one), it's taken only as the one copy reaching the model's own
  valuation row last year's row reached; else it's left for the models or a person (the review's case now open).
  A rebuilt model's unique rows are taken as before, listed to confirm. Was: the fallback took the only row within the
  numbers band: a low case beside a central case revised 20% up.

### After an internal failure

- ~~**S10. Cum becomes ex**~~ — fixed: the basis's method kept as asked, a cum value without it held; the sort that
  could raise, by date then column. Was (code re-read): If the methods inventory raises, `asked` falls back to the person's
  preference, dropping the basis's on-date method: a cum value published on the ex basis, no method need. A natural
  trigger: `methods.py` sorting `(date, None)` flags.
- ~~**S11.**~~ — fixed: `cf-error` and `interest-error` are holds a person can acknowledge. Was: **the cash-flow checks
  crashing** collapse into one `cf-error` point to check, removing `cf-split`; an interest
  check failing raises nothing.
- ~~**S12.**~~ — fixed by `store.py`: moved aside, the value held. Was: **a damaged `equity.json` or `holds.json`** freezes the engagement silently (the loop swallows the error).
- ~~**S13.**~~ — fixed: a point to check with a try again. Was: **Azure down** leaves the rows "done"; nothing reruns
  them when it's back.

### Medium

- ~~The agents confirm this year's date on expectations alone (a year on, the financial year's end), not the file~~ —
  fixed: their confirmation needs a sign from the file itself; expectations alone leave it for a person.
- ~~The facts come from every report uploaded, not only last year's report's role; on a tie, a draft's lead~~ — fixed:
  once a report is placed as last year's, the facts navigated by are its own.
- ~~A report's folder is shared across engagements and deleted with either~~ — fixed: kept while another engagement's
  report reads it. **Left**: the same report uploaded again is read again into that folder, rewriting the same page
  and table images while the other engagement may be reading one (the same bytes; a read at that instant could fail
  and be retried).
- ~~A SUMPRODUCT's factor row is laid on the columns by position: a blank factor shifts every later one~~ — fixed: each
  factor read in its own column.
- A merged row ("Operating costs and tax") settled on its words alone, one kind of evidence agreeing (in
  `variants.KNOWN_WRONG`).
- ~~The overview and compare views show an earlier result's value while it's being worked out again~~ — fixed: the list
  shows no value for this year, saying why; the side-by-side view says which one is an earlier result.
- ~~Circular references on this year's feed take last year's saved value; an unknown function in a branch not taken
  last year falls through IFERROR silently~~ — fixed: either met on the way to this year's value holds it (`circular`,
  `unknown-fn`), saying where; a person can acknowledge it.
- ~~Every fact edit reruns the roles' second opinion (cost)~~ — fixed with S8.
- ~~Model calls don't pass `store=False`~~ — fixed: every call asks Azure to keep nothing (untested against Azure
  itself: if a deployment refuses the parameter, remove the one line in `engine/llm.py`).
- A single figure off the timeline stands in at last year's where its case column's heading is reworded (reviewer).

### The review of the fixes

Four reviewers went over S1 to S13 as fixed, reproducing on synthetic workbooks. Fixed since:

- Copies (S1, S9): pasted copies, rows of another shape or kind and mirrors aren't candidates; a row holding last
  year's forecast exactly beside revised ones can't be picked; a heading exactly on one copy decides; the periods
  (annual, not the quarterly copy) and the sheet decide where the count is unchanged; a pick other than the row found
  must pass the numbers band; a pick card is kept only where the label's occurrence and count, or its history figures,
  match.
- Balances (S2): only cells outside the overlay's periods are balances; one the discountings read moves but isn't
  counted twice in the split; fixed dates by named dates only; each moved balance against last year's forecast of it
  (10% off a point to check, 50% held); last year's figure stands in, said, where this year's model has no column.
- Holds added: a distribution declared at this year's date the overlay doesn't deduct; a discounting turning the other
  way; a mid-period valuation date (inferred by months: a leap year no longer looks like one); a report figure printed
  to few digits; the overlay's own sheet differing in this year's model; a person's method not worked out.
- Decisions: row picks re-read when their file changes; one session build at a time per engagement; the AI's equity
  pick and the agents' date never written over a person's; damaged files listed at every exit; Windows waits up to
  about 15 s with temporary files cleaned; no decisions for engagements that don't exist; the page's buttons rest
  during an action; two settings made at once both land.

### Low
~~A job the loop queued just as its engagement was deleted ran on, writing its log for an engagement that's gone~~ —
fixed: a job for a deleted engagement doesn't run, and nothing is written for one. Folders deleted with errors ignored; charts and fonts from public CDNs; the diagnostics' version pattern too broad;
low ≤ high not checked after the roll; `time` and `new-terms` holds shown as "a note"; the review prompt's wording.

### To watch on the first real run

- The circular-reference and unknown-function holds (`circular`, `unknown-fn`) are quiet on the synthetic packs; how
  often they fire on real models isn't known. Only the overlay's formulas are worked out (the client models' cells come
  in as the values Excel saved), so a loop in a client model (debt sizing, interest on cash) can't trip them; but an
  overlay with an iterative loop of its own holds the value in every engagement using it until someone acknowledges it.
- The rules added for the shapes above, each with a side that could fire falsely on a real model:
  - a pasted copy (`pasted`): a row typed this year that genuinely equals last year's in three or more periods (a fixed
    schedule a client retyped) holds until acknowledged;
  - an input that moves with the client model (`assumption-moved`): an input linked to the client model on purpose
    (the growth rate as the client's CPI) is a point every year it moves;
  - a lookup's table: an exact XLOOKUP, VLOOKUP or INDEX/MATCH whose pick the saved values don't settle keeps every
    cell of its table as an input held at last year's (the noise this was meant to remove);
  - the equity cells: a row reading another row holding the report's figures gives way to it; a presentation row
    that adds something real to the row it reads, yet still rounds to the report's figure, would be passed over;
  - a figure right of the timeline: read from this year's column where the two columns' headings agree (both blank
    counts); two blank headings over different things would map in silence;
  - the row maps built before these: their models have no record of formulas saved with no result, of how they were
    calculated, or the whole labels: those checks appear only for models read (or rebuilt) after this version.

### What would make it harder to get wrong
1. ~~One store for decisions~~ — done (S6, S7): locked, atomic, a damaged file a blocking need; decisions keyed on the
   files they were made on, lapsing when one is replaced.
2. ~~Copies a doubt by default~~ — done (S1): settled only where the heading, the valuation row, the periods or the
   figures decide; cards compared on occurrence and figures too.
3. A stale-read check beside `cf-split`: each client read outside the discountings that moves the value, against this
   year's model at the new date (the balances' check against last year's forecast covers part of it).
4. A check that can't run is a hold, section by section (done for the cash flows, the interest and the methods);
   post-conditions on the result (low ≤ high after the roll, the bridge's total) still to do.
5. ~~Judgements per engagement~~ — done (S5): the library holds only what's true of a file.
6. ~~Fingerprints of what each job read~~ — done (S8), with a test editing each field a person can edit.
7. ~~The doctor's holds a person's decision; runtime health checked on this year's feed~~ — done (S4; `circular`,
   `unknown-fn`).
8. ~~The failing variants in `tests/variants.py`~~ — done: 99 of 136 rows right, 32 left open where nothing in the
   model decides, the 5 wrong all in the 2 kinds listed in `variants.KNOWN_WRONG` (small copies swapped with close
   figures; a merged row), each taken off the list when it's fixed.

## Reading the models: issue shapes tested, 3 October 2026

Each shape built on synthetic workbooks (the pack's, changed at the XML level as Excel or another tool saves a file,
or small workbooks of its own) and run end to end with the models stubbed; the rule each time: the right cells and
value, or the value held saying why, never a different figure in silence. `tests/check_ingest.py` keeps the cases.

**Files.** An encrypted package (a compound file), a document renamed as a workbook and a Strict Open XML workbook are
refused, saying what to do; protected sheets and a locked structure read as normal (real encryption with a password
isn't exercised, only its signature). About 68,000 cells a second: a million in 15 s and 360 MB, so 2.3 million in about
35 s and 700 MB; files are read one at a time. The same bytes under two names, or in two engagements, are one file
(the library keeps the first name it saw). Macro-enabled workbooks read as normal; macros never run. Formulas saved
with no result (a workbook saved without being calculated) are recorded as such: one the value reads this year holds
it, naming the cells and saying to calculate and save; last year's model's or the overlay's are a point to check
beside the rebuild's own check; a calculation that didn't finish is a point to check; manual calculation and data
tables read as saved. Every cell is read, however small the size a sheet declares.

**Structure.** Hidden and very hidden sheets, and hidden rows, are found and used (their state isn't shown on the
page). The overlay as a separate linked workbook, as the adviser's tabs behind a divider in a copy of the client model
(a divider is recognised once the adviser's names are set; one in a style the names don't cover isn't, though the
roles came out right without it), as one added sheet, or as sheets with plain names: roles and values right; a valuation date alone no longer files the client's inputs sheet as the overlay's (it froze a
client-linked assumption at last year's in that form, so the same overlay gave two values). Sheets renamed and rows
moved between years are found again. Odd sheet names (spaces, brackets, an apostrophe, leading digits, "<END>") and an
add-in's sheets read as normal; hidden sheets with no formulas no longer make this year's model look rebuilt. A row the
value reads typed this year where last year's model worked it out, equal to last year's in place or one period off,
holds as pasted (one period off passed every check before). Forty thousand names, a third broken, some external: the
same value, about twice as slow as the pack (10 s); a broken name or a deleted reference under the value holds, naming
the cell where the error starts. A link to a file that wasn't uploaded leaves last year's model empty, saying so; a
renamed link target is matched by its values. A copy's client sheets are last year's client model's.

**Layout.** Monthly, quarterly and annual sheets, a timeline row far down, 285 quarterly columns and a terminal value
far right: right. A client figure right of last year's timeline (the model's own terminal value, a total) is read as
far right of this year's timeline, else held (last year's was read in silence). Periods dated by their first day
against periods dated by their last are matched. Static input sheets, labels and units in other columns, decorated,
prefixed and misspelt labels, merged cells and array formulas: right. **Left**: a sheet with two timelines (a
quarterly block above an annual one) dates every row by one of them; such a roll is held by the time and rows checks,
not passed.

**Key datapoints.** Valuation dates: every date in a labelled row is a candidate (this year's and last year's side by
side, in either order), an implausible one (a concession's end, a small number formatted as a date) last, the later of
equally good ones; time-only and log rows are never taken. The equity value: a row reading another candidate's cells (a
report table, a 100% or scaled copy, an adjustment of nothing) gives way to the row it reads; two rows as likely, or
only some of the report's figures in the overlay, ask with the candidates and never switch the basis (both had valued
cum-distribution, +25.0, with only points to check); whole labels keep "- Base case" and "(ex-distribution)" in view,
and "Less:" reads as deducting. Discount rate decoys and per-year rates: right (**left**: a rate typed as text or as a
whole number of percent isn't sourced; a point to check, the value ties). Terminal growth: as text or a whole number of
percent, sourced; one the report states that nothing under the value reads (an exit multiple instead) is a point to
check. A terminal value added after the discounting, in the enterprise value's formula or a cell of its own, splits
with the discountings (it held as "moving outside" them) and its inputs aren't inputs held at last year's. Franking: a
switch of 1, a tax rate labelled as franking and a CHOOSE of two gammas no longer stop the utilisation being sourced;
checks across thousands and millions agree. A CHOOSE on a typed selector is traced, sourced and recomputed as the case
it picks (an inactive case's rate was sourced, depending on the run). An input of the valuation's own that reads the
client model (a growth rate that is the client's CPI) is a point to check when it moves.

**Formulas.** OFFSET, XLOOKUP, XNPV and XIRR match Excel; INDIRECT, FILTER and LET hold the value (`unknown-fn`). Array
SUM(cash flows × factors), SUM(range × 1) and SUM over a fixed OFFSET read as the discountings they are (each held
falsely). A lookup's table cells it doesn't pick aren't inputs held at last year's. An iterative loop under the value
holds it. Excel's errors where the value doesn't read them change nothing; under it, the value holds. Units starting
with # stay text. Dates with a time of day are that day; serials before March 1900 count as Excel counts them;
durations are numbers. 1E9 is a billion, not cell E9 (it pulled the high end's discountings under the low end), and
LOG10( is a function.
