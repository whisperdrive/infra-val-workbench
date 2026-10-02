# Infra Val Workbench

Recurring infrastructure valuations: last year's report, overlay and client models to this year's value.

Upload four files: last year's valuation report, last year's client model, last year's overlay (our valuation
workings, a separate workbook or sheets inside the client model) and this year's client model. The orchestrator
takes it from there:

1. **Reads the files.** The report is read from its text layer, with no model calls, and every table region on it is
   cropped to an image. The workbooks each get a row map, their external links, and their target and valuation date.
2. **Reads the key tables from their images.** A text layer can scramble a table (columns, merged headings) and has
   nothing for one pasted as a picture:
   - the tables most likely to hold the key figures, and the pictures beside them, are read from their images;
   - code checks each read against the page's own characters where there are any (each row's figures on a line under
     its own label, the headings in the page's order); otherwise a second model reads the image again, independently,
     and the two reads are compared cell by cell (a figure shifted into the next column, or swapped headings, is a
     difference);
   - a read that doesn't check out goes round the table loop (fix, check against the image, an arbiter).
3. **Extracts the report's key facts.** Only what the rebuild needs, each figure low, mid and high where the report
   gives a range (the mid as printed; on the page, the midpoint of the ends where it prints none, marked as such):
   - who and when;
   - the equity value;
   - the discount rate, terminal growth and franking credit utilisation (however the report puts it: a utilisation
     rate, gamma, "80% value ascribed to franking credits", 80% or 0.80);
   - where disclosed, the terminal value, the present values of the forecast and of the terminal value, and the value
     of franking credits;
   - how the terminal value is worked out, in the report's words, and any exit multiple. The report is searched for
     everything it says about its terminal value, and that goes to the models with the document; code then classes the
     method (none; an EV/EBITDA, EV/RAB or other exit multiple; growth on an average of the cash flows, on an adjusted
     or normalised cash flow, or on the final year's). The report's terminal value method fact decides first; its
     other sentences count only where they're about the terminal value, and a multiple the valuation implies (its
     EV/EBITDA cross-check) is never the method. Where there's no terminal value or it's an exit multiple, there's
     no growth rate to source, and the inputs card says so rather than "not found".

   One model extracts, code checks each fact against its page, a second model reviews, and they loop on what's open.
   The code check is the backstop behind the models: the quote must be on the cited page (not starting or ending
   inside a word or a number), and each figure must be in the quote in the scale and currency it's stated in (A$2.3m
   isn't A$2.3bn), under its own label rather than another fact's (in "WACC of 7.25% and terminal growth of 2.5%" the
   2.5% is the growth rate), and not inside a date or a longer number; a name is checked as text; a range needs both
   ends, in one scale.
   Then each fact with a figure is looked up, blind, on the image of where it sits (its table's, else its page's): the
   read gets the report's context with its figures and dates masked, so it can't take the answer from the text layer.
   Where the image reading differs, the reviewer looks at the image with both. A date is compared whole (day, month and
   year), not by its numbers. A correction stands when two reads of the image agree; it waives only the figure checks
   the text layer then fails, for that quote and those figures (where the quote is, and a table still to settle, must
   still check out). What doesn't settle goes to a person, with the image, and doesn't lead the rebuild meanwhile.
   Slides' own text and tables are exact, so they aren't looked at again.

   Every model that reads a table or judges a figure is also given what the report says around it (`context.py`): the
   text on its page, the transmittal letter (else the executive summary), the scope of the engagement, any definitions,
   and what's established so far (target, client, valuation date, units, the equity value's basis). It's context for
   what a figure means, not a source of figures.
4. **Works out which file is which.** It uses likeness, dates, links and where the report's figures sit, with a second
   opinion. It confirms the roles only when the evidence supports it (an overlay inside the client model must leave
   the client's sheets as the client's). An overlay set without its sheets, in a copy of the client model with the
   adviser's tabs behind a divider ("Adviser>>" up to "Client>>"), has those tabs as its own and the rest as its copy of
   the client model, where most of the rest are last year's client model's sheets. The adviser's own names for its
   divider tabs go in `VALUATION_DESK_OVERLAY_MARKERS` in `.env` (comma separated); the code names no firm.
5. **Rebuilds last year in Python.** The overlay's formulas are compiled to a module and checked cell by cell against
   Excel. The report's equity value is found in it and tied; the Python rebuild feeds it again from last year's client
   model. The inputs the report also states are sourced, not inferred (`sourced.py`): each is the cell the model's
   formulas read, found by following them, not by looking for the number, then recomputed from it and checked
   against the report. For each end:
   - the discount rate: the cell the discount factors read (the first and the last period's), the factors
     recomputed at it, the report's rate for that end (the low value at the higher rate). Where the factors blend
     two cells holding the rate (one for some revenue, one for the rest), both are sourced, and this year's rate goes
     on both. A discounting is a SUMPRODUCT of cash flows and factors (with 0 / 1 flag rows, a window of years, where
     it has them), an XNPV or NPV, or the SUM of a present-value row (or of present-value rows added up column by
     column, blocks of cash flows at one factor row: one discounting of them together, sourced from the parts'
     factors), or a present-value row summed up to a date typed in a cell (SUMIF(dates, "<"&end+1, present values)):
     that date is its cut-off, read from its cell on each feed, so a forecast end moved this year applies. A factor
     of exactly 1 on the valuation date's own period (no time to discount over) reads as nil there where no cash flow
     sits against it. A terminal value discounted on its own has one factor, which fits any date: its rate must be one a cell
     its factors read holds; where its factors follow a convention the app
     doesn't recompute (mid-year with a part-year stub at its midpoint, then whole-year steps, say), its rate and its
     valuation date are still the cells their formulas read, and the card says the factors aren't recomputed;
   - the terminal growth rate: the cell the terminal value's formula reads as g, the terminal value recomputed as
     X × (1 + g) / (r − g) at that end's own discount rate;
   - franking credit utilisation: the fraction every period's franking credits read; rerun at nil in Python, the
     equity value must fall by exactly the value of franking credits;

   the growth rate and the utilisation are also traced the other way (`forward.py`): from the cell holding the
   report's figure on a row labelled like it, up through every formula that reads it, to the equity value. That finds
   them wherever they sit, not only in the shapes the tracer down knows: a terminal value inside the last cash flow
   (=W8 + W8 × (1 + g) / (r − g)) or added after the discounting (=XNPV(…) + TV × the last factor), and franking
   credits added into the main cash flows (the first formula reading the utilisation is the franking used, its other
   operand the gross credits, discounted with the cash flows: their value is that part discounted on its own, and
   the nil rerun must fall by it). Where both ways find it, they must agree;
   - an exit multiple, where the report's terminal value is one: the cell the terminal value's formula reads as the
     multiple in multiple × metric, the terminal value recomputed from it, the metric what the report says it's a
     multiple of (EBITDA, the RAB), and the report's multiple for that end (the low value at the lower multiple). There's
     then no growth rate to source, and the card says so. It's traced up from the report's figure too, so a multiple
     inside the last cash flow (=W8 + W6 × multiple) or added after the discounting is found the same way.

   One that isn't sourced to a cell (typed into a formula, say), or doesn't check out, comes to you. What the
   report discloses of the value (the terminal value, the PV of the forecast and of the terminal value,
   the value of franking credits and its share) is reconciled to the same split of the overlay's discountings: a
   terminal value discounted on its own (one period, labelled so) is the terminal value; one inside the last cash
   flow is its term of that cell, discounted at that period's factor; one added after the discounting is the term
   that adds it (its present value); franking credits added into the cash flows are their part, discounted with
   them, and come off the forecast's present value. The split is in the
   units most of the report's figures tie in (discountings in thousands under an equity value in millions, at 100%
   under a share of it).
6. **Rolls forward onto this year's client model.** Rows are found by label, history, numbers and a trace, with row
   agents for the ones in doubt. The trace starts from the rows both models clearly share (the same label, last year's
   numbers: an output like the distributions or the cash flow available, an input like a volume or a tariff) and
   follows the dependency graph to the row, down from the outputs and up from the inputs, pairing each step with last
   year's by its label, its numbers or its words, else as the one row left once the others are paired. It follows
   formulas, not sheets, so a row moved to a new tab or with rows put in between is still reached; two ways agreeing
   count for more, and a trace that leads elsewhere than the label leaves the row in doubt. A row still in doubt comes to you with what it is: the heading it sits under and its neighbours,
   what its formula adds up or works out from, last year's figures, which overlay row reads it and whether that's a
   cash flow the value discounts; and the rows of this year's model that might be it, each with its figures for the
   same periods and a button to use it. Candidates that are the same series from last year's valuation date on are
   one answer, whichever is meant. A row that carries periods of its own, whatever its sheet's header says (an annual
   block under a quarterly header: period ends with a stub at either end, period starts, year numbers), is worked
   out, not hunted for: it moves by its own periods that ended between the two valuation dates. The valuation date
   moves wherever the discountings read it: each discounting's date is
   followed back to the cell it's typed in, through plain references and names (=Val_Date), and every one of those
   moves (a copy on a DCF sheet moves with its input, so cash flow dates counted from the input and discount periods
   counted from the copy stay in one column;
   a DCF sheet with a date of its own has it moved too), and so does the cell labelled as the valuation date (an overlay
   can keep one date for its discountings and another for the rest). Where this year's model forecasts to another date
   than last year's moved by the roll (a horizon half a year further on, rolled a year), the overlay's periods past its
   forecast are nil (not last year's figures standing in, nor this year's last period read again from a column past
   last year's timeline), and the overlay's own copy of the forecast's end follows this year's: a typed date labelled
   as an end that a terminal value is placed at, or the end a sum up to a date reads (SUMIF(dates, "<"&end+1, present
   values), the periods after it left to the terminal value), where last year's cash flows stopped on it. A row the discountings reach that isn't found this year, and whose figures don't move the
   value when nudged (a check, a label), doesn't hold the value back; the gate lists it. And each discounting's periods that end before the new date
   are cut off (its present-value row, else its factor row, else its cash flows, at nil): last year's overlay needed no
   cut-off of its own where last year's model had nothing before last year's date, but this year's model can have the
   year between the two dates filled in (a fixed horizon, a model dated before its valuation date), and a factor worked
   out from the date would compound those quarters into the value. A period's end is the one this year's feed works
   out for it where the overlay counts its periods from the valuation date (its columns are then this year's periods,
   with nothing to cut), else last year's moved by the sheet's periods. The period ending on the new date is cut off too,
   as a valuer zeroing the overlay's own period flags by hand does (keeping it is one of the methods below). This
   year's value is shown only where it can be trusted, and
   is held back otherwise, with a need that says why:
   - the rows its cash flows come from weren't found (or were found blank, or not surely);
   - a discounting still reads another date than this year's valuation date;
   - the overlay reads nothing of this year's client model (last year's figures, rolled by date alone);
   - this year's model adds a term the value takes in, or drops one last year's value took in, not yet confirmed
     (the terms, below);
   - the zero-roll check fails: at last year's valuation date, this year's model is far from last year's value;
   - the time check fails: the roll doesn't move a discounting on by its own rate. This year's model at last year's
     date (the periods to the new date cut off) against the roll at last year's rate, each discounting's present
     value of the same cash flows, matched by their amounts: rolled a year at 10%, it grows 10%. Where a discounting's
     cash flow dates and its discount periods don't move together (one counted from the valuation date's input, the
     other from a copy of it), it doesn't, and the value is wrong however plausible it looks. More than 1.5% a year off
     its rate holds the value back (`result.TIME_HOLD`); more than 0.5% is a point to check (`result.TIME_CHECK`).
     Where no discounting under the value is read exactly, the time isn't measured, and the page says so.

   A move of more than about 15% at last year's date (`overlay.ZERO_ROLL_CHECK`, a judgment call) doesn't hold the
   value but asks you to confirm it's the new forecast and not a row matched wrongly.

   **Terms added or gone** in the sums the overlay's rows sit in (the rows it reads, what they add up, and what adds
   them up): each this year is set against last year's, term by term, its terms paired where the row finder is sure
   of them, else by their label, numbers or words (a term swapped for another is one gone and one new; a term
   regrouped under a subtotal of the same sum is neither). A term this year's model has and last year's didn't, with
   figures after this year's valuation date, either feeds a row the overlay reads (a new cost under the cash flow it
   reads): the value takes it in and moves by it, so it's **held** until you confirm the term belongs; or it's beside
   them (the overlay reads the sum's other terms, not the sum): the value leaves it out, a point to check, not a hold.
   Likewise a term last year's model had and this year's doesn't, with figures after last year's valuation date: where
   last year's value took it in, this year's has lost it, **held** until you confirm it's gone; beside, a point to
   check. One the overlay read itself is a row to find, not a term gone (`result._term_changes`). A sum whose make-up
   changed (fewer than half of last year's terms pair: an annual total of a quarterly row, a model rebuilt) isn't
   compared: that's a restructure, and the rows to check say so.

   Two changes the rows the overlay reads can't show are points to check too, and don't hold the value either:
   - **cash-flow lines the overlay doesn't read**: on the client sheets it reads, a row of this year's model labelled
     as a cash flow to or from equity (an injection, a contribution, a distribution, a dividend, a capital return)
     with figures after this year's valuation date, that matches no row of last year's model (rowfind, the other way
     round) or one with nothing after last year's date, and that no row the overlay reads adds up. Last year's
     overlay had nothing of it to discount, so the roll leaves it out. Each is listed with its row, its periods and
     its total (and the model's own present value of it, where it has one);
   - **the scenario the client model is saved on**: the overlay reads the saved values, so a model saved on another
     case gives another value. The selectors (a typed scenario, case, sensitivity or switch its formulas read: a
     small whole number or an option's name) are shown on the Result page, this year's next to last year's (the
     prior model's, and the overlay's own copy where it has one), with when each model was saved (a cell of its
     own, else the file's properties), so you confirm the scenario the valuation should use. One that differs, or
     can't be matched, is a point to check (`scenarios.py`).

   A client model rebuilt between valuations is the hard case: a sheet of last year's figures pasted in under last
   year's labels (a reconciliation, say) matches last year's rows by label and by history, and the zero-roll check
   can't tell, since it reproduces last year's numbers exactly. So a pasted copy (typed values where last year's row
   was formulas, last year's figures in every period both have) is never a candidate for a row, nor the sheet last
   year's became, nor what tells the horizon. Where the two models share under half their line-item labels
   (`overlay.REBUILT`), the rows found other than by their labels are listed to check, though the value stands where
   its checks pass.
   Inputs typed into the overlay outside its discountings (a net debt, a cash balance, a declared distribution, an
   adjustment; not a units constant a name like "thousand" stands for) aren't fed by this year's model, so they stay at last year's figures: **held**, and shown so, highlighted
   on the Result page and in the workpaper until you set this year's. Each comes with a suggestion from this year's
   client model, checked first against last year's: the row of last year's model that holds last year's figure at
   last year's valuation date (its label agreeing) is the row read in this year's model at this year's date. A
   suggestion is only applied when you use it, or type your own figure; the bridge then has a step of its own for it.
7. **Bridges last year's value to this year's.** Report → rounding → rebuilt → time value → last year's cash flows paid
   → this year's forecast → discount rate → this year, for the low and the high; the mid is their average. The time
   value, the cash flows and the new forecast are at last year's discount rate; where you set this year's (the Result
   page's discount rate card: a range's two ends, the low end of the value at the higher rate), it goes on the cells
   each end's discountings read for it, and the discount rate step is the move to it (the zero-roll check stays at
   last year's rate: a new rate isn't a row matched wrongly). This year's value is also worked out other ways, the
   **methods** (`methods.py`), each against the default (the overlay's own formulas, rolled, the periods ending on or
   before the new date cut off): the period ending on the date kept in; the overlay's own forecast flags rolled in
   place of the cut-off (where it has any, the default's figure); each discounting recomputed in code as the overlay
   discounts (the default's figure, to the cent: the check that the others are like for like), mid-period, mid-year
   and on the other day count; and the mid at the midpoint rate rather than the average of the ends. The method you
   prefer is this year's value, with a bridge step of its own for the move from the default. Where a
   discounting under the value can't be read here, the roll-forward is one step, not its unwind landed in the new
   forecast. A chart compares the undiscounted forecast cash flows, last year's and this year's.
8. **Has gpt-sol review the run end to end** and says what looks implausible: once per result, like any other
   decision (a try again, or a rerun that works out the same result, keeps the points; a changed result is reviewed
   again, with the earlier points in view so the ones that still stand keep their titles).

**The conventions:**
- The primary approach is an **equity DCF**: cash flows to equity discounted at the cost of equity (an overlay that
  discounts free cash flows at a WACC and deducts net debt works the same way here).
- The conclusion is the **equity value**, low / mid / high, where the mid is the midpoint of the low and the high.
- **Ex-distribution** by default, unless the report is overwhelmingly cum-distribution or only the cum-distribution
  figure is in the model.
- This year's discount rate is **last year's** unless you set this year's on the Result page.

## The orchestrator

Plain code runs the stages in order (`engine/orchestrator.py`). gpt-sol decides only at a few bounded points:
- roles the rules and the second opinion don't agree on;
- cells that hold the equity value, where the pairing can't tell;
- the review at the end.

Each decision is checked in code before anything acts on it.

The **run log** is its memory. It records every start, outcome, decision and escalation, with the inputs each was made
on. Code, not the prompt, keeps it from going round in circles:
- A stage runs again only when its inputs change: a new file, a person's decision, a date. A stage that failed waits
  for that, or for a person's "try again".
- gpt-sol decides an issue once per set of inputs, sees what was tried before, then hands it to a person.
- A person's action changes the inputs, so the stages it affects run again by themselves.

## The page

A strip at the top of every page shows the seven stages on a track, each with what it found, and the mascot at the
one that matters. Beside it is the one thing that matters now: the files to upload, the decision that needs you (what
kind: a date to confirm, cells to pick, rows to find), how long is left and when the answer is due (from how long each
stage took last time, here or on other engagements), or the answer (last year's equity value to this year's, with
this year's low to high). A stage is a link to its card, and each thing that needs you lands on the card where it's
decided. An engagement opens where it matters: the decision that blocks it, else the Result once there is one, else
the Run page.

- **Run**: the four files as four slots (each file, how it was read, its role and who confirmed it, its valuation
  date where the roll-forward reads it), the valuation dates across the files, and, folded, changing the roles, the
  checks behind them, the engagement's profile and what happened. Upload files opens a window to drop or browse; a
  file dropped anywhere on the page opens it too.
- **Report**: what the report says (the equity value low / mid / high, the valuation date, the discount rate,
  terminal growth and franking utilisation, each with its page and image check) and how much was checked; a fact that
  needs you opens in full; every fact and table is folded below.
- **Rebuild**: the tie to the report and the model inputs sourced and checked; the reconciliation, the assumptions
  and the Python module folded, each with its result on the fold.
- **Result**: the value bridge (low, mid, high), this year's discount rate (last year's until you set it), the
  methods (and the one you prefer), the inputs held at last year's, the terms added or gone in this year's model
  (each to confirm where it moves the value), the cash-flow lines the overlay doesn't read, the
  client model's scenario next to last year's (and when each model was saved), the cash-flow chart, the review, and
  how the files link up (the map) folded. A review point names the years and the bridge step it's about (checked
  against the run), and a click marks them on the chart and the bridge. **Export workpaper** downloads it all as an Excel file: the summary (the
  equity value low / mid / high: the report's, rebuilt, this year's and the move; the inputs; the dates; the review),
  the bridge with a waterfall of the mid, the cash flows with their chart, the inputs with the cell each was sourced
  from and every check on it, the reconciliation, the key facts, the files and their roles, the review and the run
  log. Figures are values, not links back to the files; it carries a disclaimer that it's for the team to check.

**All engagements** is where the app opens (and the logo, the top of the engagement list, or the ⋯ menu come back to
it). It lists every engagement: where it is
(finished, needs you, running, waiting, no files yet), what's for you, and its equity value, last year per the report
to this year, the mid. A click opens one; tick two to compare them side by side: the equity value, the inputs and the
dates, the bridges step by step (the mid, the same steps matched; a dash where one hasn't a step) and this year's cash
flows. Differences aren't worked out between engagements in different units.

**Export diagnostics (anonymised)**, in the ⋯ menu, describes a run for diagnosis without a word of the client's:
counts, yes / no, ratios, dates and the app's own words. Each model's shape (its timelines by frequency: monthly,
quarterly, semi-annual, annual; its periods, first and last), how alike the two client models are, the profile, each
stage's outcome, the facts by status, the result's checks, the discountings traced, the roll, the reliability gate
(with how many cash-flow lines the overlay doesn't read, how many terms added and gone, in the value and held, and how
the row finder's trace did on the rows the value reads), each row to find by its shape, and the client models'
scenario selectors by how this year's compares with last year's, with when each was saved. Every string is checked
against the app's own vocabulary before it leaves; anything else is redacted and counted. It's for pasting from a machine with real files into a session that can't see them.

The trace's weights and thresholds (`rowfind.WEIGHTS`, `TRACE_*`, the 0.7 a trace needs to count on its own) were set on
synthetic models. The gate's `trace` counts in the export are how to check them on a real one: of the rows the value
reads, those found with no trace, by the trace alone, with the trace agreeing, and with the trace leading elsewhere,
how many are confident, and the trace's scores in bands. Many rows led elsewhere, or found by the trace alone and then
picked otherwise by a person, say its weight is too high; rows the trace agrees on still in doubt, too low.

Evidence is folded away, never removed. All engagements, New engagement, Models, the call log, the workpaper and
Delete are in the ⋯ menu; Delete asks for the engagement's name.

## Running it

```
uv sync
cp .env.example .env            # your Azure AI Foundry endpoints (git-ignored)
uv run python engine/llm.py     # sign in once (device code)
uv run uvicorn app.server:app --port 8003
```

Then open http://localhost:8003. **Delete this engagement**, in the ⋯ menu, removes an engagement and everything worked
out for it (its reports and key facts, the models only it uses, the Python overlay and your picks on it, the run
log and call log), so the same files can be uploaded and run again from the start; a model another engagement also
uses stays, and the report-reading rules learned so far stay. It waits until nothing of the engagement is running. Files stay on this machine, in `uploads/` and `out/`, both git-ignored. Model calls go
to your organisation's Azure AI Foundry and are logged per engagement; the spend chip in the header opens the log.
Your firm's logo goes in the git-ignored `brand/` folder, and the mascot shows without one.

## Tests

```
uv run python tests/make_pack.py        # a synthetic pack: fictional names and numbers, both overlay layouts
uv run python tests/check_workbench.py  # the whole run on it, every model call stubbed; the orchestrator's rules
uv run python tests/check_trace.py      # reading a DCF back from its factors (flags, loose conventions); the roll;
                                        # the growth rate and the utilisation traced up to the equity value
uv run python tests/check_xlruntime.py  # Excel functions in the Python runtime
```

## Layout

- `app/`: `server.py` (FastAPI), `index.html` (the page) and the mascot.
- `engine/`:
  - `orchestrator.py`: the stages, the run log and the decisions.
  - `workbench.py`: the store and the steps.
  - `keyfacts.py`: the report's key facts.
  - `visual.py`: the key tables read from their images, and each fact checked on its image.
  - `context.py`: what the report says around a figure (the page, the letter, the scope, definitions).
  - `result.py`: the tie, the bridge and the cash flows.
  - `sourced.py`: the discount rate, terminal growth and franking credit utilisation, sourced and checked.
  - `forward.py`: an assumption traced up to the equity value, through the formulas that read it.
  - `methods.py`: this year's value worked out other ways, against the default; the preferred one.
  - `scenarios.py`: the scenario each client model was saved on, this year's next to last year's, and when each was
    saved.
  - `workpaper.py`: the Excel workpaper, built in memory from the engagement's result.
  - The rest came from the Valuation Desk: the workbook library and row map, the roles, the Python overlay
    (`xlcompile.py`, `xlruntime.py`, `overlay.py`), the DCF tracer, row finding and the map.
- `web/charts.js`: the charts.
- `docs/report_rules.md`: the curated rules the report agents read. Learned lessons stay in `out/`.

Experimental: every figure must be checked by a qualified person before anyone relies on it.
