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
   - code checks each read against the page's own characters where there are any; otherwise a second model reads the
     image again, independently, and the two reads are compared;
   - a read that doesn't check out goes round the table loop (fix, check against the image, an arbiter).
3. **Extracts the report's key facts.** Only what the rebuild needs:
   - who and when;
   - the equity value (low, high and the midpoint);
   - the discount rate, terminal growth and franking credit utilisation;
   - where disclosed, the terminal value, the present values of the forecast and of the terminal value, and the value
     of franking credits.

   One model extracts, code checks each fact against its page, a second model reviews, and they loop on what's open.
   Then each fact with a figure is looked up, blind, on the image of where it sits (its table's, else its page's); where
   the image reading differs, the reviewer looks at the image with both. A date is compared whole (day, month and
   year), not by its numbers. A correction stands when two reads of the
   image agree; what doesn't settle goes to a person, with the image. Slides' own text and tables are exact, so they
   aren't looked at again.

   Every model that reads a table or judges a figure is also given what the report says around it (`context.py`): the
   text on its page, the transmittal letter (else the executive summary), the scope of the engagement, any definitions,
   and what's established so far (target, client, valuation date, units, the equity value's basis). It's context for
   what a figure means, not a source of figures.
4. **Works out which file is which.** It uses likeness, dates, links and where the report's figures sit, with a second
   opinion. It confirms the roles only when the evidence supports it.
5. **Rebuilds last year in Python.** The overlay's formulas are compiled to a module and checked cell by cell against
   Excel. The report's equity value is found in it and tied; the Python rebuild feeds it again from last year's client
   model. The inputs the report also states are sourced, not inferred (`sourced.py`): each is the cell the model's
   formulas read, found by following them, not by looking for the number, then recomputed from it and checked
   against the report. For each end:
   - the discount rate: the cell the discount factors read (the first and the last period's), the factors
     recomputed at it, the report's rate for that end (the low value at the higher rate);
   - the terminal growth rate: the cell the terminal value's formula reads as g, the terminal value recomputed as
     X × (1 + g) / (r − g) at that end's own discount rate;
   - franking credit utilisation: the fraction every period's franking credits read; rerun at nil in Python, the
     equity value must fall by exactly the value of franking credits.

   One that isn't sourced to a cell (typed into a formula, say), or doesn't check out, comes to you. What the
   report discloses of the value (the terminal value, the PV of the forecast and of the terminal value,
   the value of franking credits and its share) is reconciled to the same split of the overlay's discountings.
6. **Rolls forward onto this year's client model.** Rows are found by label, history and numbers, with row agents for
   the ones in doubt. This year's value is shown only where the rows its cash flows come from were found.
7. **Bridges last year's value to this year's.** Report → rounding → rebuilt → time value → last year's cash flows paid
   → this year's forecast → discount rate → this year, for the low and the high; the mid is their average. A chart
   compares the undiscounted forecast cash flows, last year's and this year's.
8. **Has gpt-sol review the run end to end** and says what looks implausible.

**The conventions:**
- The conclusion is the **equity value**, low / mid / high, where the mid is the midpoint of the low and the high.
- **Ex-distribution** by default, unless the report is overwhelmingly cum-distribution or only the cum-distribution
  figure is in the model.
- This year's discount rate is **last year's** (there's no control yet to set another).

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
decided. Once a run is finished the app opens on the Result.

- **Run**: the four files as four slots (each file, how it was read, its role and who confirmed it, its valuation
  date where the roll-forward reads it), the valuation dates across the files, and, folded, changing the roles, the
  checks behind them, the engagement's profile and what happened. Upload files opens a window to drop or browse; a
  file dropped anywhere on the page opens it too.
- **Report**: what the report says (the equity value low / mid / high, the valuation date, the discount rate,
  terminal growth and franking utilisation, each with its page and image check) and how much was checked; a fact that
  needs you opens in full; every fact and table is folded below.
- **Rebuild**: the tie to the report and the model inputs sourced and checked; the reconciliation, the assumptions
  and the Python module folded, each with its result on the fold.
- **Result**: the value bridge (low, mid, high) and the cash-flow chart, the review, and how the files link up (the
  map) folded. A review point names the years and the bridge step it's about (checked against the run), and a click
  marks them on the chart and the bridge. **Export workpaper** downloads it all as an Excel file: the summary (the
  equity value low / mid / high: the report's, rebuilt, this year's and the move; the inputs; the dates; the review),
  the bridge with a waterfall of the mid, the cash flows with their chart, the inputs with the cell each was sourced
  from and every check on it, the reconciliation, the key facts, the files and their roles, the review and the run
  log. Figures are values, not links back to the files; it carries a disclaimer that it's for the team to check.

Evidence is folded away, never removed. New engagement, Models, the call log, the workpaper and Delete are in the ⋯
menu; Delete asks for the engagement's name.

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
uv run python tests/check_trace.py      # reading a DCF back from its factors
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
  - `workpaper.py`: the Excel workpaper, built in memory from the engagement's result.
  - The rest came from the Valuation Desk: the workbook library and row map, the roles, the Python overlay
    (`xlcompile.py`, `xlruntime.py`, `overlay.py`), the DCF tracer, row finding and the map.
- `web/charts.js`: the charts.
- `docs/report_rules.md`: the curated rules the report agents read. Learned lessons stay in `out/`.

Experimental: every figure must be checked by a qualified person before anyone relies on it.
