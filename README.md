# Infra Val Workbench

Recurring infrastructure valuations: last year's report, overlay and client models to this year's value.

Upload four files: last year's valuation report, last year's client model, last year's overlay (our valuation
workings, a separate workbook or sheets inside the client model) and this year's client model. The orchestrator
takes it from there:

1. **Reads the files.** The report is read from its text layer, with no model calls. The workbooks each get a row
   map, their external links, and their target and valuation date.
2. **Extracts the report's key facts.** Only what the rebuild needs:
   - who and when;
   - the equity value (low, high and the midpoint);
   - the discount rate, terminal growth and franking credit utilisation;
   - where disclosed, the terminal value, the present values of the forecast and of the terminal value, and the value
     of franking credits.

   One model extracts, code checks each fact against its page, a second model reviews, and they loop on what's open.
3. **Works out which file is which.** It uses likeness, dates, links and where the report's figures sit, with a second
   opinion. It confirms the roles only when the evidence supports it.
4. **Rebuilds last year in Python.** The overlay's formulas are compiled to a module and checked cell by cell against
   Excel. The report's equity value is found in it and tied; the Python rebuild feeds it again from last year's client
   model.
5. **Rolls forward onto this year's client model.** Rows are found by label, history and numbers, with row agents for
   the ones in doubt. This year's value is shown only where the rows its cash flows come from were found.
6. **Bridges last year's value to this year's.** Report → rounding → rebuilt → time value → last year's cash flows paid
   → this year's forecast → discount rate → this year, for the low and the high; the mid is their average. A chart
   compares the undiscounted forecast cash flows, last year's and this year's.
7. **Has gpt-sol review the run end to end** and says what looks implausible.

**The conventions:**
- The conclusion is the **equity value**, low / mid / high, where the mid is the midpoint of the low and the high.
- **Ex-distribution** by default, unless the report is overwhelmingly cum-distribution or only the cum-distribution
  figure is in the model.
- This year's discount rate is **last year's** unless a person sets another.

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

- **Workbench**: files, roles, the engagement's profile, where the run is, what needs you, and what the orchestrator
  did.
- **The report**: the key facts, their checks and the agents' decisions.
- **Rebuild**: the tie to the report and the model's assumptions under the equity value.
- **Result**: the value bridge and the cash-flow chart.
- **Map**: how the files link up and what changed in the client model, with the detail a click away.

## Running it

```
uv sync
cp .env.example .env            # your Azure AI Foundry endpoints (git-ignored)
uv run python engine/llm.py     # sign in once (device code)
uv run uvicorn app.server:app --port 8003
```

Then open http://localhost:8003. Files stay on this machine, in `uploads/` and `out/`, both git-ignored. Model calls go
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
  - `result.py`: the tie, the bridge and the cash flows.
  - The rest came from the Valuation Desk: the workbook library and row map, the roles, the Python overlay
    (`xlcompile.py`, `xlruntime.py`, `overlay.py`), the DCF tracer, row finding and the map.
- `web/charts.js`: the charts.
- `docs/report_rules.md`: the curated rules the report agents read. Learned lessons stay in `out/`.

Experimental: every figure must be checked by a qualified person before anyone relies on it.
