"""The few key facts the workbench needs from last year's valuation report, checked and reviewed.

The brief is short on purpose (KEYS): who and when (target, client, valuation date, units), the equity value
(low / high, and the midpoint, ex- or cum-distribution), the discount rate, the terminal growth rate and the
franking credit utilisation, and, where the report discloses them, the terminal value, the present values of the
forecast and of the terminal value, and the value of franking credits and its share of the equity value.
Four passes, each visible on the page:
  1. extract   a model reads the report's text layer (docingest.py, no model calls) and returns each fact with the
               page and a verbatim quote
  2. check     code, no model: the quote must be on the cited page, every value must be in the quote, the
               number must match the text, and a quote from a table that hasn't been verified is marked
  3. review    a second model call sees the same pages and the facts, and accepts, corrects or rejects each
               one with a reason, and lists key facts that were missed (corrections are checked like 2)
  4. resolve   the review and remediation loop: for every fact still open (failed checks, a correction or
               rejection, a fact the reviewer added) the extractor revises, keeps or withdraws it with a reason,
               code re-checks, and the reviewer accepts or objects again; up to MAX_ROUNDS. Facts both agree on are
               "agreed" (a person still approves); what they can't settle is escalated with both positions
Every prompt carries the rules learned from earlier loops (lessons.py), and the loop's episodes feed new ones.
conclusion() then picks, in code, the equity value the bridge starts from (headline).
    uv run python engine/keyfacts.py out/docs/<dir>/document.md [model]
"""
import json
import re
import sys

import lessons
from docingest import numbers

MAX_CHARS = 150_000  # longer reports: send the pages most likely to hold conclusions and assumptions
MAX_ROUNDS = 3       # review and remediation rounds before a fact goes to a person
CHECK_VERSION = 5    # bump when check() changes: existing facts are checked again once (2: spacing-tolerant, waivers;
                     # 3: numbers worked out by code, none for identity text or a range without a preferred point;
                     # 4: a date checked whole, its month included; 5: a figure's scale and currency, the label it
                     # sits under, a quote on word and number boundaries, identity checked as text, a range's ends,
                     # a waiver tied to the quote and figures it was given for)
LOOP_CHARS = 60_000  # the loop sends only the pages the open facts cite, and their neighbours
KEYWORDS = re.compile(r"valuation|discount|wacc|terminal|growth|multiple|rab|conclu|range|preferred|assumption|"
                      r"enterprise value|equity value|methodolog|approach|summary|cost of capital|cpi|inflation", re.I)

KEYS = """identity:   target_name (the entity or asset valued), client (who engaged the valuer), valuation_date,
            currency_units (e.g. A$m)
conclusion: equity_value: the concluded equity value. low_text / high_text for the range; value_text for the
            preferred value or midpoint if the report prints one, else "". basis: ex-distribution or
            cum-distribution (or ex- / cum-dividend) as the report states it, else "". If the report gives the
            equity value on both bases, extract the one it leads with as equity_value and the other as
            equity_value_cum (cum-distribution) or equity_value_ex (ex-distribution), each with its own quote.
            Where the report states them: terminal_value, pv_forecast (present value of the discrete forecast
            cash flows), pv_terminal_value (present value of the terminal value), franking_credits_value (the
            value of franking credits in the equity value), franking_credits_share (that value as a % of the
            equity value)
assumption: discount_rate (low_text / high_text for a range; basis e.g. post-tax nominal WACC, cost of equity),
            terminal_growth_rate, franking_utilisation (the franking credit utilisation rate, or gamma)
Nothing else: no other keys."""
KNOWN = {"target_name", "client", "valuation_date", "currency_units", "equity_value", "equity_value_cum",
         "equity_value_ex", "terminal_value", "pv_forecast", "pv_terminal_value", "franking_credits_value",
         "franking_credits_share", "discount_rate", "terminal_growth_rate", "franking_utilisation"}
CRITICAL = ("valuation_date", "equity_value", "discount_rate")  # the bridge can't start without them

FIELDS = """- value_text: exactly as printed ("7.25%", "A$2,296.7m", "30 June 2025"); low_text / high_text for a range,
  else "". value: the number in value_text (7.25 for 7.25%, 2296.7 for A$2,296.7m, 20250630 for a date as
  YYYYMMDD) or null for text. unit: "%", "x", "date", "years", "text" or the currency units ("A$m").
- basis: what the figure is on (e.g. "post-tax nominal WACC", "preferred", "real"), else "".
- page: the N of the nearest "<!-- page N -->" marker above the text you used.
- quote: copied verbatim from the document, the shortest sentence or table row that states the value
  (a table row as its cells separated by spaces, without the | characters). Never paraphrase."""

EXTRACT_PROMPT = """You are reading last year's final valuation report for a recurring infrastructure valuation.
A few of its figures are the reference used to rebuild last year's valuation in Python and roll it forward, so
extract exactly the datapoints below that the report states, with these keys, and nothing else:
{keys}

Rules:
{fields}
- Only facts the document states. If the report gives a figure more than once, use the main statement
  (the executive summary or the assumptions table).

Rules learned from earlier reviews (apply where relevant):
{rules}

Document:
{doc}"""

REVIEW_PROMPT = """You are the reviewer. Another model extracted the facts below from last year's valuation report
(the document follows). Code has already checked that each quote appears on the cited page and contains the
value; those results are shown. For every fact decide:
- accept: the value, basis and page are right and it is the report's main statement of it
- correct: something is wrong; give the corrected value_text / low_text / high_text / basis / page / quote
  (quote verbatim from the document)
- reject: the document doesn't support it, or it isn't a key fact (say why)
Then list datapoints of the brief below that were missed, with the same fields (only these keys; reject a fact
whose key isn't in the brief). Identity is wanted (the target and the client) even though it isn't a figure:
don't reject a fact only because it isn't a number.
{keys} Be strict: numbers must match
the document exactly, including units and whether a value is pre- or post-tax, nominal or real.
A quote must be one continuous piece of the document, so for a figure in a table the quote is its row and the
column it sits under goes in basis: check the column against the table, but don't "correct" a fact only to add
the table's headings to its quote.

Rules learned from earlier reviews (apply where relevant):
{rules}

Facts:
{facts}

Document:
{doc}"""

FIX_PROMPT = """You extracted key facts from last year's valuation report. A reviewer and automatic checks raised the
issues below. For each one decide:
- revise: the fact needs changing; give it in full, with the quote copied verbatim from the pages below
- keep: it is right as it stands; say why, pointing to the text
- withdraw: it shouldn't be in the reference (the report doesn't support it, or it duplicates another fact); every
  datapoint in the brief below is wanted, identity included (target, client); a key outside the brief isn't
A fact the reviewer added becomes yours if you keep or revise it; withdraw it if you disagree. The automatic
checks need the quote to be one continuous piece of the cited page that contains the value, and the number to
match the text. The brief:
{keys}
Field conventions:
{fields}

Rules learned from earlier reviews (list the IDs you apply in rules_applied):
{rules}

Issues:
{issues}

Pages:
{doc}"""

VERIFY_PROMPT = """You are the reviewer of key facts extracted from last year's valuation report. You, or the automatic
checks, raised the issues below and the extractor has answered each one. For each item decide:
- accept: the fact as it now stands is right and a key fact (for a withdrawal: it should indeed go)
- object: it is still wrong or shouldn't go; say why, and give the corrected value_text / low_text / high_text /
  basis / page / quote (quote verbatim) if you can, else leave them ""
Be strict about numbers, units, basis (pre or post tax, nominal or real) and the page. The automatic checks on
the fact as it now stands are shown: don't accept a fact whose checks fail unless you say why the check is wrong.
The reference wants every datapoint below that the report states, identity included (target, client), so a
withdrawal is right only when the report doesn't support the fact, duplicates another, or isn't in the brief:
{keys}

Rules learned from earlier reviews (list the IDs you apply in rules_applied):
{rules}

Items:
{items}

Pages:
{doc}"""

_S, _N, _I = {"type": "string"}, {"type": ["number", "null"]}, {"type": ["integer", "null"]}
CATEGORIES = ["identity", "conclusion", "assumption"]
_FACT = {"category": {"type": "string", "enum": CATEGORIES},
         "key": _S, "label": _S, "value_text": _S, "low_text": _S, "high_text": _S, "value": _N, "unit": _S,
         "basis": _S, "page": _I, "quote": _S}
EXTRACT_SCHEMA = {"type": "json_schema", "name": "report_facts", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["facts", "notes"],
    "properties": {"facts": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                        "required": list(_FACT), "properties": _FACT}},
                   "notes": _S}}}
_REV = {"id": {"type": "integer"}, "verdict": {"type": "string", "enum": ["accept", "correct", "reject"]},
        "reason": _S, "value_text": _S, "low_text": _S, "high_text": _S, "basis": _S, "page": _I, "quote": _S}
_MISS = {**_FACT, "why": _S}
_IDS = {"type": "array", "items": _S}
_FIXR = {"id": {"type": "integer"}, "action": {"type": "string", "enum": ["revise", "keep", "withdraw"]}, "reason": _S,
         **_FACT, "rules_applied": _IDS}
FIX_SCHEMA = {"type": "json_schema", "name": "facts_fix", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["responses"],
    "properties": {"responses": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                            "required": list(_FIXR), "properties": _FIXR}}}}}
_VER = {"id": {"type": "integer"}, "verdict": {"type": "string", "enum": ["accept", "object"]}, "reason": _S,
        "value_text": _S, "low_text": _S, "high_text": _S, "basis": _S, "page": _I, "quote": _S, "rules_applied": _IDS}
VERIFY_SCHEMA = {"type": "json_schema", "name": "facts_verify", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["verdicts"],
    "properties": {"verdicts": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                           "required": list(_VER), "properties": _VER}}}}}
REVIEW_SCHEMA = {"type": "json_schema", "name": "facts_review", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["reviews", "missing", "summary"],
    "properties": {
        "reviews": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                               "required": list(_REV), "properties": _REV}},
        "missing": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                               "required": list(_MISS), "properties": _MISS}},
        "summary": _S}}}


# ---- the document -------------------------------------------------------------------------------------------

def pages(markdown: str) -> dict[int, str]:
    parts = re.split(r"<!-- page (\d+) -->\n", markdown)
    return {int(parts[i]): parts[i + 1] for i in range(1, len(parts) - 1, 2)}


def select(markdown: str, limit: int = MAX_CHARS) -> str:
    """The whole document if it fits, else the first pages plus the pages with the most valuation keywords."""
    if len(markdown) <= limit:
        return markdown
    pg = pages(markdown)
    import context
    around = set(context.pages_for(context.find(markdown)))  # the letter, the scope, definitions: always in
    score = {n: len(KEYWORDS.findall(t)) + (1000 if n in around else 0) for n, t in pg.items()}
    keep, used = set(), 0
    for n in sorted(pg, key=lambda n: (n > 3, -score[n], n)):
        if used + len(pg[n]) > limit:
            continue
        keep.add(n)
        used += len(pg[n])
    return "\n\n".join(f"<!-- page {n} -->\n{pg[n]}" for n in sorted(keep))


def _norm(s: str) -> str:
    s = re.sub(r"<!--.*?-->|[|*#`>]", " ", s or "")
    return re.sub(r"\s+", " ", s).strip().lower()


# ---- checks (no model) --------------------------------------------------------------------------------------

def _squash(s: str) -> str:
    return re.sub(r"\s+", "", _norm(s))


def _in_squashed(num: str, text: str) -> bool:
    """A number (as numbers() gives it: "5223.0", "-70.0", "7.25%") in text with all spacing ignored, reading glued
    figures by the number's own shape: 5,223.0 is in "5 , 2 2 3 . 0 5 , 2 2 3 . 0" (cells run together)."""
    neg, pct = num.startswith("-"), num.endswith("%")
    core = num.lstrip("-").rstrip("%")
    d = len(core.split(".")[1]) if "." in core else 0
    frac = rf"\.\d{{{d}}}" if d else ""
    shape = re.compile(rf"\d{{1,3}}(?:,\d{{3}})+{frac}|\d+{frac}")
    flat = re.sub(r"\s+", "", text or "")
    for m in shape.finditer(flat):
        if abs(float(m.group(0).replace(",", "")) - float(core)) > 1e-9 * max(1, abs(float(core))):
            continue
        before, after = flat[max(0, m.start() - 1):m.start()], flat[m.end():m.end() + 1]
        if neg != (before in ("(", "-", "−", "–")):  # a bracket or minus in front: negative, both ways
            continue
        rest = re.match(r"[.,](\d+)", flat[m.end():])
        if after.isdigit() or (rest and rest[1].strip("0")) or re.search(r"\d[.,]$", flat[:m.start()]):
            continue  # the front or the tail of a longer number: 22 isn't in 22.5, nor 25 in 7.25 (22.0 is 22)
        if pct and after != "%":
            continue
        return True
    return False


MONTHS = {m: i for i, m in enumerate(("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"), 1)}
_MON = r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?"
_DATES = ((re.compile(rf"\b(\d{{1,2}})(?:st|nd|rd|th)?[\s-]+(?:of\s+)?{_MON}[,\s-]+(\d{{4}}|\d{{2}})\b", re.I), "dmy"),
          (re.compile(rf"\b{_MON}\s+(\d{{1,2}})(?:st|nd|rd|th)?,?\s+(\d{{4}})\b", re.I), "mdy"),
          (re.compile(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b"), "iso"),
          (re.compile(r"\b(\d{1,2})[/.-](\d{1,2})[/.-](\d{4}|\d{2})\b"), "num"))


def is_date(f: dict) -> bool:
    return f.get("unit") == "date" or f.get("key") == "valuation_date"


def dates_in(text: str | None) -> list:
    """The dates in a text, in order, as a report prints them: "30 June 2025", "30-Jun-25", "June 30, 2025",
    "30/06/2025" (day first, as Australian reports write it), "2025-06-30". A date's numbers alone (30 and 2025)
    don't make one: its month is a word, and a check on the numbers passes "30 September 2025" for "30 June 2025"."""
    from datetime import date
    hits = []
    for rx, kind in _DATES:
        for m in rx.finditer(text or ""):
            g = m.groups()
            if kind == "dmy":
                d, mo, y = int(g[0]), MONTHS[g[1].lower()[:3]], int(g[2])
            elif kind == "mdy":
                d, mo, y = int(g[1]), MONTHS[g[0].lower()[:3]], int(g[2])
            elif kind == "iso":
                y, mo, d = map(int, g)
            else:
                d, mo, y = map(int, g)
            try:
                hits.append((m.start(), date(y + (2000 if y < 100 else 0), mo, d)))
            except ValueError:
                continue
    return [d for _, d in sorted(set(hits))]


def date_of(text: str | None):
    """The first date in a text, or None."""
    found = dates_in(text)
    return found[0] if found else None


def settle_value(f: dict) -> dict:
    """f with its number worked out from its text, in place: the first number in value_text. None for identity
    text (a name holding digits is still a name), text, or a range the report gives without a preferred point
    (value_text empty, or the range itself); a date is its YYYYMMDD, from the date in the text (the model's where the
    text holds no whole date). Code sets it, so a model's number can't contradict the text, and a model's null isn't
    overwritten."""
    if is_date(f):
        d = date_of(f.get("value_text"))
        if d:
            f["value"] = d.year * 10000 + d.month * 100 + d.day
        return f
    nums, lo, hi = (numbers(f.get(k) or "") for k in ("value_text", "low_text", "high_text"))
    bare = {x.lstrip("-") for x in nums}  # "1.25%–1.75%" reads as 1.25% and -1.75%: the dash isn't a sign
    range_only = bool(lo and hi) and (not nums or (lo[0].lstrip("-") in bare and hi[0].lstrip("-") in bare)) or \
        (not lo and not hi and bool(_range_of(f.get("value_text"))))
    f["value"] = None if f.get("category") == "identity" or not nums or range_only else float(nums[0].rstrip("%"))
    return f


def _unrange(text: str) -> str:
    """A dash between two figures is a range, not a minus: "7.25% - 7.75%" reads as 7.25% to 7.75%, and
    "A$1,900m – A$2,100m" as A$1,900m to A$2,100m. A dash after a word ("Net debt - 850.0") stays a minus."""
    return re.sub(r"(\d(?:%|\s?(?:bn|mn|m|k)\b)?)\s*[-–—]\s*(?=(?:[A-Z]{0,3}\$|€|£)?\s?\d)", r"\1 to ", text or "")


_SCALE = {"k": 1e3, "'000": 1e3, "thousand": 1e3, "m": 1e6, "mn": 1e6, "million": 1e6, "bn": 1e9, "billion": 1e9}
_FIG = re.compile(r"(?P<cur>[A-Z]{0,3}\$|€|£)?\s?(?P<num>\d[\d,]*(?:\.\d+)?)(?P<pct>\s?%)?"
                  r"(?:\s?(?P<scale>thousand|million|billion|bn|mn|m|k|'000)(?![A-Za-z]))?")


def _figs(text: str | None) -> list[dict]:
    """The figures in a text with what they're stated in: [{"value", "scale" ("m", "bn", "k" or None), "cur" ("A$",
    "NZ$", "$", "€" or None), "start", "end"}]. A figure inside a date isn't one."""
    out = []
    for m in _FIG.finditer(_undated(text or "")):
        sc = (m["scale"] or "").lower() or None
        out.append({"value": float(m["num"].replace(",", "")), "scale": None if m["pct"] else sc,
                    "cur": m["cur"].upper() if m["cur"] else None, "start": m.start("num"), "end": m.end()})
    return out


def _undated(text: str) -> str:
    """The text with its dates blanked out (to the same length): the 30 in "30 June 2025" isn't a figure."""
    for rx, _kind in _DATES:
        text = rx.sub(lambda m: " " * len(m.group(0)), text)
    return text


def _range_of(text: str | None) -> tuple[str, str] | None:
    """A range printed as one text ("A$1,900m – A$2,100m", "7.25% to 7.75%"): its two ends, else None."""
    t = _unrange(text or "")
    fig = r"(?:[A-Z]{0,3}\$|€|£)?\s?\d[\d,]*(?:\.\d+)?\s?(?:%|bn|mn|m|k|million|billion)?"
    m = re.fullmatch(rf"\s*({fig})\s+(?:to|and)\s+({fig})\s*", t, re.I)
    return (m[1].strip(), m[2].strip()) if m else None


# the words a figure's label uses, by key: a figure checked against its quote must sit under its own label, not
# another fact's ("WACC of 7.25% and terminal growth of 2.5%": the 2.5% is the growth rate's)
TERMS = {"equity_value": r"equity|net assets|\bshares?\b|unitholder|securit",
         "discount_rate": r"discount|wacc|cost of (?:capital|equity)|hurdle|required return",
         "terminal_growth_rate": r"growth|\bcpi\b|inflation|perpetu",
         "franking_utilisation": r"utili[sz]|gamma|theta",
         "franking_credits_value": r"franking|imputation",
         "terminal_value": r"terminal|exit|perpetu",
         "pv_forecast": r"present value|\bpv\b|discrete|forecast|explicit",
         "pv_terminal_value": r"present value|\bpv\b|terminal"}
TERMS.update(equity_value_ex=TERMS["equity_value"], equity_value_cum=TERMS["equity_value"],
             franking_credits_share=TERMS["franking_credits_value"])
OTHER_TERMS = r"enterprise value|net debt|\bdebt\b|ebitda|revenue|capex|capital expenditure|tax rate"


def _under_other(f: dict, v: str, quote: str) -> str | None:
    """Where every place the quote holds v's figure sits under another fact's label (the nearest label before it,
    looking back past the figures in between), that label; None where one sits under its own, or no label is seen."""
    own = TERMS.get(f.get("key"))
    if not own:
        return None
    others = "|".join([t for k, t in TERMS.items() if t != own] + [OTHER_TERMS])
    q = _unrange(quote or "").lower()
    figs, wanted = _figs(q), [x["value"] for x in _figs(v)]
    seen = None
    for want in wanted:
        at = [i for i, x in enumerate(figs) if abs(x["value"] - want) < 1e-9 * max(1, abs(want))]
        if not at:
            return None  # found only with spacing ignored: the number check speaks for it
        for i in at:
            j, verdict = i, None
            while j >= 0 and verdict is None:
                seg = q[figs[j - 1]["end"] if j else 0:figs[j]["start"]]
                if re.search(own, seg):
                    verdict = "own"
                elif re.search(others, seg):
                    verdict = seg[re.search(others, seg).start():].strip(" :,;(")
                j -= 1
            if verdict in (None, "own"):
                break
            seen = verdict
        else:
            return seen
    return None


def _scale_off(v: str, quote: str) -> str | None:
    """A figure's scale (m, bn, k) and currency (A$, NZ$) must be the quote's where the quote states them: A$2.3m
    isn't A$2.3bn. -> what differs, or None."""
    qf = _figs(quote)
    for x in _figs(v):
        same = [y for y in qf if abs(y["value"] - x["value"]) < 1e-9 * max(1, abs(x["value"]))]
        if not same:
            continue
        if x["scale"] and all(y["scale"] and _SCALE[y["scale"]] != _SCALE[x["scale"]] for y in same):
            return f"{v}: the quote gives it in {same[0]['scale']}, not {x['scale']}"
        named = lambda c: c and c not in ("$",)  # a bare $ is the report's own currency
        if named(x["cur"]) and all(named(y["cur"]) and y["cur"] != x["cur"] for y in same):
            return f"{v}: the quote gives it in {same[0]['cur']}, not {x['cur']}"
    return None


def _at(hay: str, q: str, squashed: bool = False) -> bool:
    """q in hay, not starting or ending inside a number ("25%" isn't in "7.25%") or, spacing kept, inside a word."""
    i = hay.find(q)
    while i >= 0:
        e = i + len(q)
        pre, nxt = hay[max(0, i - 2):i], hay[e:e + 2]
        bad = (q[:1].isdigit() and bool(re.search(r"\d[.,]?$", pre))) or \
              (q[-1:].isdigit() and bool(re.match(r"[.,]?\d", nxt))) or \
              (not squashed and q[:1].isalpha() and pre[-1:].isalpha()) or \
              (not squashed and q[-1:].isalpha() and nxt[:1].isalpha())
        if not bad:
            return True
        i = hay.find(q, i + 1)
    return False


def check(f: dict, pg: dict[int, str]) -> dict:
    """Is the fact supported by the text it cites? Returns {"ok", "items": [{"ok", "text"}]}.
    Strict first; where a report's text layer spaces letters oddly or runs table cells together, a match that
    ignores spacing still counts, and says so. A check an arbiter waived (f["waivers"]) counts, with its note."""
    items = []  # (ok, text, kind): "quote" (where it is), "figure" (a value against its quote), "table", "range"
    q = _norm(f.get("quote"))
    where = [n for n, t in pg.items() if q and _at(_norm(t), q)]
    loose = [] if where or not q else [n for n, t in pg.items() if _at(_squash(t), _squash(f.get("quote")), True)]
    if not q:
        items.append((False, "no quote", "quote"))
    elif f.get("page") in where:
        items.append((True, f"quote found on page {f['page']}", "quote"))
    elif f.get("page") in loose:
        items.append((True, f"quote found on page {f['page']} (spacing ignored)", "quote"))
    elif where or loose:
        items.append((False, f"quote is on page {(where or loose)[0]}, not page {f.get('page')}", "quote"))
    else:
        items.append((False, "quote not found in the document", "quote"))
    quote = f.get("quote") or ""
    undated = _undated(quote)  # the day and year of a date aren't figures: "30 June" doesn't hold A$30m
    for k in ("value_text", "low_text", "high_text"):
        v = (f.get(k) or "").strip()
        if not v or (k == "value_text" and is_date(f) and date_of(v)):  # a date: checked whole, below
            continue
        nums = [] if f.get("category") == "identity" else numbers(_unrange(v))  # a name holding digits is a name
        if nums:
            missing = [n for n in nums if n not in numbers(_unrange(undated))]
            if not missing:
                items.append((True, f"{v} in the quote", "figure"))
            elif all(_in_squashed(n, undated) for n in missing):
                items.append((True, f"{v} in the quote (spacing ignored)", "figure"))
            else:
                items.append((False, f"{v} is not in the quote", "figure"))
                continue
            off = _scale_off(v, quote)
            if off:
                items.append((False, off, "figure"))
            other = _under_other(f, v, undated)
            if other:
                items.append((False, f"{v} sits under '{other}' in the quote, not under the {f.get('label') or f.get('key')}",
                              "figure"))
        elif _norm(v) in q:
            items.append((True, f"'{v}' in the quote", "figure"))
        elif _squash(v) and _squash(v) in _squash(quote):
            items.append((True, f"'{v}' in the quote (spacing ignored)", "figure"))
        else:
            items.append((False, f"'{v}' is not in the quote", "figure"))
    lo, hi = ((f.get(k) or "").strip() for k in ("low_text", "high_text"))
    if bool(lo) != bool(hi):
        items.append((False, f"only the {'low' if lo else 'high'} end of the range is given", "range"))
    elif lo and hi:
        a, b = (_figs(x) for x in (lo, hi))
        if a and b and a[0]["scale"] and b[0]["scale"] and _SCALE[a[0]["scale"]] != _SCALE[b[0]["scale"]]:
            items.append((False, f"the range's ends are in different scales ({lo} and {hi})", "range"))
    nums = numbers(f.get("value_text") or "")
    said = date_of(f.get("value_text")) if is_date(f) else None
    if said:  # a date is checked whole, month included: its numbers alone pass a different month
        ymd = said.year * 10000 + said.month * 100 + said.day
        if said in dates_in(f.get("quote")):
            items.append((True, f"{said:%d %B %Y} is a date in the quote", "figure"))
        elif said in dates_in(re.sub(r"(?<=[A-Za-z])\s+(?=[a-z])", "", f.get("quote") or "")):
            items.append((True, f"{said:%d %B %Y} is a date in the quote (spacing ignored)", "figure"))
        else:
            items.append((False, f"{said:%d %B %Y} is not a date in the quote", "figure"))
        if f.get("value") is not None:
            items.append((int(f["value"]) == ymd, "date matches the text" if int(f["value"]) == ymd
                          else f"date {f['value']} doesn't match {f['value_text']}", "figure"))
    elif f.get("value") is not None and nums and not is_date(f):
        want = float(nums[0].rstrip("%"))
        items.append((abs(want - float(f["value"])) < 1e-9 * max(1, abs(want)),
                      "number matches the text" if abs(want - float(f["value"])) < 1e-9 * max(1, abs(want))
                      else f"number {f['value']} doesn't match {f['value_text']}", "figure"))
    # A quote taken from a table inherits that table's status (it's in the page's <!-- table id (status) --> marker).
    for tid, status in source_tables(f, pg):
        if status == "unread":  # quoted from the PDF's own text while the table waits to be read: that text is real
            items.append((True, f"from table {tid}, still being read; the quote is the PDF's own text", "table"))
        elif status not in ("verified", "approved", "edited", "resolved"):
            items.append((False, f"from table {tid}, which is {status}: settle the table first", "table"))
    # a waiver holds for the quote and figures it was given for: a new quote or figure is checked afresh
    now = {"quote": f.get("quote") or "", **{k: f.get(k) or "" for k in ("value_text", "low_text", "high_text")}}
    holds = lambda w: not w.get("given") or all(w["given"].get(k, "") == v for k, v in now.items())
    # what a waiver may waive: a figure's check; an arbiter's also where the quote is (one from before waivers were
    # tied to their kind keeps that, unless the image gave it)
    reach = lambda w: w.get("any", not str(w.get("by", "")).startswith("the image"))
    waived = {w["check"]: w for w in f.get("waivers") or [] if holds(w)}
    items = [(True, f"{t} — waived by {waived[t]['by']}: {waived[t]['note']}", kind) if not ok and t in waived
             and (kind == "figure" or reach(waived[t])) else (ok, t, kind) for ok, t, kind in items]
    return {"ok": all(ok for ok, _, _ in items), "items": [{"ok": ok, "text": t, "kind": kind} for ok, t, kind in items]}


def waiver(f: dict, check_text: str, by: str, note: str, any_kind: bool = False) -> dict:
    """A waiver of one failing check, tied to the quote and figures it was given for. any_kind: it may waive a check
    of where the quote is (an arbiter's reading of a garbled text layer); else only a figure's check against its
    quote is waived."""
    return {"check": check_text, "by": by, "note": note, "any": any_kind,
            "given": {"quote": f.get("quote") or "", **{k: f.get(k) or "" for k in ("value_text", "low_text", "high_text")}}}


def source_tables(f: dict, pg: dict[int, str]) -> list[tuple[str, str]]:
    """[(table id, status)] for tables on the cited page whose text contains the quote."""
    q = _norm(f.get("quote"))
    out = []
    for m in re.finditer(r"<!-- table (\S+) \((\w+)\) -->\n(.*?)(?=\n\n(?!\|)|\Z)", pg.get(f.get("page"), ""), re.S):
        if q and q in _norm(m.group(3)):
            out.append((m.group(1), m.group(2)))
    return out


# ---- model passes -------------------------------------------------------------------------------------------

def _call(model: str, prompt: str, schema: dict, purpose: str, on_usage) -> dict:
    """One structured call. A reply that isn't valid JSON (a model sometimes runs on in whitespace until it hits
    the output limit) is asked for once more before the step fails."""
    from llm import client, create
    for attempt in (1, 2):
        r = create(client(interactive=False), model, input=prompt, text={"format": schema}, max_output_tokens=12000,
                   purpose=purpose if attempt == 1 else f"{purpose} (again: the first reply wasn't valid JSON)")
        if r.usage and on_usage:
            on_usage(model, r.usage, purpose)
        try:
            return json.loads(r.output_text)
        except json.JSONDecodeError:
            if attempt == 2:
                raise
    raise AssertionError("unreachable")


def extract(markdown: str, model: str, on_usage=None) -> dict:
    return _call(model, EXTRACT_PROMPT.format(keys=KEYS, fields=FIELDS, rules=lessons.rules_text("facts"),
                                              doc=select(markdown)), EXTRACT_SCHEMA, "facts", on_usage)


def review(markdown: str, facts: list[dict], model: str, on_usage=None) -> dict:
    brief = [{"id": f["id"], **{k: f.get(k) for k in _FACT}, "automatic_checks": [i["text"] for i in f["check"]["items"]]}
             for f in facts]
    return _call(model, REVIEW_PROMPT.format(facts=json.dumps(brief, indent=1), rules=lessons.rules_text("facts"), keys=KEYS,
                                             doc=select(markdown)), REVIEW_SCHEMA, "facts-review", on_usage)


# ---- the review and remediation loop --------------------------------------------------------------------------

def cited_pages(markdown: str, page_numbers, limit: int = LOOP_CHARS) -> str:
    """The pages the open facts cite, with their neighbours (then without them, if that's too long), and the report's
    letter, scope and definitions (context.py): what the figures mean depends on them."""
    import context
    pg = pages(markdown)
    cited = {n for n in page_numbers if n in pg}
    around = {n for n in context.pages_for(context.find(markdown)) if n in pg}
    for want in ({m for n in cited for m in (n - 1, n, n + 1)} | around, cited | around, cited):
        keep = sorted(n for n in want if n in pg)
        text = "\n\n".join(f"<!-- page {n} -->\n{pg[n]}" for n in keep)
        if keep and len(text) <= limit:
            return text
    return select(markdown, limit)


def _quote_pages(f: dict, pg: dict[int, str]) -> list[int]:
    """Pages where the fact's quote actually is (a wrong page number is a common slip)."""
    q = _norm(f.get("quote"))
    return [n for n, t in pg.items() if q and q in _norm(t)]


def fact_pages(f: dict, pg: dict[int, str]) -> set[int]:
    """The pages a fact rests on: the one it gives, and those its quote is actually on."""
    return {n for n in [f.get("page"), *_quote_pages(f, pg)] if n}


def _fields(f: dict) -> dict:
    return {k: f.get(k) for k in _FACT}


def _failed(f: dict) -> list[str]:
    return [i["text"] for i in (f.get("check") or {}).get("items", []) if not i["ok"]]


def _open_issue(f: dict) -> dict | None:
    """What's unsettled about a fact after review: failed checks, the reviewer's correction or rejection, or a fact
    the reviewer added that the extractor hasn't confirmed. None if both sides and the checks agree."""
    rv = f.get("review") or {}
    if rv.get("verdict") == "accept" and not _failed(f):
        return None
    sug = rv.get("suggestion")
    return {"verdict": rv.get("verdict"), "reason": rv.get("reason") or "",
            "correction": {k: sug.get(k) for k in ("value_text", "low_text", "high_text", "basis", "page", "quote")} if sug else None}


# ---- the arbiter: a third model for what the loop couldn't settle --------------------------------------------------

ARBITER_PROMPT = """You are an independent arbiter for the key facts taken from last year's valuation report. An
extractor and a reviewer (two other models) went back and forth on each fact below and couldn't settle it: an
automatic check kept failing, or they disagreed on a detail. Read the cited pages and decide each one:
- waive: the fact is right and a failing automatic check is wrong on a clerical point (spacing, a table row
  copied into the quote, separators, formatting). Copy the failing check's text exactly into check, and give a
  one-sentence note for the file saying why it is clerical.
- use_correction: the reviewer's latest correction is right.
- keep: the extractor's version is right as it stands.
- escalate: a person should decide: the report is ambiguous, the figure may really be wrong, or you can't see
  the value in the cited text.
Only waive or pick a version if you can see the value in the page text yourself. When unsure, escalate; a
person will see your note. The brief (what the reference wants):
{keys}

Items:
{items}

Pages:
{doc}"""
_ARB = {"id": {"type": "integer"}, "decision": {"type": "string", "enum": ["waive", "use_correction", "keep", "escalate"]},
        "check": _S, "note": _S}
ARBITER_SCHEMA = {"type": "json_schema", "name": "arbiter", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["decisions"],
    "properties": {"decisions": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                            "required": list(_ARB), "properties": _ARB}}}}}


def arbitrate(markdown: str, facts: list[dict], model: str, on_usage=None) -> int:
    """An independent model decides facts the loop escalated (updated in place). A waiver is kept on the fact
    (f["waivers"]) so every later check honours it and shows its note; nothing is settled unless the checks then
    pass. Returns how many it settled; the rest stay escalated, with the arbiter's note."""
    if not facts:
        return 0
    pg = pages(markdown)
    items = []
    for f in facts:
        sug = (f.get("review") or {}).get("suggestion")
        items.append({"id": f["id"], "fact": _fields(f), "checks_failed": _failed(f),
                      "open_point": (f.get("agent") or {}).get("open", {}).get("reason"),
                      "reviewer_correction": {k: sug.get(k) for k in _FACT} if sug else None,
                      "rounds": [{"extractor": t.get("extractor"), "reviewer": t.get("reviewer")}
                                 for t in (f.get("agent") or {}).get("thread", [])[-3:]]})
    cites = [f.get("page") for f in facts] + [n for f in facts for n in _quote_pages(f, pg)]
    out = _call(model, ARBITER_PROMPT.format(keys=KEYS, items=json.dumps(items, indent=1, ensure_ascii=False),
                                             doc=cited_pages(markdown, [c for c in cites if c] or [1])),
                ARBITER_SCHEMA, "facts-arbiter", on_usage)
    by_id, settled = {f["id"]: f for f in facts}, 0
    for dec in out["decisions"]:
        f = by_id.get(dec["id"])
        if not f:
            continue
        a = f["agent"]
        entry = {"round": "arbiter", "arbiter": {"by": model, "decision": dec["decision"], "check": dec["check"],
                                                 "note": dec["note"]}}
        a["thread"].append(entry)
        ok = False
        if dec["decision"] == "waive" and dec["check"]:
            waived = {**f, "waivers": (f.get("waivers") or []) + [waiver(f, dec["check"].strip(), model, dec["note"],
                                                                          any_kind=True)]}
            chk = check(waived, pg)
            if chk["ok"]:
                f["waivers"], f["check"], ok = waived["waivers"], chk, True
        elif dec["decision"] == "use_correction":
            sug = (f.get("review") or {}).get("suggestion")
            if sug:
                chk = check({**sug, "waivers": f.get("waivers")}, pg)
                if chk["ok"]:
                    f.update({k: sug.get(k) for k in _FACT})
                    f["check"], ok = chk, True
        elif dec["decision"] == "keep":
            chk = check(f, pg)
            f["check"], ok = chk, chk["ok"]
        if ok:
            a.update(status="agreed", round="arbiter", waivers=f.get("waivers") or [])
            a.pop("open", None)
            settled += 1
        else:
            entry["arbiter"]["outcome"] = "escalated" if dec["decision"] == "escalate" else "didn't pass the checks"
            a["open"] = {**(a.get("open") or {}), "arbiter": dec["note"]}
    return settled


def resolve(markdown: str, facts: list[dict], model: str, reviewer_model: str, on_usage=None, progress=None,
            rounds: int = MAX_ROUNDS, arbiter_model: str | None = None) -> dict:
    """The loop on reviewed facts (updated in place: each gets "agent" = {status, round, thread}). A fact is
    agreed when the reviewer accepts it and the checks pass, withdrawn when the reviewer accepts the extractor's
    withdrawal, and escalated when rounds run out. Returns {"summary", "episodes"} (episodes feed lessons.py)."""
    progress = progress or (lambda f, m: None)
    pg = pages(markdown)
    by_id = {f["id"]: f for f in facts}
    issues = {}
    for f in facts:
        iss = _open_issue(f)
        rv = f.get("review") or {}
        thread = [{"round": 0, "extractor": {"action": "extract" if f.get("origin") != "reviewer" else "-"},
                   "reviewer": {"verdict": rv.get("verdict"), "reason": rv.get("reason")}, "checks_failed": _failed(f)}]
        f["agent"] = {"status": "agreed" if iss is None else "open", "round": 0, "thread": thread}
        if iss:
            issues[f["id"]] = iss
    k = 0
    while issues and k < rounds:
        k += 1
        progress((k - 1) / rounds, f"Fact review loop, round {k}: {len(issues)} fact(s) open ({model} fixes, {reviewer_model} checks)")
        cites = [by_id[i].get("page") for i in issues] + [(iss["correction"] or {}).get("page") for iss in issues.values()] \
            + [n for i in issues for n in _quote_pages(by_id[i], pg)]
        doc = cited_pages(markdown, [c for c in cites if c])
        brief = [{"id": i, "added_by_reviewer": by_id[i].get("origin") == "reviewer", "fact": _fields(by_id[i]),
                  "checks_failed": _failed(by_id[i]), "reviewer": iss} for i, iss in issues.items()]
        fix = _call(model, FIX_PROMPT.format(fields=FIELDS, keys=KEYS, rules=lessons.rules_text("facts"),
                                             issues=json.dumps(brief, indent=1, ensure_ascii=False), doc=doc),
                    FIX_SCHEMA, "facts-fix", on_usage)
        answers = {r["id"]: r for r in fix["responses"] if r["id"] in issues}
        lessons.applied([x for r in answers.values() for x in r["rules_applied"]])
        pending, items = {}, []
        for i in issues:
            f, r = by_id[i], answers.get(i) or {"action": "keep", "reason": "(no answer)", "rules_applied": []}
            cand = _fields(f)
            if r["action"] == "revise":
                cand.update({x: r[x] for x in _FACT if x in r and r[x] not in (None, "") or x in ("low_text", "high_text")
                             and x in r})
                cand["category"] = cand["category"] if cand.get("category") in CATEGORIES else f["category"]
                cand["key"] = cand.get("key") or f["key"]
                settle_value(cand)
            cand["check"] = check(cand, pg) if r["action"] == "revise" else f["check"]
            pending[i] = (r, cand)
            items.append({"id": i, "you_said": issues[i], "extractor": {"action": r["action"], "reason": r["reason"]},
                          "fact_now": _fields(cand), "automatic_checks_now": [x["text"] for x in cand["check"]["items"]]})
        cites = [c.get("page") for _, c in pending.values()] + [n for _, c in pending.values() for n in _quote_pages(c, pg)]
        progress((k - 0.5) / rounds, f"Fact review loop, round {k}: {reviewer_model} checks {model}'s answers on "
                                     f"{len(pending)} fact(s)")
        ver = _call(reviewer_model, VERIFY_PROMPT.format(rules=lessons.rules_text("facts"), keys=KEYS, items=json.dumps(
            items, indent=1, ensure_ascii=False), doc=cited_pages(markdown, [c for c in cites if c] or [1])),
            VERIFY_SCHEMA, "facts-verify", on_usage)
        verdicts = {v["id"]: v for v in ver["verdicts"]}
        lessons.applied([x for v in verdicts.values() for x in v["rules_applied"]])
        issues = {}
        for i, (r, cand) in pending.items():
            f = by_id[i]
            v = verdicts.get(i) or {"verdict": "object", "reason": "the reviewer gave no verdict"}
            ok = cand["check"]["ok"]
            f["agent"]["thread"].append({"round": k, "extractor": {"action": r["action"], "reason": r["reason"]},
                                         "fact": _fields(cand) if r["action"] == "revise" else None,
                                         "reviewer": {"verdict": v["verdict"], "reason": v["reason"]},
                                         "checks_failed": [x["text"] for x in cand["check"]["items"] if not x["ok"]]})
            if r["action"] == "revise":  # the discussion moves on to the latest version
                f.update(_fields(cand))
                f["check"] = cand["check"]
            if v["verdict"] == "accept" and r["action"] == "withdraw":
                f["agent"].update(status="withdrawn", round=k)
            elif v["verdict"] == "accept" and ok:
                f["agent"].update(status="agreed", round=k)
            else:
                corr = {x: v.get(x) for x in ("value_text", "low_text", "high_text", "basis", "page", "quote")
                        if v.get(x) not in (None, "")}
                reason = v["reason"] if v["verdict"] == "object" else \
                    "accepted by the reviewer, but the automatic checks still fail: " + "; ".join(_failed(cand))
                issues[i] = {"verdict": "object", "reason": reason, "correction": corr or None,
                             "accepted": v["verdict"] == "accept"}
    for i, iss in issues.items():  # rounds ran out: a person decides, with the reviewer's latest correction to hand
        f = by_id[i]
        f["agent"].update(status="escalated", round=k, open=iss)
        if iss.get("correction"):
            fix = settle_value({**_fields(f), **iss["correction"]})
            fix["check"] = check(fix, pg)
            f["review"] = {**(f.get("review") or {}), "suggestion": fix}
    arbitrated = 0
    if arbiter_model and issues:
        progress(0.95, f"Arbiter ({arbiter_model}) on {len(issues)} fact(s) the loop couldn't settle")
        arbitrated = arbitrate(markdown, [by_id[i] for i in issues], arbiter_model, on_usage)
    count = lambda st: sum(f["agent"]["status"] == st for f in facts)
    summary = {"rounds": k, "agreed": count("agreed"), "withdrawn": count("withdrawn"), "escalated": count("escalated"),
               "settled_in_loop": sum(f["agent"]["status"] in ("agreed", "withdrawn") and f["agent"]["round"] not in (0, None)
                                      for f in facts), "arbitrated": arbitrated}
    episodes = [{"category": f["category"], "key": f["key"], "added_by_reviewer": f.get("origin") == "reviewer",
                 "thread": f["agent"]["thread"], "outcome": f["agent"]["status"]}
                for f in facts if len(f["agent"]["thread"]) > 1]
    progress(1.0, f"Fact review loop: {summary['agreed']} agreed, {summary['withdrawn']} withdrawn, {summary['escalated']} for you")
    return {"summary": summary, "episodes": episodes}


def run(markdown: str, model: str = "gpt-6-luna", reviewer_model: str = "gpt-6-sol",
        on_usage=None, progress=None, loop_rounds: int = MAX_ROUNDS, arbiter_model: str | None = None) -> dict:
    """All four passes. Returns {"facts": [...], "notes", "review_summary", "loop"}; each fact carries check,
    review (with a checked suggestion if the reviewer corrected it) and agent (the loop's outcome and thread)."""
    progress = progress or (lambda f, m: None)
    pg = pages(markdown)

    def checked(f):
        f["check"] = check(settle_value(f), pg)
        return f

    progress(0.1, f"Extracting key facts ({model})")
    ext = extract(markdown, model, on_usage)
    facts = [checked({"id": i, "origin": "extractor", **f})
             for i, f in enumerate((f for f in ext["facts"] if f["key"] in KNOWN), start=1)]
    progress(0.45, f"Reviewing {len(facts)} facts ({reviewer_model})")
    rev = review(markdown, facts, reviewer_model, on_usage)
    by_id = {f["id"]: f for f in facts}
    for r in rev["reviews"]:
        f = by_id.get(r["id"])
        if not f:
            continue
        f["review"] = {"verdict": r["verdict"], "reason": r["reason"]}
        if r["verdict"] == "correct":
            fix = {**{k: f.get(k) for k in _FACT}, **{k: r[k] for k in ("value_text", "low_text", "high_text", "basis",
                                                                          "page", "quote") if r[k] not in (None, "")}}
            f["review"]["suggestion"] = checked(fix)
    next_id = len(facts) + 1
    for m in (x for x in rev["missing"] if x["key"] in KNOWN):  # the brief only
        f = checked({"id": next_id, "origin": "reviewer", **{k: m[k] for k in _FACT}})
        f["review"] = {"verdict": "added", "reason": m["why"]}
        facts.append(f)
        next_id += 1
    loop = resolve(markdown, facts, model, reviewer_model, on_usage,
                   lambda f, m: progress(0.75 + 0.25 * f, m), arbiter_model=arbiter_model) if loop_rounds else None
    progress(1.0, "Facts extracted and reviewed")
    return {"facts": facts, "notes": ext.get("notes"), "review_summary": rev.get("summary"), "loop": loop}


# ---- the equity value the bridge starts from ----------------------------------------------------------------------
# Always the equity value, as a low / high range with the midpoint as the mid (the convention: the midpoint of the
# low and the high, whatever preferred value the report prints). Ex-distribution unless the report is overwhelmingly
# cum-distribution; the rebuild can still switch to cum where only that figure ties (the model's cash flows are
# cum-distribution by default).

_EX = re.compile(r"\bex[- ]?(?:distribution|dividend|div)\b", re.I)
_CUM = re.compile(r"\bcum[- ]?(?:distribution|dividend|div)\b", re.I)
OVERWHELMING = 3  # cum-distribution mentions per ex-distribution mention that make the report "overwhelmingly" cum


def _basis(f: dict) -> str | None:
    if f.get("key") == "equity_value_cum":
        return "cum"
    if f.get("key") == "equity_value_ex":
        return "ex"
    text = " ".join(f.get(k) or "" for k in ("basis", "label", "quote"))
    ex, cum = bool(_EX.search(text)), bool(_CUM.search(text))
    return "ex" if ex and not cum else "cum" if cum and not ex else None


def _num(text: str | None) -> float | None:
    got = numbers(text or "")
    return float(got[0].rstrip("%")) if got else None


def conclusion(facts: list[dict], markdown: str = "") -> dict | None:
    """The equity value to bridge from, among the facts (approved or agreed; rejected and withdrawn ones left out):
    {"fact_id", "key", "basis" ("ex", "cum" or None: not stated, taken as ex), "low", "high", "mid", "printed" (the
    report's own preferred or midpoint value, if any), "units", "why": [...], "other": the other basis's figure}."""
    live = [f for f in facts if f.get("key") in ("equity_value", "equity_value_ex", "equity_value_cum")
            and f.get("status") != "rejected" and (f.get("agent") or {}).get("status") != "withdrawn"]
    if not live:
        return None
    n_ex, n_cum = len(_EX.findall(markdown or "")), len(_CUM.findall(markdown or ""))
    by = {b: [f for f in live if _basis(f) == b] for b in ("ex", "cum", None)}
    why = []
    if by["cum"] and n_cum >= OVERWHELMING * max(1, n_ex) and not (by["ex"] and n_ex > 1):
        pick, basis = by["cum"][0], "cum"
        why.append(f"the report is overwhelmingly cum-distribution ({n_cum} mentions against {n_ex} ex-distribution)")
    elif by["ex"]:
        pick, basis = by["ex"][0], "ex"
        why.append("ex-distribution, as the report states it" + (" (the cum-distribution figure is kept beside it)"
                                                                  if by["cum"] else ""))
    elif by[None]:
        pick, basis = by[None][0], None
        why.append("the report doesn't say ex- or cum-distribution: taken as ex-distribution, the usual basis")
    else:
        pick, basis = by["cum"][0], "cum"
        why.append("only a cum-distribution equity value is given")
    texts = {k: pick.get(k) for k in ("value_text", "low_text", "high_text")}
    rng = _range_of(texts["value_text"]) if not texts["low_text"] and not texts["high_text"] else None
    if rng:  # the range printed as one text: its ends are the low and the high, and it has no preferred point
        texts = {"value_text": "", "low_text": rng[0], "high_text": rng[1]}
        why.append(f"the range {pick.get('value_text')} is the low and the high")
    low, high, printed = _num(texts["low_text"]), _num(texts["high_text"]), _num(texts["value_text"])
    mid = (low + high) / 2 if low is not None and high is not None else printed
    if low is not None and high is not None:
        why.append(f"mid = the midpoint of {texts['low_text']} and {texts['high_text']}")
        if printed is not None:
            d = len((numbers(texts["value_text"])[0].rstrip("%").split(".") + [""])[1])
            if abs(round(mid, d) - printed) > 0.5 * 10 ** -d + 1e-9:
                why.append(f"the report prints {pick.get('value_text')}, not the midpoint: the midpoint is used")
    else:
        why.append("no range in the report: its one figure is the low, the mid and the high")
        low = high = mid
    other = next((f for b in ("ex", "cum") if b != basis for f in by[b]), None)
    units = next((f.get("value_text") for f in facts if f.get("key") == "currency_units"), None) or pick.get("unit")
    return {"fact_id": pick.get("id"), "key": pick.get("key"), "basis": basis, "low": low, "high": high, "mid": mid,
            "printed": printed, "texts": texts,
            "page": pick.get("page"), "units": units, "why": why,
            "other": {"fact_id": other.get("id"), "basis": _basis(other), "low": _num(other.get("low_text")),
                      "high": _num(other.get("high_text")), "value": _num(other.get("value_text")),
                      "texts": {k: other.get(k) for k in ("value_text", "low_text", "high_text")}} if other else None}


if __name__ == "__main__":
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent))
    md = Path(sys.argv[1]).read_text(encoding="utf-8")
    model = sys.argv[2] if len(sys.argv) > 2 else "gpt-6-luna"
    res = run(md, model, progress=lambda f, m: print(f"{f:4.0%} {m}"))
    for f in res["facts"]:
        rv = f.get("review", {})
        print(f"{f['id']:>2} {f['category']:10s} {f['key']:24s} {f['value_text']!r:18} p{f['page']} "
              f"check={'ok' if f['check']['ok'] else [i['text'] for i in f['check']['items'] if not i['ok']]} "
              f"review={rv.get('verdict')}: {rv.get('reason', '')[:80]}")
    print("summary:", res["review_summary"])
