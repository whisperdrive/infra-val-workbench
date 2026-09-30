"""Seeing the report: its key tables read from their images, and each key fact confirmed on the image of where it sits.

A report's text layer can't always be trusted. A multi-column layout, or a table with merged headings, comes out with
its lines interleaved or its figures under the wrong column; a table pasted as a picture has no text at all. So:

  tables  before the facts, the tables most likely to hold the key figures (valuation words in and around them, and
          every picture on a page that has them) are read from their images (docingest.read_and_check: luna reads;
          code checks the read against the page's own characters where there are any, else sol reads it again,
          independently, and the two reads are compared). A read the checks don't pass goes round the table loop
          (docingest.resolve_tables: luna fixes, sol checks against the image, an arbiter for what's left). What's
          read replaces the text layer's lines in the document the facts are extracted from.
  facts   after them, each fact with a figure is looked up on the image of where it sits (its table's image, else
          its page's) blind: luna reads the item's figures and the column each sits under without being told what
          was extracted, and code compares. Where they differ, sol looks at the image with both versions. A
          correction stands when two reads of the image agree on it (the text layer is then outranked, and the check
          it fails is waived with that reason); what that doesn't settle goes to a person, with the image. A slide's
          own text and tables are exact (read from the file), so they aren't looked at again.
"""
import json
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import calllog
import docingest
import keyfacts

KEY_TABLES = 8       # tables read from their images, most likely to hold the key figures first
KEY_PICTURES = 6     # and pictures on pages that talk about the valuation, beside them
PAGE_DPI = 130       # a page's image, for a fact in its prose
NATIVE = ("pptx table", "pptx chart data")
NUMERIC = ("conclusion", "assumption")

BLIND_PROMPT = """This image is from {where} of last year's valuation report. Find each item below in it and read its
figures exactly as printed (thousands separators, decimals, %, currency and units, brackets for negatives): value is
the single figure (the preferred value, or the midpoint where the report gives one), low and high the ends of a range,
"" for what the image doesn't show. column: the heading of the column each figure sits under ("Low", "Mid", "High",
a year), or "" in running text; row: the row's label, or what the sentence is about. Only what you can see in this
image: found false if the item isn't in it.

Items:
{items}"""
VERDICT_PROMPT = """You are the reviewer. Two readings of last year's valuation report disagree on the items below: the
extraction from the report's text layer (a layout with columns or merged headings can scramble it) and a reading of
this image of {where}. Look at the image yourself. For each item: choice "extracted" if the extraction is right,
"image" if the image reading is right, "neither" if both are wrong (then give the figures as the image shows them),
"not shown" if the image doesn't show the item. Figures exactly as printed; say why in one sentence.

Items:
{items}"""
_S = {"type": "string"}
_READ = {"id": {"type": "integer"}, "found": {"type": "boolean"}, "value": _S, "low": _S, "high": _S, "column": _S,
         "row": _S}
BLIND_SCHEMA = {"type": "json_schema", "name": "visual_read", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["reads"],
    "properties": {"reads": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                        "required": list(_READ), "properties": _READ}}}}}
_VER = {"id": {"type": "integer"}, "choice": {"type": "string", "enum": ["extracted", "image", "neither", "not shown"]},
        "value": _S, "low": _S, "high": _S, "reason": _S}
VERDICT_SCHEMA = {"type": "json_schema", "name": "visual_verdict", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["verdicts"],
    "properties": {"verdicts": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                           "required": list(_VER), "properties": _VER}}}}}


# ---- the key tables, from their images -------------------------------------------------------------------------------

def key_tables(doc: dict) -> list[dict]:
    """The unread tables most likely to hold the key figures, and the pictures on pages that talk about the valuation."""
    unread = [t for t in doc["tables"] if t.get("status") == "unread" and t.get("png")]
    ranked = docingest.rank_tables(unread, doc["pages"])
    pick = [t for score, t in ranked if score > 0][:KEY_TABLES]
    pictures = [t for score, t in ranked if score > 0 and t["source"] in ("picture", "page image") and t not in pick]
    return pick + pictures[:KEY_PICTURES]


def read_tables(doc: dict, out_dir: str | Path, reader, progress=None, arbiter_model: str | None = None) -> dict:
    """Read the key tables from their images (in place), the loop on those whose reads don't check out, and the
    document's Markdown again. -> {"read", "verified", "resolved", "flagged", "figures", "tables": [ids]}."""
    progress = progress or (lambda frac, msg: None)
    todo = key_tables(doc)
    before = [t for t in doc["tables"] if t.get("read_for") == "key figures"]  # read on an earlier run
    if not todo:
        count = lambda st: sum(t.get("status") == st for t in before)
        return {"read": len(before), "verified": count("verified"), "resolved": count("resolved"),
                "flagged": count("flagged"), "figures": count("figure"), "tables": [t["id"] for t in before]}
    out_dir = Path(out_dir)
    done = [0]

    def one(t):
        try:
            t.update(docingest.read_and_check(reader, t, (out_dir / t["png"]).read_bytes(), t.get("text_lines")))
        except Exception as e:  # the table keeps the page's own text; the facts still see it
            t.update(status="unread", read_error=f"{type(e).__name__}: {e}")
        t["read_for"] = "key figures"
        done[0] += 1
        progress(done[0] / len(todo), f"Read {done[0]} of {len(todo)} key tables from their images")
        return t

    with ThreadPoolExecutor(docingest.READERS) as pool:
        list(pool.map(calllog.carry(one), todo))
    flagged = {t["id"] for t in todo if t.get("status") == "flagged"}
    loop = None
    if flagged:
        progress(0.9, f"Table review loop on {len(flagged)} key table(s) whose reads don't check out")
        loop = docingest.resolve_tables(doc, out_dir, reader.model, reader.reviewer_model, None, reader.on_usage,
                                        ids=flagged, arbiter_model=arbiter_model)
    docingest.render(doc)
    todo = before + todo
    count = lambda st: sum(t.get("status") == st for t in todo)
    return {"read": len(todo), "verified": count("verified"), "resolved": count("resolved"), "flagged": count("flagged"),
            "figures": count("figure"), "tables": [t["id"] for t in todo], "loop": {k: v for k, v in (loop or {}).items()
                                                                                    if k != "episodes"} or None}


# ---- each key fact, on its image ----------------------------------------------------------------------------------------

def _nums(text: str | None) -> list[float]:
    return [float(n.rstrip("%")) for n in keyfacts.numbers(keyfacts._unrange(text or ""))]


def _figures(x: dict, keys=("value_text", "low_text", "high_text")) -> dict:
    return {k.removesuffix("_text"): (_nums(x.get(k)) or [None])[0] for k in keys}


def _same(fact: dict, read: dict) -> bool:
    """The image reading gives the fact's figures in the fact's places (a low read as a high is a difference)."""
    f, r = _figures(fact), _figures(read, ("value", "low", "high"))
    have = {k: v for k, v in f.items() if v is not None}
    if not have:
        return True
    if all(r.get(k) is not None and abs(r[k] - v) < 1e-9 * max(1, abs(v)) for k, v in have.items()):
        return True
    # a single figure the reading put as the value where the extraction has it as both ends, or the reverse
    one = {v for v in have.values()} | set()
    got = {v for v in r.values() if v is not None}
    return len(one) == 1 and got == one


def image_for(f: dict, doc: dict, pg: dict, out_dir: Path, source_path: str | None) -> tuple[str, str] | None:
    """(the image of where a fact sits, relative to out_dir; what it is) or None where it needn't be looked at again
    (a slide's own text or table, read exactly from the file)."""
    tabs = {t["id"]: t for t in doc["tables"]}
    for tid, _status in keyfacts.source_tables(f, pg):
        t = tabs.get(tid)
        if t and t.get("source") in NATIVE:
            return None
        if t and t.get("png") and (out_dir / t["png"]).exists():
            return t["png"], f"its table ({t['source']}, {t['where']})"
    if doc["kind"] != "pdf" or not f.get("page"):
        return None
    rel = f"pages/p{int(f['page']):03d}.png"
    if not (out_dir / rel).exists():
        if not source_path:
            return None
        import pdfplumber
        with pdfplumber.open(source_path) as pdf:
            if not 1 <= int(f["page"]) <= len(pdf.pages):
                return None
            (out_dir / "pages").mkdir(parents=True, exist_ok=True)
            (out_dir / rel).write_bytes(docingest._png(pdf.pages[int(f["page"]) - 1].to_image(resolution=PAGE_DPI).original))
    return rel, f"page {f['page']}"


def _apply(f: dict, fig: dict, pg: dict, why: str) -> None:
    """Set a fact's figures to what two reads of the image agree on; a check the text layer then fails is waived."""
    for k in ("value", "low", "high"):
        if fig.get(k) is not None:
            f[f"{k}_text"] = fig[k]
    keyfacts.settle_value(f)
    chk = keyfacts.check(f, pg)
    failed = [i["text"] for i in chk["items"] if not i["ok"]]
    if failed:
        f["waivers"] = (f.get("waivers") or []) + [{"check": c, "by": "the image (two reads)", "note": why} for c in failed]
        chk = keyfacts.check(f, pg)
    f["check"] = chk


def confirm(doc: dict, facts: list[dict], reader, out_dir: str | Path, source_path: str | None, progress=None) -> dict:
    """Look up each fact with a figure on its image (in place: fact["visual"], and the fact corrected or handed to a
    person where the image says otherwise). -> {"looked", "confirmed", "corrected", "escalated", "images"}."""
    progress = progress or (lambda frac, msg: None)
    out_dir = Path(out_dir)
    pg = keyfacts.pages(doc["markdown"])
    groups: dict[str, list[dict]] = {}
    where: dict[str, str] = {}
    for f in facts:
        if (f.get("agent") or {}).get("status") == "withdrawn" or not (
                f.get("category") in NUMERIC or f.get("key") == "valuation_date") or not any(
                _nums(f.get(k)) for k in ("value_text", "low_text", "high_text")):
            continue
        img = image_for(f, doc, pg, out_dir, source_path)
        if not img:
            f["visual"] = {"status": "native", "note": "read exactly from the slide, not from an image"}
            continue
        groups.setdefault(img[0], []).append(f)
        where[img[0]] = img[1]
    n = {"looked": 0, "confirmed": 0, "corrected": 0, "escalated": 0, "images": len(groups)}
    for i, (rel, fs) in enumerate(groups.items(), 1):
        progress(i / max(1, len(groups)), f"Checking {len(fs)} fact(s) on {where[rel]} ({i} of {len(groups)} images)")
        png = (out_dir / rel).read_bytes()
        items = [{"id": k, "item": f.get("label") or f["key"], "key": f["key"],
                  "kind": "a range (low and high)" if f.get("low_text") or f.get("high_text") else "one figure"}
                 for k, f in enumerate(fs)]
        reads = {r["id"]: r for r in reader._call(reader.model, BLIND_PROMPT.format(
            where=where[rel], items=json.dumps(items, indent=1)), png, BLIND_SCHEMA, "visual-read")["reads"]}
        disputed = []
        for k, f in enumerate(fs):
            n["looked"] += 1
            r = reads.get(k) or {"found": False}
            base = {"image": rel, "where": where[rel], "read": {x: r.get(x) for x in ("value", "low", "high", "column", "row")},
                    "by": reader.model}
            if r.get("found") and _same(f, r):
                f["visual"] = {**base, "status": "confirmed"}
                n["confirmed"] += 1
            else:
                f["visual"] = {**base, "status": "disputed", "found": bool(r.get("found"))}
                disputed.append(k)
        if not disputed:
            continue
        brief = [{"id": k, "item": fs[k].get("label") or fs[k]["key"], "key": fs[k]["key"],
                  "extracted": {x: fs[k].get(x) for x in ("value_text", "low_text", "high_text", "basis", "quote")},
                  "image_reading": fs[k]["visual"]["read"] if fs[k]["visual"]["found"] else "not found in the image"}
                 for k in disputed]
        verdicts = {v["id"]: v for v in reader._call(reader.reviewer_model, VERDICT_PROMPT.format(
            where=where[rel], items=json.dumps(brief, indent=1, ensure_ascii=False)), png, VERDICT_SCHEMA,
            "visual-verdict")["verdicts"]}
        for k in disputed:
            f, v = fs[k], verdicts.get(k) or {"choice": "not shown", "reason": "the reviewer gave no verdict"}
            vis = f["visual"]
            vis.update(verdict={x: v.get(x) for x in ("choice", "value", "low", "high", "reason")}, reviewer=reader.reviewer_model)
            read = {x: vis["read"].get(x) for x in ("value", "low", "high")}
            said = {x: v.get(x) or None for x in ("value", "low", "high")}
            was = {x: f.get(f"{x}_text") for x in ("value", "low", "high")}
            agree_read = vis["found"] and (v["choice"] == "image" or (v["choice"] == "neither" and _same(
                {f"{x}_text": said[x] for x in said}, read)))
            if v["choice"] == "extracted" or (v["choice"] == "neither" and _same(f, {x: said[x] for x in said})):
                vis["status"] = "confirmed by the reviewer"
                n["confirmed"] += 1
            elif agree_read:
                _apply(f, {x: read[x] or None for x in read}, pg,
                       f"{reader.model} and {reader.reviewer_model} read {where[rel]} alike: {v.get('reason') or ''}")
                vis.update(status="corrected", was={x: y for x, y in was.items() if y})
                n["corrected"] += 1
            else:
                vis["status"] = "escalated"
                a = f.setdefault("agent", {"status": "open", "round": 0, "thread": []})
                a.update(status="escalated", open={"verdict": "object", "reason": (
                    f"on the image of {where[rel]}: " + (v.get("reason") or "the reads don't agree")), "correction": None,
                    "image": rel})
                n["escalated"] += 1
    return n
