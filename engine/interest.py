"""The interest valued: the share of the entity the equity value is for, as the report states it and as the overlay
applies it to the cash flows.

  applied  the share the overlay applies on the way to the equity value: a row of one figure between 0 and 1 multiplied
           into a discounting (dcftrace: core["share"]), or a cell labelled as an interest, a share or a stake
           multiplying a term on the path above the discountings (after the discounting); else none (100%)
  check    the report's interest_valued against it: different, held; applied where the report doesn't say, held (a
           person can say why); stated below 100% with none applied, a point to check (the client model may give the
           cash flows at the share already)
"""
import re

import sourced

SHARE_WORDS = re.compile(r"\binterest\b|ownership|\bstake\b|\bholding\b|shareholding|attributable|\bproportion\b|"
                         r"\bportion\b|equity share|share of|% share|\bshare\b", re.I)
NOT_SHARE = re.compile(r"tax|frank|gamma|theta|utili[sz]|growth|inflation|cpi|escalat|margin|discount|cost of|"
                       r"probabilit|\brate\b|rate of|interest (?:rate|expense|income|paid|received|cover)", re.I)
SAME = 1e-6  # the report's % and the overlay's share, as fractions


def _frac(v) -> float | None:
    if not isinstance(v, (int, float)) or isinstance(v, bool):
        return None
    return v if 0 < v < 1 else v / 100 if 1 < v < 100 else None


def applied(tree: dict | None) -> dict | None:
    """The share the overlay applies on the way to the equity value: {"value" (a fraction), "cell" (or row), "label",
    "where" ("inside" a discounting / "after" it)}, or None where none is found (the whole of the cash flows)."""
    if not tree:
        return None
    found = []

    def walk(n, depth=0):
        if n.get("again") or depth > 40:
            return
        for c in n.get("cores") or []:
            if c.get("share"):
                found.append({"value": c["share"]["value"], "cell": c["share"]["row"], "label": c["share"].get("label"),
                              "where": "inside", "at": c["cell"]})
        if n.get("on_path") and "*" in (n.get("formula") or ""):
            for ch in n.get("children") or []:
                lab = ch.get("label") or ""
                v = _frac(ch.get("value"))
                if not ch.get("on_path") and v is not None and SHARE_WORDS.search(lab) and not NOT_SHARE.search(lab):
                    found.append({"value": v, "cell": ch["cell"], "label": lab, "where": "after", "at": n["cell"]})
        for ch in n.get("children") or []:
            walk(ch, depth + 1)

    walk(tree)
    if not found:
        return None
    one = found[0]
    if any(abs(x["value"] - one["value"]) > SAME for x in found[1:]):
        one = {**one, "others": [x for x in found[1:] if abs(x["value"] - one["value"]) > SAME]}
    return one


def check(summary: dict, facts: list[dict], trees: dict) -> dict:
    """{"report" (a fraction or None), "report_text", "model" (applied() for the low, or None), "holds"}."""
    import result
    stated = sourced._stated(facts, "interest_valued")
    rep = stated[0][0] / 100 if stated else None
    model = applied((trees.get("low") or {}).get("tree")) or applied((trees.get("high") or {}).get("tree"))
    holds = []
    pct = lambda x: f"{100 * x:.4g}%"
    if model and model.get("others"):
        holds.append(result.hold(summary, "interest-two", [model["cell"], [x["cell"] for x in model["others"]]],
                                 "The overlay applies more than one ownership share on the way to the value",
                                 f"{pct(model['value'])} at {model['cell']} ({model.get('label') or ''}) and "
                                 + ", ".join(f"{pct(x['value'])} at {x['cell']}" for x in model["others"])))
    if rep is not None and model and abs(rep - model["value"]) > SAME:
        holds.append(result.hold(summary, "interest", [rep, model["value"], model["cell"]],
                                 f"The report values a {pct(rep)} interest; the overlay applies {pct(model['value'])}",
                                 f"the report: {stated[0][1]}; the overlay: {model['cell']} ({model.get('label') or ''}), "
                                 f"{'inside the discounting at ' if model['where'] == 'inside' else 'after the discounting, at '}"
                                 f"{model['at']}. The value rolled forward is on the overlay's share"))
    elif rep is None and model:
        holds.append(result.hold(summary, "interest", [None, model["value"], model["cell"]],
                                 f"The overlay values {pct(model['value'])} of the cash flows; the report doesn't say what "
                                 "interest it values",
                                 f"{model['cell']} ({model.get('label') or ''}) multiplies "
                                 f"{'the cash flows inside the discounting at ' if model['where'] == 'inside' else 'the value after the discounting, at '}"
                                 f"{model['at']}: confirm the interest valued (the report's statement of it, or why)"))
    elif rep is not None and rep < 1 - SAME and not model:
        holds.append(result.hold(summary, "interest", [rep, None],
                                 f"The report values a {pct(rep)} interest; no share is applied to the overlay's cash flows",
                                 f"the report: {stated[0][1]}. Either the client model gives the cash flows at the share "
                                 "already, or the overlay's value is for the whole: check which", severity="check"))
    return {"report": rep, "report_text": stated[0][1] if stated else None, "model": model, "holds": holds}
