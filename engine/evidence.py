"""What the row tools agreed on, and whether they were right: learned from a person's decisions, without a word of
the client's. Each time a person picks this year's row for one of last year's, a line goes to a local log: the kinds
of evidence that agreed on the row the finder had found (who it is, where it is, what it does, its numbers), its
confidence, whether the person's pick was that row, and the kinds of evidence for the row they picked. No label,
figure, sheet or file. From the log, how often each combination of evidence has been right: the confidence a
combination has earned, rather than one guessed (calibration). The diagnostics carry the table, so decisions made on
one machine can be learned from on another.
"""
import json
import time
from pathlib import Path

import workbench as wb


def log_file() -> Path:
    return wb.OUT / "evidence.jsonl"


def record(found_agreed: list | None, confidence: float | None, right: bool, picked_agreed: list | None,
           how: str | None = None) -> None:
    """One decision: the finder's row and what agreed on it, against the person's pick."""
    entry = {"at": round(time.time()), "agreed": sorted(found_agreed or []), "confidence": confidence,
             "right": bool(right), "picked": sorted(picked_agreed or []), "how": how}
    f = log_file()
    f.parent.mkdir(parents=True, exist_ok=True)
    with f.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(entry) + "\n")


def calibration() -> list[dict]:
    """[{"agreed": [kinds], "n", "right", "share"}]: how often the finder's row was the person's pick, by the kinds of
    evidence that agreed on it, most decisions first."""
    by = {}
    try:
        lines = log_file().read_text(encoding="utf-8").splitlines()
    except OSError:
        return []
    for line in lines:
        try:
            x = json.loads(line)
        except ValueError:
            continue
        key = tuple(x.get("agreed") or [])
        n, r = by.get(key, (0, 0))
        by[key] = (n + 1, r + bool(x.get("right")))
    return sorted(({"agreed": list(k), "n": n, "right": r, "share": round(r / n, 3)} for k, (n, r) in by.items()),
                  key=lambda x: -x["n"])
