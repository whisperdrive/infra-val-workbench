"""A person's decisions on disk (acknowledgements, row picks, figures set for held inputs, this year's discount rate and
valuation date, the method, confirmed terms, the equity cells, the cells held at Excel's value), read and written so
that none is lost:

  one at a time   each file has its own lock; a change is read, made and written under it (update), so two decisions
                  made at once both land (a page's two quick clicks, an agents' run writing beside a person)
  whole           written to a temporary file beside it, then swapped in (os.replace): a crash leaves the old file or
                  the new one, never half of one, and a reader never sees it half written. On Windows a reader holding
                  the file open for a moment is waited out
  never empty     a file that doesn't parse isn't read as empty, which would drop every decision in it without a word:
                  it's moved aside (<name>.damaged-<time>), kept, and the engagement's value holds until a person has
                  looked (damaged(), result.compute)
"""
import json
import os
import threading
import time
from pathlib import Path

_LOCKS: dict[str, threading.RLock] = {}
_GUARD = threading.Lock()
TRIES = 40      # Windows: a file another reader (antivirus, an indexer, a sync client) has open can't be replaced or
WAIT = 0.05     # opened for a moment: tried again, the wait doubling from WAIT up to MAX_WAIT (about 15 s in all)
MAX_WAIT = 0.5


def lock(path) -> threading.RLock:
    key = os.path.normcase(os.path.abspath(str(path)))
    with _GUARD:
        return _LOCKS.setdefault(key, threading.RLock())


def _retry(fn):
    wait = WAIT
    for i in range(TRIES):
        try:
            return fn()
        except PermissionError:
            if i == TRIES - 1:
                raise
            time.sleep(wait)
            wait = min(wait * 2, MAX_WAIT)


def read(path, default=None):
    """The file's contents, or default where there's no file. A file that doesn't parse is moved aside, kept, and
    default returned: damaged() says so, and the value holds until a person has looked."""
    p = Path(path)
    if not p.exists():
        return default
    try:
        return json.loads(_retry(lambda: p.read_text(encoding="utf-8")))
    except FileNotFoundError:
        return default
    except ValueError:
        with lock(p):  # still damaged once no one is writing it: moved aside
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except FileNotFoundError:
                return default
            except ValueError:
                aside = p.with_name(f"{p.name}.damaged-{time.strftime('%Y%m%d-%H%M%S')}")
                _retry(lambda: os.replace(p, aside))
                return default


def write(path, data) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with lock(p):
        tmp = p.with_name(f".{p.name}.{os.getpid()}.{threading.get_ident()}.tmp")
        tmp.write_text(json.dumps(data, indent=1, default=str), encoding="utf-8")
        try:
            _retry(lambda: os.replace(tmp, p))
        except OSError:
            tmp.unlink(missing_ok=True)  # the old file stands; the half-step isn't left behind
            raise


def update(path, change, default=None):
    """change(contents) -> the new contents (or None to leave the file as it is), under the file's lock: two changes at
    once both land. Returns the contents as written (or as they were)."""
    with lock(path):
        got = read(path, default)
        new = change(got)
        if new is None:
            return got
        write(path, new)
        return new


def damaged(folder) -> list[dict]:
    """The decision files moved aside as damaged in this folder: [{"file", "aside"}], oldest first."""
    try:
        xs = sorted(Path(folder).glob("*.damaged-*"))
    except OSError:
        return []
    return [{"file": x.name.split(".damaged-")[0], "aside": x.name} for x in xs]
