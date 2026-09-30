"""Stay a safety margin below each deployment's rate limits instead of hitting Azure's 429s.

Standard deployments get capacity x 1,000 tokens/min and a per-model requests/min ratio. Before each call
we estimate its tokens and wait until the last 60 seconds of calls to that deployment leave room under
(1 - BUFFER) of both limits. Actual token counts replace the estimate once the response arrives.
BUFFER defaults to 10%; set RATE_LIMIT_BUFFER=0.05 for 5%.
"""
import os
import threading
import time
from collections import deque

BUFFER = float(os.environ.get("RATE_LIMIT_BUFFER", "0.10"))
WINDOW = 60.0

# Requests/min per 1,000 tokens/min, from Microsoft's quota tables (learn.microsoft.com, updated 2026-08-20).
# gpt-4o: 6, confirmed from the deployment's limits in the portal (450,000 TPM / 2,700 RPM, 2026-09-25).
RPM_PER_1K_TPM = {"gpt-4o": 6, "gpt-4o-mini": 10, "gpt-5-nano": 1, "gpt-6-sol": 1, "gpt-6-luna": 1}
# Capacity (x 1,000 tokens/min) as deployed on 2026-09-25; refreshed from the project by load_capacities().
CAPACITY = {"gpt-4o": 450, "gpt-4o-mini": 200, "gpt-5-nano": 200, "gpt-6-sol": 500, "gpt-6-luna": 500}
SOURCE = "built-in defaults"

_lock = threading.Condition()
_calls: dict[str, deque] = {}  # model -> deque of [timestamp, tokens]


def limits(model: str) -> dict | None:
    cap = CAPACITY.get(model)
    if not cap:
        return None
    tpm, rpm = cap * 1000, cap * RPM_PER_1K_TPM.get(model, 1)
    return {"tpm": tpm, "rpm": rpm, "tpm_budget": int(tpm * (1 - BUFFER)), "rpm_budget": max(1, int(rpm * (1 - BUFFER)))}


def load_capacities() -> str:
    """Read each deployment's capacity from the Foundry project (so raising it in the portal takes effect)."""
    global SOURCE
    try:
        from azure.ai.projects import AIProjectClient
        from llm import PROJECT_ENDPOINT, credential
        project = AIProjectClient(endpoint=PROJECT_ENDPOINT, credential=credential(interactive=False))
        for d in project.deployments.list():
            x = d.as_dict()
            if x.get("sku", {}).get("capacity"):
                CAPACITY[x["name"]] = x["sku"]["capacity"]
        SOURCE = f"your Foundry project, read {time.strftime('%H:%M')}"
    except Exception as e:
        SOURCE = f"built-in defaults (couldn't read the project: {type(e).__name__})"
    return SOURCE


def estimate_tokens(*parts, reserve_output: int = 1000) -> int:
    """Rough prompt size (4 characters per token) plus room for the reply, which Azure also counts."""
    return sum(len(str(p)) for p in parts) // 4 + reserve_output


def acquire(model: str, est_tokens: int, on_wait=None) -> list | None:
    """Block until a call of est_tokens fits under the budget. Returns a handle for settle()."""
    lim = limits(model)
    if lim is None:
        return None
    told = False
    with _lock:
        q = _calls.setdefault(model, deque())
        while True:
            now = time.time()
            while q and now - q[0][0] > WINDOW:
                q.popleft()
            used = sum(t for _, t in q)
            # A single call bigger than the whole budget goes once the window is empty.
            fits_tokens = used + est_tokens <= lim["tpm_budget"] or not q
            if len(q) < lim["rpm_budget"] and fits_tokens:
                entry = [now, est_tokens]
                q.append(entry)
                return entry
            wait = max(0.2, WINDOW - (now - q[0][0]))
            if on_wait and not told:
                on_wait(round(wait), lim)
                told = True
            _lock.wait(timeout=min(wait, 5))


def wait_needed(model: str, est_tokens: int) -> float:
    """Seconds until a call of est_tokens would fit (0 if it fits now). Doesn't reserve anything."""
    lim = limits(model)
    if lim is None:
        return 0.0
    with _lock:
        now = time.time()
        q = [e for e in _calls.get(model, ()) if now - e[0] <= WINDOW]
        used = sum(t for _, t in q)
        if not q or (len(q) < lim["rpm_budget"] and used + est_tokens <= lim["tpm_budget"]):
            return 0.0
        return max(0.0, WINDOW - (now - q[0][0]))


def settle(entry: list | None, actual_tokens: int) -> None:
    """Swap the estimate for the real total once the response reports it."""
    if entry is not None:
        with _lock:
            entry[1] = actual_tokens
            _lock.notify_all()


def status() -> dict:
    now = time.time()
    out = {}
    with _lock:
        for m in CAPACITY:
            if m not in RPM_PER_1K_TPM:
                continue  # embeddings / transcription aren't used by the app
            lim = limits(m)
            q = [e for e in _calls.get(m, ()) if now - e[0] <= WINDOW]
            out[m] = {**lim, "tokens_last_min": sum(t for _, t in q), "requests_last_min": len(q)}
    return {"buffer": BUFFER, "source": SOURCE, "models": out}
