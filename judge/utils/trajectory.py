"""Per-call trajectory recording for the judge.

Analogue of coding-agents' in-container API relay, but recorded at
the source: the judge owns its loop, so every record carries the exact wire
messages as sent, the parsed response, harness state the wire never shows
(pressure tier, evictions, tool executions, scratchpad outcomes), and
semantic tags (category / round / purpose) instead of raw HTTP.

Output: ``trajectory.jsonl`` inside judge_results/, written line-by-line and
flushed per record (crash-safe — a partial file is still kept with the rest
of judge_results/), gzipped to ``trajectory.jsonl.gz`` on close.

Record types (one JSON object per line, all carry "type" and "ts"):
  header  — run config: versions, rubric/weights md5s, check_order, model
  blob    — deduplicated large content: {"id", "chars", "content"}
  call    — one API call: mode/category/round/purpose, request (messages
            with big content replaced by {"$blob": id}, params, tool names),
            response, usage, latency_ms, error
  event   — harness happenings: pressure, tool_exec, nudge, eviction via
            tool_exec, api_retry, empty_choices, context_overflow,
            parse_failure, forced_finalization_start/end
  outcome — per-category result: judgement, coverage, pending, tool stats
  end     — totals written at close

Requests repeat their whole message prefix every round; the blob store keeps
each unique large string (CSV payloads, base64 images, tool results) once,
so repeated prefixes cost only short references. Reconstruction is a walk
replacing {"$blob": id} with the blob's content.
"""

import gzip
import hashlib
import json
import threading
import time
from pathlib import Path

# Strings at or above this many chars are stored once as a blob record and
# referenced by hash. Small strings stay inline for readability.
BLOB_THRESHOLD = 2048


def _jsonable(obj):
    """Best-effort conversion of SDK objects / dicts / lists to plain JSON."""
    if obj is None or isinstance(obj, (bool, int, float, str)):
        return obj
    if isinstance(obj, dict):
        return {str(k): _jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_jsonable(v) for v in obj]
    dump = getattr(obj, "model_dump", None)
    if callable(dump):
        try:
            return _jsonable(dump(mode="json"))
        except Exception:
            pass
    return str(obj)


class TrajectoryRecorder:
    """Appends trajectory records to a JSONL file; gzips on close.

    All methods are no-ops when ``enabled`` is False, so call sites never
    need to guard. Each grading run gets its own instance/file; the lock
    only defends against accidental cross-thread sharing.
    """

    def __init__(self, path, enabled: bool = True):
        self.path = Path(path)
        self.enabled = enabled
        self._fh = None
        self._lock = threading.Lock()
        self._step = 0
        self._blob_ids = set()
        self._closed = False
        if self.enabled:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self._fh = open(self.path, "w", encoding="utf-8")

    # ---- low-level ----

    def _write(self, record: dict) -> None:
        if not self.enabled or self._closed:
            return
        with self._lock:
            self._fh.write(json.dumps(record, ensure_ascii=False) + "\n")
            self._fh.flush()

    def _dedup(self, obj):
        """Replace large strings with {"$blob": id}, emitting blob records."""
        if isinstance(obj, str):
            if len(obj) >= BLOB_THRESHOLD:
                blob_id = hashlib.sha256(obj.encode("utf-8")).hexdigest()[:16]
                if blob_id not in self._blob_ids:
                    self._blob_ids.add(blob_id)
                    self._write(
                        {"type": "blob", "id": blob_id, "chars": len(obj), "content": obj}
                    )
                return {"$blob": blob_id, "chars": len(obj)}
            return obj
        if isinstance(obj, dict):
            return {k: self._dedup(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [self._dedup(v) for v in obj]
        return obj

    def _base(self, rtype: str) -> dict:
        self._step += 1
        return {"type": rtype, "step": self._step, "ts": round(time.time(), 3)}

    # ---- record types ----

    def record_header(self, **fields) -> None:
        rec = self._base("header")
        rec.update(_jsonable(fields))
        self._write(rec)

    def record_call(
        self,
        *,
        mode: str,
        category: str,
        round: int,
        purpose: str,
        model: str,
        request_messages,
        request_params: dict | None = None,
        tools: list | None = None,
        response=None,
        error: str | None = None,
        t0: float | None = None,
    ) -> None:
        if not self.enabled:
            return
        rec = self._base("call")
        response_j = _jsonable(response) if response is not None else None
        usage = None
        if isinstance(response_j, dict):
            usage = response_j.get("usage")
        rec.update(
            {
                "mode": mode,
                "category": category,
                "round": round,
                "purpose": purpose,
                "model": model,
                "request": {
                    "messages": self._dedup(_jsonable(request_messages)),
                    "params": _jsonable(request_params or {}),
                    "tools": tools or [],
                },
                "response": self._dedup(response_j),
                "usage": usage,
                # NB: the `round` parameter above shadows the builtin here
                "latency_ms": int((time.time() - t0) * 1000) if t0 else None,
                "error": error,
            }
        )
        self._write(rec)

    def record_event(self, kind: str, *, category: str | None = None, **data) -> None:
        if not self.enabled:
            return
        rec = self._base("event")
        rec["kind"] = kind
        rec["category"] = category
        rec.update(self._dedup(_jsonable(data)))
        self._write(rec)

    def record_outcome(self, *, category: str, **data) -> None:
        if not self.enabled:
            return
        rec = self._base("outcome")
        rec["category"] = category
        rec.update(self._dedup(_jsonable(data)))
        self._write(rec)

    # ---- lifecycle ----

    def close(self) -> str | None:
        """Write the end record, gzip the file, remove the plain JSONL."""
        if not self.enabled or self._closed:
            return None
        self._write({"type": "end", "steps": self._step + 1, "ts": round(time.time(), 3)})
        with self._lock:
            self._closed = True
            self._fh.close()
        gz_path = self.path.with_suffix(".jsonl.gz")
        with open(self.path, "rb") as src, gzip.open(gz_path, "wb") as dst:
            dst.write(src.read())
        self.path.unlink()
        return str(gz_path)
