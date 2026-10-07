#!/usr/bin/env python3
"""Trajectory-recording relay (runs INSIDE the sandbox container).

The coding CLI is pointed at http://127.0.0.1:$TRAJ_PORT via its base-URL env;
this relay forwards every call to $TRAJ_UPSTREAM (the vendor API) and appends
one record per call to $TRAJ_PATH:

  {"step", "ts", "method", "path", "request_headers", "request",
   "status", "response", "latency_ms"}

request/response are parsed JSON when possible; SSE streams are stored as raw
text. Auth headers are scrubbed. Streaming responses are passed through
chunk-by-chunk, so agent latency is unaffected. Stdlib only.
"""
import base64
import copy
import io
import json
import os
import re
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

UPSTREAM = os.environ["TRAJ_UPSTREAM"].rstrip("/")
OUT = os.environ.get("TRAJ_PATH", "/trajectory/trajectory.jsonl")
PORT = int(os.environ.get("TRAJ_PORT", "9877"))
# Seconds of upstream silence — connecting, waiting for the reply to start, or
# between body bytes — before the call counts as dropped. One hour for the
# vendor APIs. The TensorBlock Forge identities set 600 through their env
# (agent_identities.yaml): Forge has left calls unanswered for hours, and
# Codex has no timeout of its own while it waits for a reply to START
# (2026-09-21 probe: stream_idle_timeout_ms 15 s, still waiting at 150 s), so
# every such hang would sit here for the full hour.
UPSTREAM_TIMEOUT = float(os.environ.get("TRAJ_UPSTREAM_TIMEOUT") or 3600)
# An upstream 429 (rate limit) is retried HERE, with backoff, so the CLI never sees it:
# Codex treats one 429 as "exceeded retry limit" and ends the task (2026-09-22, two
# Kimi attempts lost that way when TensorBlock rate-limited five concurrent streams).
# Retry-After wins when the provider sends one (capped at 120 s); otherwise the delay
# doubles from TRAJ_RETRY_429_BASE seconds up to 60 s. 0 attempts = old behaviour.
RETRY_429_MAX = int(os.environ.get("TRAJ_RETRY_429_MAX") or 8)
RETRY_429_BASE = float(os.environ.get("TRAJ_RETRY_429_BASE") or 5)
# Request rewrites for providers whose Responses API is stricter than OpenAI's,
# switched on per identity by TRAJ_REQUEST_FIXES (a comma list in the
# identity's env, so every row records it). A record keeps the request as the
# CLI sent it and names the rewrites applied ("request_fixes").
#   reasoning_null_content  drop "content": null from the reasoning items Codex
#                           sends back: xAI (Grok through TensorBlock) answers
#                           400 to every call that carries one (2026-09-21).
REQUEST_FIXES = {f.strip() for f in os.environ.get("TRAJ_REQUEST_FIXES", "").split(",") if f.strip()}
KNOWN_FIXES = {"reasoning_null_content"}

SCRUB = {"authorization", "x-api-key", "cookie", "set-cookie", "openai-organization"}
HOP = {"transfer-encoding", "content-length", "connection", "keep-alive", "content-encoding"}

_lock = threading.Lock()
_step = [0]


def _jsonable(raw: bytes, content_type: str):
    if not raw:
        return None
    if "json" in (content_type or ""):
        try:
            return json.loads(raw)
        except (json.JSONDecodeError, UnicodeDecodeError):
            pass
    try:
        return {"_raw_text": raw.decode()}
    except UnicodeDecodeError:
        return {"_raw_b64": base64.b64encode(raw).decode()}


def rewrite_request(body: bytes):
    """(the body to forward, names of the rewrites applied to it)."""
    if not REQUEST_FIXES or not body:
        return body, []
    try:
        data = json.loads(body)
    except (json.JSONDecodeError, UnicodeDecodeError):
        return body, []
    items = data.get("input") if isinstance(data, dict) else None
    applied = set()
    if "reasoning_null_content" in REQUEST_FIXES and isinstance(items, list):
        for item in items:
            if (isinstance(item, dict) and item.get("type") == "reasoning"
                    and "content" in item and item["content"] is None):
                del item["content"]
                applied.add("reasoning_null_content")
    if not applied:
        return body, []
    return json.dumps(data).encode(), sorted(applied)


# ---------------------------------------------------------------------------
# Chat Completions wire (TRAJ_WIRE=chat, set per identity). Codex speaks only
# the Responses API, and some providers are only complete on TensorBlock's
# /v1/chat/completions (2026-09-21): on /v1/responses Fireworks cuts every
# Kimi K3 reply at 4,096 tokens and refuses its "max" tier, and Google's
# Gemini never gets back the thought signature it needs after each tool call.
# In chat mode each /v1/responses call becomes one streamed chat call and the
# reply is turned back into the Responses events Codex reads. Content is only
# re-packaged: the instructions and leading developer messages become the one
# system message, later developer messages user messages, function calls and
# their outputs assistant tool_calls and tool messages. Every assistant
# message is kept exactly as the provider returned it (Kimi's
# reasoning_content, Gemini's extra_content signatures) and sent back verbatim
# whenever Codex's history replays that turn. Records carry "wire": "chat",
# the chat request actually sent and the provider's raw stream.
WIRE = os.environ.get("TRAJ_WIRE", "responses")
KNOWN_WIRES = {"responses", "chat"}
# Reply cap sent as max_completion_tokens (the identity's TRAJ_CHAT_MAX_TOKENS):
# the model's own output limit, so the provider's default never cuts a reply.
CHAT_MAX_TOKENS = int(os.environ.get("TRAJ_CHAT_MAX_TOKENS") or 0)
_turns = {}  # turn number -> the assistant message exactly as returned
_turn_lock = threading.Lock()
_turn_seq = [0]
_last_tools = [None]  # the attempt's latest tool list, for requests that arrive without one
_OUR_ITEM = re.compile(r"^(?:rs|msg|fc)_tr(\d+)_\d+$")


def _text(content) -> str:
    if isinstance(content, str):
        return content
    return "\n\n".join(p.get("text", "") for p in content or []
                       if isinstance(p, dict) and p.get("type") in ("input_text", "output_text", "text"))


def _user_content(content):
    """A string, or a list of chat parts when the message carries images."""
    if isinstance(content, str):
        return content
    parts = []
    for p in content or []:
        if not isinstance(p, dict):
            continue
        if p.get("type") in ("input_text", "output_text", "text"):
            parts.append({"type": "text", "text": p.get("text", "")})
        elif p.get("type") == "input_image" and p.get("image_url"):
            parts.append({"type": "image_url", "image_url": {"url": p["image_url"]}})
    if all(p["type"] == "text" for p in parts):
        return "\n\n".join(p["text"] for p in parts)
    return parts


_CACHE = {"cache_control": {"type": "ephemeral"}}


def _mark_last_text(m):
    """cache_control on the message's last non-empty text block; a string content
    becomes one text block. False when there is nothing to mark (an empty block is refused)."""
    c = m.get("content")
    if isinstance(c, str) and c:
        m["content"] = [{"type": "text", "text": c, **_CACHE}]
        return True
    if isinstance(c, list):
        for i in range(len(c) - 1, -1, -1):
            if isinstance(c[i], dict) and c[i].get("type") == "text" and c[i].get("text"):
                c[i] = {**c[i], **_CACHE}
                return True
    return False


def _cache_breakpoints(msgs):
    """Anthropic prompt caching through Forge's chat route (2026-09-24, probed on the
    shared key). The route passes caching through, but only with explicit breakpoints,
    as Anthropic's own API: without them every call re-bills its whole context (the
    Opus 5 run: 96% input tokens, cached_tokens 0 on every reply). Breakpoints, four
    at most as Anthropic allows: the system message; the last user message when the
    request ends with one (first call, compaction); the last two assistant messages
    that carry text. Each call then reads the prefix its predecessor wrote and writes
    only the tail (probe: call 2 read 28,103 of 29,696 tokens). Never a tool message:
    a cache_control block inside a tool result is refused by one of Forge's two
    backends (400 "provider rejected", 5 of 6 tries) and a message-level marker is
    accepted but ignored. A marked assistant message is the replayed copy, its string
    content turned into one text block, which hashes the same (probed): the model sees
    the same text. Billing only. Returns the number of marks placed."""
    marked = 0
    if msgs and msgs[0].get("role") == "system" and _mark_last_text(msgs[0]):
        marked += 1
    if len(msgs) > 1 and msgs[-1].get("role") == "user" and _mark_last_text(msgs[-1]):
        marked += 1
    assistant_marks = 0
    for m in reversed(msgs[1:]):
        if assistant_marks >= 2 or marked >= 4:
            break
        if m.get("role") == "assistant" and _mark_last_text(m):
            marked += 1
            assistant_marks += 1
    return marked


def responses_to_chat(req: dict):
    """(chat request body, notes on anything that could not be carried over)."""
    notes = []
    items = req.get("input") or []
    if isinstance(items, str):
        items = [{"type": "message", "role": "user", "content": items}]
    system = [req["instructions"]] if req.get("instructions") else []
    k = 0
    while (k < len(items) and items[k].get("type", "message") == "message"
           and items[k].get("role") in ("system", "developer")):
        system.append(_text(items[k].get("content")))
        k += 1
    msgs = [{"role": "system", "content": "\n\n".join(s for s in system if s)}] if system else []
    call_ids = {it.get("call_id") for it in items if it.get("type") == "function_call"}
    replayed, loose = set(), None

    def flush():
        nonlocal loose
        if loose is not None:
            if not loose["tool_calls"]:
                del loose["tool_calls"]
            msgs.append(loose)
            loose = None

    for it in items[k:]:
        kind, role = it.get("type", "message"), it.get("role")
        if kind in ("reasoning", "function_call") or (kind == "message" and role == "assistant"):
            ours = _OUR_ITEM.match(str(it.get("id") or ""))
            turn = int(ours.group(1)) if ours else None
            with _turn_lock:
                orig = copy.deepcopy(_turns.get(turn)) if turn is not None else None
            if orig is not None:
                if turn not in replayed:
                    flush()
                    replayed.add(turn)
                    if orig.get("tool_calls"):  # only calls still in Codex's history
                        orig["tool_calls"] = [c for c in orig["tool_calls"] if c.get("id") in call_ids]
                        if not orig["tool_calls"]:
                            del orig["tool_calls"]
                    if orig.get("content") or orig.get("tool_calls"):
                        msgs.append(orig)
                continue
            if kind == "reasoning":
                notes.append("reasoning item from another source dropped")
                continue
            if loose is None:
                loose = {"role": "assistant", "content": None, "tool_calls": []}
            if kind == "message":
                loose["content"] = (loose["content"] or "") + _text(it.get("content"))
            else:
                loose["tool_calls"].append({"id": it.get("call_id"), "type": "function",
                                            "function": {"name": it.get("name"),
                                                         "arguments": it.get("arguments") or "{}"}})
            continue
        flush()
        if kind == "function_call_output":
            out = it.get("output")
            images = []
            if isinstance(out, list):
                parts = _user_content(out)
                if not isinstance(parts, str):
                    images = [p for p in parts if p["type"] == "image_url"]
                out = _text(out)
            msgs.append({"role": "tool", "tool_call_id": it.get("call_id"), "content": out or ""})
            if images:
                msgs.append({"role": "user", "content": images})
        elif kind == "message":
            msgs.append({"role": "user" if role in ("system", "developer", None) else role,
                         "content": _user_content(it.get("content"))})
        else:
            notes.append(f"input item of type {kind} dropped")
    flush()

    body = {"model": req.get("model"), "messages": msgs, "stream": True,
            "stream_options": {"include_usage": True}}
    tools = []
    for tool in req.get("tools") or []:
        if tool.get("type") != "function":
            notes.append(f"tool of type {tool.get('type')} dropped")
            continue
        fn = {"name": tool.get("name"),
              "parameters": tool.get("parameters") or {"type": "object", "properties": {}}}
        if tool.get("description"):
            fn["description"] = tool["description"]
        tools.append({"type": "function", "function": fn})
    if tools:
        body["tools"] = tools
        choice = req.get("tool_choice")
        if isinstance(choice, str):
            body["tool_choice"] = choice
        elif isinstance(choice, dict) and choice.get("name"):
            body["tool_choice"] = {"type": "function", "function": {"name": choice["name"]}}
        if "parallel_tool_calls" in req:
            body["parallel_tool_calls"] = req["parallel_tool_calls"]
        with _turn_lock:
            _last_tools[0] = copy.deepcopy(tools)
    elif any(m.get("tool_calls") or m.get("role") == "tool" for m in msgs):
        # 2026-09-24: Codex's context compaction ("CONTEXT CHECKPOINT COMPACTION", at
        # ~245k tokens) sends the whole history, tool calls and results included, with
        # no tools. Bedrock refuses a history that uses tools without a tool list
        # (Forge: 400 "The configured provider rejected the request"); Codex retried 15
        # times and the attempt would have failed (Opus 5 task 13). One relay serves one
        # attempt, so its own last tool list goes back, with no tool_choice: Bedrock
        # has no "none" (400), and "auto" is the default.
        with _turn_lock:
            last = copy.deepcopy(_last_tools[0])
        if last:
            body["tools"] = last
            notes.append("tools re-attached: history has tool calls, request had no tools")
    effort = (req.get("reasoning") or {}).get("effort")
    if effort:
        body["reasoning_effort"] = effort
    if "claude" in str(req.get("model") or "").lower():  # the chat wire is Forge-only
        # 2026-09-23: Forge sends no byte while a Claude model thinks and cuts its
        # upstream connection after ~600 s of silence ("The model request timed out
        # before completion" after exactly 40 keep-alive pings; 214 of 641 Fable-max
        # calls on the Stage 5 v15 launch, every long think). Asking for the thinking
        # summary makes Bedrock stream summary deltas while the model thinks, so the
        # connection stays busy and a 25-minute think completes - the cli pipeline's
        # fix (task_executor._forge_claude_extra_body, commit 9bd2ac1). Display only:
        # reasoning depth, output and billing are unchanged. Claude on Forge only.
        body["thinking"] = {"type": "adaptive", "display": "summarized"}
        marks = _cache_breakpoints(msgs)
        if marks:
            notes.append(f"cache_control on {marks} block(s)")
    if CHAT_MAX_TOKENS:
        body["max_completion_tokens"] = CHAT_MAX_TOKENS
    return body, notes


class ChatStream:
    """Turns one streamed chat completion into the Responses events Codex
    reads (the same event shapes Forge's own /v1/responses sends)."""

    def __init__(self, emit, model, turn):
        self.emit, self.model, self.turn = emit, model, turn
        self.rid, self.created, self.seq = f"resp_tr{turn}", int(time.time()), 0
        self.items = []  # output items in order of first appearance
        self.reasoning = self.message = None
        self.calls = {}  # chat tool-call index -> accumulated call
        self.extra = {}  # other assistant-message fields the provider sent
        self.usage = self.finish_reason = self.error = None

    def _ev(self, kind, **fields):
        self.seq += 1
        self.emit(kind, {"type": kind, **fields, "sequence_number": self.seq})

    def _response(self, status, **fields):
        return {"id": self.rid, "object": "response", "created_at": self.created, "status": status,
                "model": self.model, "output": [], **fields}

    def start(self):
        self._ev("response.created", response=self._response("in_progress"))

    def _open(self, kind):
        entry = {"kind": kind, "index": len(self.items), "text": "",
                 "id": f"{'rs' if kind == 'reasoning' else 'msg'}_tr{self.turn}_{len(self.items)}"}
        self.items.append(entry)
        if kind == "reasoning":
            self._ev("response.output_item.added", output_index=entry["index"],
                     item={"id": entry["id"], "type": "reasoning", "summary": [], "content": []})
            self._ev("response.reasoning_summary_part.added", item_id=entry["id"], output_index=entry["index"],
                     summary_index=0, part={"type": "summary_text", "text": ""})
        else:
            self._ev("response.output_item.added", output_index=entry["index"],
                     item={"id": entry["id"], "type": "message", "role": "assistant",
                           "status": "in_progress", "content": []})
            self._ev("response.content_part.added", item_id=entry["id"], output_index=entry["index"],
                     content_index=0, part={"type": "output_text", "annotations": [], "logprobs": [], "text": ""})
        return entry

    def feed(self, chunk: dict):
        if chunk.get("error"):
            self.error = chunk["error"]
            return
        if chunk.get("usage"):
            self.usage = chunk["usage"]
        for choice in chunk.get("choices") or []:
            self.finish_reason = choice.get("finish_reason") or self.finish_reason
            delta = choice.get("delta") or {}
            thought = delta.get("reasoning_content") or delta.get("reasoning")
            if isinstance(thought, str) and thought:
                if self.reasoning is None:
                    self.reasoning = self._open("reasoning")
                self.reasoning["text"] += thought
                self._ev("response.reasoning_summary_text.delta", item_id=self.reasoning["id"],
                         output_index=self.reasoning["index"], summary_index=0, delta=thought)
            text = delta.get("content")
            if isinstance(text, str) and text:
                if self.message is None:
                    self.message = self._open("message")
                self.message["text"] += text
                self._ev("response.output_text.delta", item_id=self.message["id"],
                         output_index=self.message["index"], content_index=0, delta=text, logprobs=[])
            for call in delta.get("tool_calls") or []:
                acc = self.calls.setdefault(call.get("index", len(self.calls)),
                                            {"id": None, "name": "", "arguments": "", "extra": {}})
                acc["id"] = call.get("id") or acc["id"]
                fn = call.get("function") or {}
                if fn.get("name") and not acc["name"]:
                    acc["name"] = fn["name"]
                acc["arguments"] += fn.get("arguments") or ""
                acc["extra"].update({k: v for k, v in call.items()
                                     if k not in ("index", "id", "type", "function") and v is not None})
            self.extra.update({k: v for k, v in delta.items()
                               if k not in ("role", "content", "reasoning_content", "reasoning", "tool_calls")
                               and v is not None})

    def _usage(self):
        u = self.usage or {}
        prompt, completion = u.get("prompt_tokens") or 0, u.get("completion_tokens") or 0
        # Thinking a provider bills but leaves out of completion_tokens (Gemini).
        hidden = max(0, (u.get("total_tokens") or 0) - prompt - completion)
        reasoning = ((u.get("completion_tokens_details") or {}).get("reasoning_tokens")
                     or (u.get("output_tokens_details") or {}).get("reasoning_tokens") or 0)
        return {"input_tokens": prompt,
                "input_tokens_details": {"cached_tokens": (u.get("prompt_tokens_details") or {}).get("cached_tokens") or 0},
                "output_tokens": completion + hidden,
                "output_tokens_details": {"reasoning_tokens": reasoning + hidden},
                "total_tokens": prompt + completion + hidden}

    def _outcome(self):
        """("completed" | "incomplete" | "failed", detail) for the whole reply."""
        if self.error is not None:
            err = self.error if isinstance(self.error, dict) else {"message": str(self.error)}
            return "failed", {"code": str(err.get("code") or "server_error"),
                              "message": str(err.get("message") or err)[:2000]}
        if self.finish_reason == "length":
            return "incomplete", {"reason": "max_output_tokens"}
        if self.finish_reason == "content_filter":
            return "incomplete", {"reason": "content_filter"}
        if self.finish_reason is None:
            # The provider closed the stream before its finish_reason chunk: a
            # dropped call, whatever had already arrived. Passed on as
            # "completed", a reply of reasoning alone made Codex end the task
            # (2026-09-23, Kimi task 43: 471 s of thinking, stream cut, no
            # answer, no tool call, exit 0, no workbook). Failed → Codex retries.
            partial = bool(self.reasoning or self.message or self.calls)
            return "failed", {"code": "server_error", "message": "upstream stream ended without a finish_reason"
                              + (" (partial reply discarded)" if partial else " (no reply)")}
        if self.finish_reason not in (None, "stop", "tool_calls", "function_call"):
            # The provider stopped the reply for its own reason, e.g. Gemini's
            # "function_call_filter: MALFORMED_FUNCTION_CALL" (a tool call it
            # could not parse, sent back as plain text): passed on as a
            # finished answer, Codex would end the task there (2026-09-21,
            # Gemini light test).
            return "failed", {"code": "server_error",
                              "message": f"provider stopped the reply: {self.finish_reason}"}
        return "completed", None

    def finish(self):
        status, detail = self._outcome()
        if status != "completed":
            # No item of a reply that did not finish is completed: Codex adds
            # completed items to its history, and the retry it makes must not
            # carry a cut-off or garbled turn.
            if status == "failed":
                self._ev("response.failed", response=self._response("failed", error=detail))
            else:
                self._ev("response.incomplete", response=self._response(
                    "incomplete", incomplete_details=detail, usage=self._usage()))
            return
        output = []
        if self.reasoning:
            r = self.reasoning
            part = {"type": "summary_text", "text": r["text"]}
            self._ev("response.reasoning_summary_text.done", item_id=r["id"], output_index=r["index"],
                     summary_index=0, text=r["text"])
            self._ev("response.reasoning_summary_part.done", item_id=r["id"], output_index=r["index"],
                     summary_index=0, part=part)
            item = {"id": r["id"], "type": "reasoning", "summary": [part], "content": []}
            self._ev("response.output_item.done", output_index=r["index"], item=item)
            output.append(item)
        if self.message:
            m = self.message
            part = {"type": "output_text", "annotations": [], "logprobs": [], "text": m["text"]}
            self._ev("response.output_text.done", item_id=m["id"], output_index=m["index"], content_index=0,
                     text=m["text"], logprobs=[])
            self._ev("response.content_part.done", item_id=m["id"], output_index=m["index"], content_index=0, part=part)
            item = {"id": m["id"], "type": "message", "role": "assistant", "status": "completed",
                    "content": [{"type": "output_text", "text": m["text"], "annotations": []}]}
            self._ev("response.output_item.done", output_index=m["index"], item=item)
            output.append(item)
        tool_calls = []
        for n in sorted(self.calls):
            acc = self.calls[n]
            call_id = acc["id"] or f"call_tr{self.turn}_{n}"
            index = len(self.items) + len(tool_calls)
            item = {"id": f"fc_tr{self.turn}_{index}", "type": "function_call", "status": "completed",
                    "call_id": call_id, "name": acc["name"], "arguments": acc["arguments"] or "{}"}
            self._ev("response.output_item.done", output_index=index, item=item)
            output.append(item)
            tool_calls.append({"id": call_id, "type": "function",
                               "function": {"name": acc["name"], "arguments": acc["arguments"] or "{}"},
                               **acc["extra"]})
        message = {"role": "assistant", "content": (self.message or {}).get("text") or None, **self.extra}
        if tool_calls:
            message["tool_calls"] = tool_calls
        if self.reasoning:
            message["reasoning_content"] = self.reasoning["text"]
        with _turn_lock:
            _turns[self.turn] = message
        self._ev("response.completed", response=self._response("completed", output=output, usage=self._usage()))


class _Replayed:
    """An upstream error reply whose body was read to inspect it, handed on unchanged."""

    def __init__(self, body: bytes, resp):
        self.code, self.headers, self._buf = resp.code, resp.headers, io.BytesIO(body)

    def read(self, n=-1):
        return self._buf.read(n)

    def readline(self):
        return self._buf.readline()

    def close(self):
        pass


def record(entry: dict):
    with _lock:
        _step[0] += 1
        entry["step"] = _step[0]
        with open(OUT, "a") as f:
            f.write(json.dumps(entry) + "\n")


class Relay(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):  # quiet; relay.log carries errors only
        pass

    def _handle(self):
        t0 = time.time()
        length = int(self.headers.get("content-length") or 0)
        body = self.rfile.read(length) if length else b""

        fwd_headers = {k: v for k, v in self.headers.items()
                       if k.lower() not in ("host", "content-length", "accept-encoding")}
        fwd_headers["accept-encoding"] = "identity"
        fwd_body, fixes = rewrite_request(body)
        path = self.path.split("?")[0].rstrip("/")
        chat = WIRE == "chat" and self.command == "POST" and path.endswith("/responses")
        url, chat_body, notes, sent = UPSTREAM + self.path, None, [], []
        chunks = []
        retries = []  # upstream 429s answered here before the final reply
        # Where a failure happened, so a status -1 record says WHO dropped the
        # connection: "upstream_open" (no response headers ever arrived),
        # "upstream_read" (the API side closed or errored mid-body — the
        # client was still connected) or "client_write" (Claude Code / Codex
        # closed its side while the upstream body was still flowing, e.g. its
        # own timeout or watchdog). Before 2026-09-11 the exception was
        # dropped and every drop looked the same. "translate": the chat-wire
        # translation itself failed (answered like a failed connection).
        phase = "upstream_open"
        error = None
        headers_sent = False
        try:
            if chat:
                phase = "translate"
                chat_body, notes = responses_to_chat(json.loads(fwd_body))
                url = UPSTREAM + path[: -len("/responses")] + "/chat/completions"
                fwd_body = json.dumps(chat_body).encode()
                with _turn_lock:
                    _turn_seq[0] += 1
                    turn = _turn_seq[0]
                phase = "upstream_open"
            while True:
                req = urllib.request.Request(url, data=fwd_body if fwd_body else None,
                                             headers=fwd_headers, method=self.command)
                try:
                    resp = urllib.request.urlopen(req, timeout=UPSTREAM_TIMEOUT)
                except urllib.error.HTTPError as e:
                    resp = e
                status = resp.code
                quota = False
                if status == 400:
                    # 2026-09-24: Forge passes a Bedrock throttle through as a 400
                    # provider_error "The configured provider rate limit, quota, or
                    # billing limit was reached" - a 429 in all but name, lasting ~10 s
                    # (11:03 and 12:44 bursts: every in-flight call on the box, 24
                    # attempts lost, Codex treats a 400 as final). Retried here like a
                    # 429; any other 400 is handed on unchanged.
                    peek = resp.read()
                    resp.close()
                    quota = b"rate limit, quota, or billing limit" in peek
                    resp = _Replayed(peek, resp)
                if not (status == 429 or quota) or len(retries) >= RETRY_429_MAX:
                    break
                snippet = resp.read(500).decode(errors="replace")
                resp.close()
                delay = min(RETRY_429_BASE * (2 ** len(retries)), 60.0)
                try:
                    if resp.headers.get("retry-after"):
                        delay = min(float(resp.headers.get("retry-after")), 120.0)
                except ValueError:
                    pass
                retries.append({"ts": datetime.now(timezone.utc).isoformat(), "status": status,
                                "delay_s": round(delay, 2), "body": snippet})
                phase = "upstream_retry"
                time.sleep(delay)
                phase = "upstream_open"
            rheaders = list(resp.headers.items())
            if chat and status == 200:
                self.send_response(200)
                self.send_header("Content-Type", "text/event-stream")
                self.send_header("Transfer-Encoding", "chunked")
                self.end_headers()
                headers_sent = True

                def emit(kind, payload):
                    data = f"event: {kind}\ndata: {json.dumps(payload)}\n\n".encode()
                    sent.append(data)
                    self.wfile.write(f"{len(data):X}\r\n".encode() + data + b"\r\n")
                    self.wfile.flush()

                stream = ChatStream(emit, chat_body.get("model"), turn)
                phase = "client_write"
                stream.start()
                while True:
                    phase = "upstream_read"
                    line = resp.readline()
                    if not line:
                        break
                    chunks.append(line)
                    data = line.strip()
                    if not data.startswith(b"data:"):
                        continue
                    data = data[5:].strip()
                    if data == b"[DONE]":
                        break
                    try:
                        obj = json.loads(data)
                    except ValueError:
                        continue
                    phase = "client_write"
                    stream.feed(obj)
                resp.close()
                phase = "client_write"
                stream.finish()
                self.wfile.write(b"0\r\n\r\n")
            else:
                self.send_response(status)
                for k, v in rheaders:
                    if k.lower() not in HOP:
                        self.send_header(k, v)
                self.send_header("Transfer-Encoding", "chunked")
                self.end_headers()
                headers_sent = True
                while True:
                    phase = "upstream_read"
                    chunk = resp.read(8192)
                    if not chunk:
                        break
                    chunks.append(chunk)
                    phase = "client_write"
                    self.wfile.write(f"{len(chunk):X}\r\n".encode() + chunk + b"\r\n")
                    self.wfile.flush()
                phase = "client_write"
                self.wfile.write(b"0\r\n\r\n")
        except Exception as e:  # noqa: BLE001 — recorded, then propagated as 502
            error = {"phase": phase, "type": type(e).__name__, "repr": repr(e)[:500],
                     "bytes_relayed": sum(len(c) for c in chunks),
                     "elapsed_ms": round((time.time() - t0) * 1000)}
            try:
                with open(OUT + ".errors.log", "a") as f:
                    f.write(json.dumps({"ts": datetime.now(timezone.utc).isoformat(), **error}) + "\n")
            except Exception:
                pass
            try:
                if headers_sent:
                    # Mid-body failure, unchanged: the status line lands inside
                    # the chunked body, the client's parser rejects it and the
                    # CLI retries (272 such drops on record, none hung).
                    self.send_response(502)
                    self.end_headers()
                else:
                    # The upstream connection failed before any reply. This 502
                    # must carry a length AND end the connection: a bare 502 on
                    # a keep-alive HTTP/1.1 socket leaves the client waiting
                    # forever for a body. 2026-09-20: Claude Code sat 68 min on
                    # one (task 84; with API_FORCE_IDLE_TIMEOUT=0 no runtime
                    # timeout breaks the wait) until the run was stopped by hand.
                    # A complete 502 is an ordinary retryable error to both CLIs.
                    self.close_connection = True
                    self.send_response(502)
                    self.send_header("Content-Length", "0")
                    self.send_header("Connection", "close")
                    self.end_headers()
            except Exception:
                pass
            status = -1
            rheaders = [("x-relay-error", repr(e))]

        raw = b"".join(chunks)
        ct = dict((k.lower(), v) for k, v in rheaders).get("content-type", "")
        entry = {
            "ts": datetime.now(timezone.utc).isoformat(),
            "method": self.command,
            "path": self.path,
            "request_headers": {k.lower(): ("<scrubbed>" if k.lower() in SCRUB else v)
                                for k, v in self.headers.items()},
            "request": _jsonable(body, self.headers.get("content-type", "")),
            "status": status,
            "response": _jsonable(b"".join(sent), "text/event-stream") if sent else _jsonable(raw, ct),
            "latency_ms": round((time.time() - t0) * 1000),
            "error": error,
        }
        if fixes:
            entry["request_fixes"] = fixes
        if retries:
            entry["upstream_retries"] = retries
        if chat:
            entry["wire"] = "chat"
            entry["upstream_request"] = chat_body
            if sent:  # what the provider streamed, before translation
                entry["upstream_response"] = _jsonable(raw, ct)
            if notes:
                entry["translation_notes"] = notes
        record(entry)

    do_POST = do_GET = do_PUT = do_DELETE = _handle


if __name__ == "__main__":
    unknown = sorted(REQUEST_FIXES - KNOWN_FIXES)
    if unknown:  # a misspelt fix must not run silently; the entrypoint then aborts (97, no row)
        raise SystemExit(f"traj_relay: unknown TRAJ_REQUEST_FIXES {unknown}; known: {sorted(KNOWN_FIXES)}")
    if WIRE not in KNOWN_WIRES:
        raise SystemExit(f"traj_relay: unknown TRAJ_WIRE {WIRE!r}; known: {sorted(KNOWN_WIRES)}")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Relay)
    print(f"traj_relay: 127.0.0.1:{PORT} -> {UPSTREAM}, recording to {OUT}", flush=True)
    server.serve_forever()
