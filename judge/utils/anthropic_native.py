"""Native Anthropic Messages-API adapter for the judge's tool loop (judge v6).

Why this exists (2026-09-02): the judge has always reached Claude through
Anthropic's OpenAI-compatibility endpoint, which — per Anthropic's own
compatibility page — does NOT support prompt caching, IGNORES
`reasoning_effort`, and returns empty `prompt_tokens_details`. So on that
path (a) every round re-bills the whole conversation at full price (the
sonnet canary: $21/grading at 100% fresh input vs sol's $9.46 at 92%
cached), and (b) the "low"/"max" effort arms were never real. Both are
fixed by speaking the native API: `output_config.effort` for real effort
tiers, `cache_control` markers for the ~90% cache-read discount, and
`usage.cache_read_input_tokens` for honest cost accounting.

Same design as utils/openai_responses.py: the loop keeps speaking
chat-completions shapes; this module translates both directions and
returns a shim that quacks like a ChatCompletion (real SDK
ChatCompletionMessage / tool-call objects; CompletionUsage whose
prompt_tokens is the TOTAL input — uncached + cache-read + cache-write — so
the loop's chars-per-token calibration and read gate stay correct, with
cached tokens in prompt_tokens_details.cached_tokens and the cache-write
count as `cache_write_tokens`).

Thinking blocks: Claude 5 models think adaptively and require an assistant
turn that contains tool_use to be replayed WITH its thinking blocks intact.
The adapter keeps each assistant turn's native content blocks in a
side-table keyed by the turn's first tool-call id and replays them
verbatim; eviction (dropping whole tool rounds from the wire) simply stops
emitting those turns. Text-only assistant turns replay as text.

Caching layout: one explicit breakpoint on the (hoisted, frozen) system
prompt — which caches tools + system together — plus the top-level
automatic `cache_control` that walks forward with the growing conversation.
Eviction rewrites the history at the evicted position, so the tail re-caches
from there; the seed prefix stays cached throughout.

Scope: wired into single_pass_judge_case; other providers keep the
compat-endpoint client byte-for-byte.
"""

from __future__ import annotations

import base64
import json
import os
import re

from openai.types.chat import ChatCompletionMessage
from openai.types.chat.chat_completion_message_function_tool_call import (
    ChatCompletionMessageFunctionToolCall,
    Function,
)
from openai.types.completion_usage import CompletionUsage, PromptTokensDetails

try:
    from .logger import logger
except ImportError:  # imported as a bare module (utils/ on sys.path)
    from logger import logger

# 128K = the models' own output maximum (Sonnet 5 / Opus 5). The old 32K was
# a hand-set cap inherited from the compat path (raised from 16K in August
# when Opus 5 thinking truncated a JSON reply); the rung-1 handshake at
# effort=max hit exactly 32,000 completion tokens and lost the turn's tool
# calls. max_tokens is a ceiling, not a spend — only generated tokens bill —
# and requests stream (see create()) so a long turn cannot trip an HTTP
# timeout. The OpenAI paths set no output cap (model default = 128K).
DEFAULT_MAX_TOKENS = 128000
EFFORT_LEVELS = ("low", "medium", "high", "xhigh", "max")
_OK_IMAGE_MEDIA = {"image/jpeg", "image/png", "image/gif", "image/webp"}
_DATA_URL_RE = re.compile(r"^data:(?P<mime>[^;]+);base64,(?P<data>.+)$", re.S)


def wants_native_anthropic(provider: str) -> bool:
    """Every provider=anthropic call in single-pass mode routes here."""
    return provider == "anthropic"


def map_effort(reasoning_effort):
    """Judge effort label -> Anthropic output_config.effort (or None = omit).

    The registry's null effort means "the model's default" (omit). The
    OpenAI-only labels 'none'/'minimal' have no Anthropic equivalent; they
    map to 'low' (the floor) and are logged, so a mis-pinned identity is
    visible rather than silently defaulting to high.
    """
    if reasoning_effort in (None, ""):
        return None
    e = str(reasoning_effort).lower()
    if e in EFFORT_LEVELS:
        return e
    if e in ("none", "minimal"):
        logger.warning(
            f"  [anthropic-native] effort {reasoning_effort!r} is not an "
            f"Anthropic tier; using 'low' (the floor)"
        )
        return "low"
    raise ValueError(f"unknown reasoning effort {reasoning_effort!r} for Anthropic")


class NativeState:
    """Side-table: first tool-call id of an assistant turn -> that turn's
    native content blocks (plain dicts, thinking included) for verbatim
    replay. `strip_thinking` is the one-shot recovery for a prefix-binding
    400 (see create())."""

    def __init__(self):
        self.by_first_call_id: dict[str, list[dict]] = {}
        self.strip_thinking: bool = False


def _msg_get(m, key, default=None):
    if isinstance(m, dict):
        return m.get(key, default)
    return getattr(m, key, default)


def convert_tools(chat_tools: list[dict]) -> list[dict]:
    """Chat-completions function tools -> Anthropic tool definitions."""
    out = []
    for t in chat_tools:
        fn = t["function"]
        out.append(
            {
                "name": fn["name"],
                "description": fn.get("description", ""),
                "input_schema": fn.get("parameters") or {"type": "object", "properties": {}},
            }
        )
    return out


def _text_block(text: str) -> dict | None:
    if text is None or str(text) == "":
        return None
    return {"type": "text", "text": str(text)}


def _content_to_blocks(content) -> list[dict]:
    """Chat user/system content (str or parts list) -> Anthropic blocks."""
    if isinstance(content, str):
        b = _text_block(content)
        return [b] if b else []
    blocks = []
    for part in content or []:
        if not isinstance(part, dict):
            continue
        ptype = part.get("type")
        if ptype == "text" or ("text" in part and ptype is None):
            b = _text_block(part.get("text", ""))
            if b:
                blocks.append(b)
        elif ptype == "image_url":
            url = (part.get("image_url") or {}).get("url", "")
            m = _DATA_URL_RE.match(url or "")
            if not m or m.group("mime") not in _OK_IMAGE_MEDIA:
                continue  # non-data or unsupported image: drop (compat did too)
            blocks.append(
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": m.group("mime"),
                        "data": m.group("data"),
                    },
                }
            )
        elif ptype == "image":
            src = part.get("source") or {}
            if src.get("media_type") in _OK_IMAGE_MEDIA:
                blocks.append({"type": "image", "source": src})
    return blocks


def _assistant_blocks(m, state: NativeState) -> list[dict]:
    """Native content for an assistant turn: replayed verbatim when we have
    the original blocks (tool-call turns), else rebuilt from the chat shape."""
    tool_calls = _msg_get(m, "tool_calls") or []
    if tool_calls:
        first_id = _msg_get(tool_calls[0], "id")
        stored = state.by_first_call_id.get(first_id)
        if stored:
            if state.strip_thinking:
                return [b for b in stored if b.get("type") not in ("thinking", "redacted_thinking")]
            return list(stored)
    blocks = []
    content = _msg_get(m, "content")
    if isinstance(content, str):
        b = _text_block(content)
        if b:
            blocks.append(b)
    elif isinstance(content, list):
        blocks.extend(_content_to_blocks(content))
    for tc in tool_calls:
        fn = _msg_get(tc, "function")
        try:
            args = json.loads(_msg_get(fn, "arguments") or "{}")
        except (json.JSONDecodeError, TypeError):
            args = {"_raw": _msg_get(fn, "arguments")}
        blocks.append(
            {
                "type": "tool_use",
                "id": _msg_get(tc, "id"),
                "name": _msg_get(fn, "name"),
                "input": args if isinstance(args, dict) else {"_value": args},
            }
        )
    return blocks


def convert_messages(messages, state: NativeState) -> tuple[str, list[dict]]:
    """Chat-completions history -> (system_text, Anthropic messages).

    - system messages are hoisted and joined (Anthropic takes one system);
    - tool results become tool_result blocks in a user message, placed
      BEFORE any text in that message;
    - consecutive same-role messages are merged (the judge's loop appends a
      pressure note right after tool results, and evict stubs are user
      text), keeping strict user/assistant alternation;
    - an empty user message is rendered as a placeholder line, because the
      API rejects empty text blocks.
    """
    system_parts: list[str] = []
    out: list[dict] = []

    def _push(role: str, blocks: list[dict]) -> None:
        if not blocks:
            return
        if out and out[-1]["role"] == role:
            existing = out[-1]["content"]
            if role == "user":
                # tool_result blocks must lead the user message
                results = [b for b in existing + blocks if b["type"] == "tool_result"]
                others = [b for b in existing + blocks if b["type"] != "tool_result"]
                out[-1]["content"] = results + others
            else:
                out[-1]["content"] = existing + blocks
            return
        out.append({"role": role, "content": list(blocks)})

    for m in messages:
        role = _msg_get(m, "role")
        if role == "system":
            c = _msg_get(m, "content")
            text = c if isinstance(c, str) else "\n".join(
                b["text"] for b in _content_to_blocks(c) if b.get("type") == "text"
            )
            if text:
                system_parts.append(text)
        elif role == "user":
            _push(
                "user",
                _content_to_blocks(_msg_get(m, "content"))
                or [{"type": "text", "text": "(continue)"}],
            )
        elif role == "assistant":
            _push("assistant", _assistant_blocks(m, state))
        elif role == "tool":
            content = str(_msg_get(m, "content") or "")
            block = {
                "type": "tool_result",
                "tool_use_id": _msg_get(m, "tool_call_id"),
                "content": content if content else "(empty result)",
            }
            if content.startswith("Error:"):
                block["is_error"] = True
            _push("user", [block])
        else:
            logger.warning(f"  [anthropic-native] dropping message with role {role!r}")

    if out and out[0]["role"] != "user":
        out.insert(0, {"role": "user", "content": [{"type": "text", "text": "(begin)"}]})
    # Claude 5 rejects assistant prefill: the request must END on a user turn.
    if out and out[-1]["role"] != "user":
        out.append({"role": "user", "content": [{"type": "text", "text": "(continue)"}]})
    for m in out:
        if m["role"] == "user" and not m["content"]:
            m["content"] = [{"type": "text", "text": "(continue)"}]
    return "\n".join(system_parts), out


class NativeShim:
    """Quacks like a ChatCompletion for the judge's loop and recorder."""

    class _Choice:
        def __init__(self, message):
            self.message = message

    def __init__(self, raw, message, usage, cache_write_tokens: int):
        self._raw = raw
        self.choices = [self._Choice(message)]
        self.usage = usage
        self.cache_write_tokens = cache_write_tokens
        self.model_extra = {}
        self.stop_reason = getattr(raw, "stop_reason", None)

    def model_dump(self, **kwargs):
        return {
            "adapter": "anthropic_native",
            "usage": self.usage.model_dump(**kwargs) if self.usage else None,
            "cache_write_tokens": self.cache_write_tokens,
            "stop_reason": self.stop_reason,
            "chat_message": self.choices[0].message.model_dump(**kwargs),
            "raw_response": self._raw.model_dump(**kwargs) if hasattr(self._raw, "model_dump") else str(self._raw),
        }


class _UsageWithCache(CompletionUsage):
    """CompletionUsage plus the Anthropic cache-write count (read by
    llm_utils.usage_breakdown via getattr)."""

    cache_write_tokens: int = 0


def build_request(*, model, messages, chat_tools, reasoning_effort, state: NativeState,
                  max_tokens: int = DEFAULT_MAX_TOKENS) -> dict:
    """The exact kwargs for client.messages.create (also what tests inspect)."""
    system_text, native_msgs = convert_messages(messages, state)
    kwargs = {
        "model": model,
        "max_tokens": max_tokens,
        "messages": native_msgs,
        "tools": convert_tools(chat_tools),
        # Adaptive thinking is the Claude 5 default; stating it makes the
        # Opus/Sonnet 4.x graders think too instead of running thinking-off.
        "thinking": {"type": "adaptive"},
        # Automatic breakpoint that walks forward with the conversation.
        "cache_control": {"type": "ephemeral"},
    }
    if system_text:
        # Explicit breakpoint on the frozen prefix: caches tools + system.
        kwargs["system"] = [
            {"type": "text", "text": system_text, "cache_control": {"type": "ephemeral"}}
        ]
    effort = map_effort(reasoning_effort)
    if effort:
        kwargs["output_config"] = {"effort": effort}
    return kwargs


def _shim_from_response(response, state: NativeState) -> NativeShim:
    tool_calls = []
    text_parts = []
    native_blocks = []
    for block in response.content or []:
        btype = getattr(block, "type", None)
        native_blocks.append(block.model_dump(exclude_none=True) if hasattr(block, "model_dump") else dict(block))
        if btype == "tool_use":
            tool_calls.append(
                ChatCompletionMessageFunctionToolCall(
                    id=block.id,
                    type="function",
                    function=Function(name=block.name, arguments=json.dumps(block.input)),
                )
            )
        elif btype == "text" and getattr(block, "text", None):
            text_parts.append(block.text)
    if tool_calls:
        state.by_first_call_id[tool_calls[0].id] = native_blocks

    message = ChatCompletionMessage(
        role="assistant",
        content="\n".join(text_parts) if text_parts else None,
        tool_calls=tool_calls or None,
    )
    u = getattr(response, "usage", None)
    usage = None
    cache_write = 0
    if u is not None:
        uncached = getattr(u, "input_tokens", 0) or 0
        cache_read = getattr(u, "cache_read_input_tokens", 0) or 0
        cache_write = getattr(u, "cache_creation_input_tokens", 0) or 0
        prompt_total = uncached + cache_read + cache_write
        completion = getattr(u, "output_tokens", 0) or 0
        usage = _UsageWithCache(
            prompt_tokens=prompt_total,
            completion_tokens=completion,
            total_tokens=prompt_total + completion,
            prompt_tokens_details=PromptTokensDetails(cached_tokens=int(cache_read)),
            cache_write_tokens=int(cache_write),
        )
    return NativeShim(response, message, usage, int(cache_write))


def create(client, *, model, messages, chat_tools, reasoning_effort, state: NativeState,
           max_tokens: int = DEFAULT_MAX_TOKENS):
    """One native Messages call, presented as a chat completion.

    Raises whatever the SDK raises so the loop's retry/overflow handling
    applies unchanged (context overflows are 400s mentioning the context /
    prompt length). One recovery is built in: a 400 saying a thinking block
    is "bound to a different conversation" (preserved-thinking enforcement,
    which eviction can trigger on models that enforce it) strips thinking
    blocks from the replayed history and retries once.
    """
    kwargs = build_request(
        model=model, messages=messages, chat_tools=chat_tools,
        reasoning_effort=reasoning_effort, state=state, max_tokens=max_tokens,
    )
    try:
        response = _send(client, kwargs)
    except Exception as e:  # noqa: BLE001 — targeted recovery, then re-raise
        text = str(e)
        if ("bound to a different conversation" in text or "Invalid `signature`" in text) \
                and not state.strip_thinking:
            logger.warning(
                "  [anthropic-native] preserved-thinking mismatch after a history "
                "edit; stripping replayed thinking blocks and retrying once"
            )
            state.strip_thinking = True
            kwargs = build_request(
                model=model, messages=messages, chat_tools=chat_tools,
                reasoning_effort=reasoning_effort, state=state, max_tokens=max_tokens,
            )
            response = _send(client, kwargs)
        else:
            raise
    stop = getattr(response, "stop_reason", None)
    if stop == "refusal":
        details = getattr(response, "stop_details", None)
        raise RuntimeError(
            f"Anthropic refusal (stop_reason=refusal, category="
            f"{getattr(details, 'category', None)!r})"
        )
    if stop == "max_tokens":
        logger.warning(
            f"  [anthropic-native] reply truncated at max_tokens={kwargs['max_tokens']} "
            f"(thinking + tool calls exceeded the output ceiling); tool calls "
            f"after the cut are lost for this round"
        )
    return _shim_from_response(response, state)


# Content-silence watchdog. The HTTP read timeout only catches a silent
# wire; the API keeps a wedged stream alive with `ping` frames (which the SDK
# swallows before they reach the event iterator), so a turn that has stopped
# producing content can sit for hours without ever tripping it — twelve
# Opus@max gradings did exactly that on 2026-09-03 (68-92 min each, killed by
# hand). This watchdog measures silence in *events*: if no stream event
# (thinking/text/tool-input delta, block start/stop, message delta) arrives
# for CONTENT_SILENCE_SECONDS, the stream is closed from a helper thread,
# the blocked read raises, and `_send` re-raises as ContentSilenceTimeout so
# the caller's normal API-retry path resends the turn.
CONTENT_SILENCE_SECONDS = float(os.environ.get("ANTHROPIC_CONTENT_SILENCE_SECONDS", "600"))
_WATCHDOG_TICK_SECONDS = 5.0


class ContentSilenceTimeout(TimeoutError):
    """No stream event for CONTENT_SILENCE_SECONDS; the turn was aborted."""


def _send(client, kwargs, silence_seconds=None):
    """Stream the request and return the final Message.

    Streaming is what the SDK requires for large max_tokens without an HTTP
    timeout; the loop only needs the final message, so
    `stream.get_final_message()` gives it the same object a non-streaming
    call would. A client without `.stream` (test fakes) falls back to
    `.create`. Iteration is watched for content silence (see above).
    """
    import threading
    import time as _time

    stream_fn = getattr(client.messages, "stream", None)
    if stream_fn is None:
        return client.messages.create(**kwargs)
    limit = CONTENT_SILENCE_SECONDS if silence_seconds is None else float(silence_seconds)
    last_event = [_time.monotonic()]
    tripped = threading.Event()
    done = threading.Event()

    with stream_fn(**kwargs) as stream:
        def _watch():
            tick = min(_WATCHDOG_TICK_SECONDS, max(limit / 4, 0.05))
            while not done.wait(tick):
                if _time.monotonic() - last_event[0] > limit:
                    tripped.set()
                    try:
                        stream.close()
                    except Exception:  # noqa: BLE001 — closing is best-effort
                        pass
                    return

        watcher = threading.Thread(target=_watch, name="anthropic-silence-watchdog", daemon=True)
        watcher.start()
        try:
            for _event in stream:
                last_event[0] = _time.monotonic()
            message = stream.get_final_message()
        except Exception as e:  # noqa: BLE001 — translate the watchdog's close
            if tripped.is_set():
                raise ContentSilenceTimeout(
                    f"no stream event for {limit:.0f}s (pings excluded); "
                    f"stream closed by the content-silence watchdog"
                ) from e
            raise
        finally:
            done.set()
        if tripped.is_set():
            raise ContentSilenceTimeout(
                f"no stream event for {limit:.0f}s (pings excluded); "
                f"stream closed by the content-silence watchdog"
            )
        return message
