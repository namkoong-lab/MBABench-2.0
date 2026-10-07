"""OpenAI Responses-API adapter for the judge's tool loop.

Why this exists (probed live 2026-09-01): on /v1/chat/completions, gpt-5.6-sol
and gpt-5.6-terra reject function tools combined with any real reasoning tier —

    "Function tools with reasoning_effort are not supported for gpt-5.6-sol in
     /v1/chat/completions. To use function tools, use /v1/responses or set
     reasoning_effort to 'none'."

— which is exactly why the judge has always forced 'none' for OpenAI graders.
The judge-effort canary needs sol/terra at real reasoning tiers WITH tools, so
this module routes those calls through /v1/responses and translates both
directions, letting the loop keep speaking chat-completions shapes untouched:

    loop (chat messages/tools) -> convert -> responses.create -> shim that
    looks like a ChatCompletion (choices[0].message with real SDK
    ChatCompletionMessage / tool-call objects, usage with prompt_tokens etc.)

Reasoning statefulness: we run store=False (nothing retained server-side) and
request encrypted reasoning content back. The Responses API requires that a
resent function_call be accompanied by its reasoning item, so the adapter
keeps a side-table mapping each assistant turn (keyed by its first tool-call
id) to that turn's reasoning items, and re-emits them just before the calls on
every subsequent request. Eviction needs no special handling: when the loop
drops an assistant message from wire context, its function_calls are no longer
converted, so its reasoning items simply stop being emitted with it.

Scope: wired into single_pass_judge_case only (the canary's mode).
Other providers keep the chat-completions path byte-for-byte.
"""

import json

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


def wants_responses_api(provider: str, reasoning_effort) -> bool:
    """True when this call must route via /v1/responses.

    Only the OpenAI wire (direct, or via the TensorBlock Forge gateway —
    probed 2026-09-15: Forge forwards /v1/responses with store=False and
    returns reasoning items with encrypted_content, and a two-round tool
    exchange resending them succeeds), and only for a real reasoning tier —
    effort None or 'none' stays on chat/completions, the path every existing
    OpenAI grading used.
    """
    return provider in ("openai", "tensorblock") and reasoning_effort not in (None, "none")


class ReasoningState:
    """Side-table: first tool-call id of an assistant turn -> that turn's
    reasoning items (plain dicts with encrypted_content), for re-emission."""

    def __init__(self):
        self.by_first_call_id: dict[str, list[dict]] = {}


def _msg_get(m, key, default=None):
    """Field access across plain dicts and SDK message objects."""
    if isinstance(m, dict):
        return m.get(key, default)
    return getattr(m, key, default)


def convert_tools(chat_tools: list[dict]) -> list[dict]:
    """Chat-completions function tools -> Responses-API tool entries."""
    out = []
    for t in chat_tools:
        fn = t["function"]
        out.append(
            {
                "type": "function",
                "name": fn["name"],
                "description": fn.get("description", ""),
                "parameters": fn.get("parameters", {}),
            }
        )
    return out


def _content_to_input_parts(content):
    """Chat message content (str or parts list) -> Responses input parts."""
    if isinstance(content, str):
        return [{"type": "input_text", "text": content}]
    parts = []
    for part in content or []:
        if not isinstance(part, dict):
            continue
        if part.get("type") == "text" or "text" in part:
            parts.append({"type": "input_text", "text": part.get("text", "")})
        elif part.get("type") == "image_url":
            url = (part.get("image_url") or {}).get("url", "")
            if url:
                parts.append({"type": "input_image", "image_url": url})
    return parts or [{"type": "input_text", "text": ""}]


def convert_messages(messages, state: ReasoningState) -> list:
    """Chat-completions history -> Responses `input` item list.

    Handles the four shapes the judge's loop produces: dict system/user
    messages (seed, pressure notes, nudges), dict tool results, assistant
    messages from this adapter's own shims (SDK objects with tool_calls),
    and plain-text assistant messages.
    """
    items = []
    for m in messages:
        role = _msg_get(m, "role")
        if role in ("system", "user"):
            items.append(
                {"role": role, "content": _content_to_input_parts(_msg_get(m, "content"))}
            )
        elif role == "assistant":
            content = _msg_get(m, "content")
            if content:
                text = content if isinstance(content, str) else json.dumps(content)
                items.append({"role": "assistant", "content": text})
            tool_calls = _msg_get(m, "tool_calls") or []
            if tool_calls:
                first_id = _msg_get(tool_calls[0], "id")
                for r_item in state.by_first_call_id.get(first_id, []):
                    items.append(r_item)
            for tc in tool_calls:
                fn = _msg_get(tc, "function")
                items.append(
                    {
                        "type": "function_call",
                        "call_id": _msg_get(tc, "id"),
                        "name": _msg_get(fn, "name"),
                        "arguments": _msg_get(fn, "arguments") or "{}",
                    }
                )
        elif role == "tool":
            items.append(
                {
                    "type": "function_call_output",
                    "call_id": _msg_get(m, "tool_call_id"),
                    "output": str(_msg_get(m, "content") or ""),
                }
            )
        else:
            logger.warning(
                f"  [responses-adapter] dropping message with role {role!r}"
            )
    return items


class ResponsesShim:
    """Quacks like a ChatCompletion for the judge's loop and recorder.

    choices[0].message is a REAL ChatCompletionMessage (with real tool-call
    objects) so _msg_to_dict / _measure_message_chars / append_wire treat it
    exactly like a chat-completions reply. usage maps input->prompt tokens.
    model_dump() carries the raw Responses payload for the trajectory.
    """

    class _Choice:
        def __init__(self, message):
            self.message = message

    def __init__(self, raw, message, usage):
        self._raw = raw
        self.choices = [self._Choice(message)]
        self.usage = usage
        self.model_extra = {}

    def model_dump(self, **kwargs):
        return {
            "adapter": "openai_responses",
            "usage": self.usage.model_dump(**kwargs) if self.usage else None,
            "chat_message": self.choices[0].message.model_dump(**kwargs),
            "raw_response": self._raw.model_dump(**kwargs),
        }


def create(client, *, model, messages, chat_tools, reasoning_effort,
           state: ReasoningState):
    """One /v1/responses call, presented as a chat completion.

    Raises whatever the SDK raises — the loop's existing retry/overflow
    handling applies unchanged (Responses context overflows also surface as
    400s mentioning context)."""
    response = client.responses.create(
        model=model,
        input=convert_messages(messages, state),
        tools=convert_tools(chat_tools),
        reasoning={"effort": reasoning_effort},
        store=False,
        include=["reasoning.encrypted_content"],
    )

    tool_calls = []
    text_parts = []
    reasoning_items = []
    for item in response.output or []:
        if item.type == "function_call":
            tool_calls.append(
                ChatCompletionMessageFunctionToolCall(
                    id=item.call_id,
                    type="function",
                    function=Function(name=item.name, arguments=item.arguments),
                )
            )
        elif item.type == "reasoning":
            reasoning_items.append(item.model_dump(exclude_none=True))
        elif item.type == "message":
            for part in getattr(item, "content", None) or []:
                text = getattr(part, "text", None)
                if text:
                    text_parts.append(text)

    if tool_calls and reasoning_items:
        state.by_first_call_id[tool_calls[0].id] = reasoning_items

    message = ChatCompletionMessage(
        role="assistant",
        content="\n".join(text_parts) if text_parts else None,
        tool_calls=tool_calls or None,
    )

    u = response.usage
    usage = None
    if u is not None:
        # Cached input is billed at 25% on OpenAI; surface it in the same
        # prompt_tokens_details.cached_tokens slot chat/completions uses so
        # the loop's cost accounting (llm_utils.usage_breakdown) sees it.
        details = getattr(u, "input_tokens_details", None)
        cached = getattr(details, "cached_tokens", 0) if details is not None else 0
        usage = CompletionUsage(
            prompt_tokens=getattr(u, "input_tokens", 0) or 0,
            completion_tokens=getattr(u, "output_tokens", 0) or 0,
            total_tokens=getattr(u, "total_tokens", 0) or 0,
            prompt_tokens_details=PromptTokensDetails(cached_tokens=int(cached or 0)),
        )

    return ResponsesShim(response, message, usage)
