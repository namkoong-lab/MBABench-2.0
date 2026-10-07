"""LLM communication utilities for the judge system."""

import random
import threading
import time

import httpx
from openai import OpenAI

from .judge_identity import JudgeIdentity
from .logger import logger
from .repo_config import resolve_api_key

# Retry constants
BASE_DELAY = 1
MAX_DELAY = 60
RATE_LIMIT_DELAY = 30
MAX_ATTEMPTS = 10

# ---- Provider routing ------------------------------------------------------
# Which endpoint a grader hits is pinned by its JudgeIdentity (resolved from
# judge_identities.yaml) — never sniffed from the model slug or the ambient
# env. The identity's provider names the base_url and the api-key entry
# (keys resolve env-first, then config/config.yaml keys.*; see
# repo_config.resolve_api_key). Clients are cached per-provider so the
# underlying httpx connection pool is reused across calls (the OpenAI SDK is
# thread-safe).

_clients: dict[str, OpenAI] = {}
_clients_lock = threading.Lock()


_OK_IMAGE_MEDIA_TYPES = {"image/jpeg", "image/png", "image/gif", "image/webp"}


def strip_unsupported_anthropic_images(messages):
    """Return a copy of `messages` with image parts Anthropic can't accept removed.

    Anthropic's API accepts only jpeg/png/gif/webp image inputs and 400s on
    anything else. Excel workbooks embed emf/wmf/bmp/tiff pictures (which
    Gemini tolerates), so those parts must be dropped before an Anthropic call.
    Handles both OpenAI-compat `image_url` data-URI parts and native Anthropic
    `image` / `source.base64` parts; any image whose media type is missing,
    unparseable, or unsupported is dropped. Non-image content is untouched.
    Grading is driven by cell values/formulas, not the embedded pictures.
    """
    cleaned = []
    for m in messages:
        content = m.get("content") if isinstance(m, dict) else None
        if isinstance(content, list):
            new_content = []
            for item in content:
                if isinstance(item, dict):
                    itype = item.get("type")
                    media_type = None
                    if itype == "image_url":
                        url = item.get("image_url", {}).get("url", "")
                        media_type = (
                            url[5 : url.find(";")]
                            if url.startswith("data:") and ";" in url
                            else ""
                        )
                    elif itype == "image":
                        media_type = (item.get("source", {}) or {}).get(
                            "media_type", ""
                        )
                    if media_type is not None and media_type not in _OK_IMAGE_MEDIA_TYPES:
                        continue  # drop unsupported / unparseable image part
                new_content.append(item)
            m = {**m, "content": new_content}
        cleaned.append(m)
    return cleaned


def get_client(identity: JudgeIdentity) -> OpenAI:
    """Return a cached OpenAI-compatible client for the identity's provider.

    First call per provider builds the client; subsequent calls are dict
    lookups. Safe to call from worker threads — the lock only contends
    on cold start.
    """
    provider = identity.provider
    client = _clients.get(provider)
    if client is not None:
        return client
    with _clients_lock:
        client = _clients.get(provider)
        if client is not None:
            return client
        kwargs = {"api_key": resolve_api_key(identity.api_key_provider)}
        if identity.base_url is not None:
            kwargs["base_url"] = identity.base_url
        # 2026-09-15: the SDK's default httpx timeout allows only 5 s to
        # CONNECT. Under evening load the Forge gateway took 5-10 s to accept
        # a TCP connection and ~20 s to finish TLS, so ~30% of toy gradings
        # died on round 1 with "Request timed out." before any request was
        # sent. Keep the 600 s read budget; give connect a full minute.
        kwargs["timeout"] = httpx.Timeout(600.0, connect=60.0)
        client = OpenAI(**kwargs)
        _clients[provider] = client
        return client


_native_clients: dict = {}

# Native Anthropic HTTP timeouts, in seconds. Requests stream (see
# anthropic_native._send), so `read` is the longest silence tolerated
# BETWEEN bytes, not the length of a turn: a Sonnet turn thinking for 11
# minutes keeps sending thinking deltas the whole time. The old single
# `timeout=1800.0` made every phase 30 minutes; on 2026-09-03 three parallel
# gradings sat 44 minutes each on one dead read (last bytes ~14 min in, then
# 30 minutes of silence before "The read operation timed out") and lost the
# turn to a retry that then finished in seconds. Ten minutes of silence is
# already far past anything a live stream produces.
NATIVE_CONNECT_TIMEOUT_S = 30.0
NATIVE_READ_TIMEOUT_S = 600.0
NATIVE_WRITE_TIMEOUT_S = 120.0


def native_anthropic_timeout():
    """httpx.Timeout for the native client (connect / read / write / pool)."""
    import httpx

    return httpx.Timeout(
        connect=NATIVE_CONNECT_TIMEOUT_S,
        read=NATIVE_READ_TIMEOUT_S,
        write=NATIVE_WRITE_TIMEOUT_S,
        pool=NATIVE_CONNECT_TIMEOUT_S,
    )


def get_native_anthropic_client(identity: JudgeIdentity):
    """Cached native `anthropic.Anthropic` client for provider=anthropic.

    Why native (2026-09-02): Anthropic's OpenAI-compat endpoint documents no
    prompt caching, `reasoning_effort` ignored, and empty
    prompt_tokens_details — so neither the cache savings nor real effort
    arms are reachable through `get_client`. The single-pass loop routes
    provider=anthropic here (utils/anthropic_native.py); the
    OpenAI-compatible client above stays for every other provider.

    The explicit timeout matters: the SDK refuses non-streaming requests
    whose max_tokens implies more than ~10 minutes unless a timeout is set.
    See NATIVE_*_TIMEOUT_S for why the read phase is short.
    """
    if identity.provider != "anthropic":
        raise ValueError(f"native Anthropic client requested for provider {identity.provider!r}")
    client = _native_clients.get("anthropic")
    if client is not None:
        return client
    with _clients_lock:
        client = _native_clients.get("anthropic")
        if client is None:
            import anthropic

            client = anthropic.Anthropic(
                api_key=resolve_api_key(identity.api_key_provider),
                timeout=native_anthropic_timeout(),
                max_retries=2,
            )
            _native_clients["anthropic"] = client
        return client

# Pricing per 1M tokens (input, output) for common OpenRouter models
_MODEL_PRICING = {
    "openai/gpt-4o": (5.0, 15.0),
    "openai/gpt-4o-mini": (0.15, 0.6),
    "openai/gpt-4-turbo": (10.0, 30.0),
    "openai/gpt-4": (30.0, 60.0),
    "anthropic/claude-opus-4-8": (5.0, 25.0),
    "anthropic/claude-sonnet-4-5-20250929": (3.0, 15.0),
    "anthropic/claude-3.5-sonnet": (3.0, 15.0),
    "anthropic/claude-3-opus": (15.0, 75.0),
    "google/gemini-pro-1.5": (3.5, 10.5),
    "google/gemini-2.5-pro": (1.25, 10.0),
    "google/gemini-2.5-flash": (0.15, 0.60),
    "google/gemini-2.5-flash-lite": (0.075, 0.30),
    "google/gemini-3-flash-preview": (0.5, 3.0),
    "google/gemini-3.5-flash": (0.5, 3.0),
    "google/gemini-3.7-flash": (0.75, 3.75),  # OpenRouter list 2026-08-31

    "google/gemini-3.6-flash": (0.5, 3.0),
    "google/gemini-3.1-pro-preview": (2.0, 12.0),
    "openai/gpt-5.5": (5.0, 30.0),   # judge-robustness study 2026-07-26 (OpenRouter)
    # Judge bake-off 2026-08 — PROVISIONAL list prices; ground truth comes
    # from the maintainer's isolated-window credits-dashboard deltas after the runs
    # ($/MTok = delta / token_tracking totals), then recalibrate here.
    "openai/gpt-5.6-sol": (5.0, 30.0),       # OpenAI list (matches cli models_config)
    # Forge does not return per-call cost and publishes no price in
    # models.list; OpenAI's list rate is assumed until a Forge credit diff
    # verifies it (2026-09-15; verified the same day against Forge's usage
    # log: $5/$30 exactly, cached input at 10% — see _CACHE_PRICING).
    "tensorblock/gpt-5.6-sol": (5.0, 30.0),
    "anthropic/claude-opus-5": (5.0, 25.0),  # Anthropic list price
    # (gemini-3.7-flash now priced above at the 2026-08-31 OpenRouter list —
    # a stale duplicate entry here was silently shadowing it, since the
    # LAST duplicate key in a dict literal wins.)
    # Judge-effort canary graders (2026-09) — PROVISIONAL; the canary's
    # before/after balance readings are the calibration for all of these.
    "openai/gpt-5.6-terra": (5.0, 30.0),      # assume sol-tier until billed
    "anthropic/claude-sonnet-5": (2.0, 10.0),  # Anthropic list price
}
_DEFAULT_PRICING = (10.0, 30.0)  # Fallback pricing per 1M tokens


def calculate_message_size(messages):
    """Calculate the total string length of all messages.

    Returns two counts: one for text-only content and one that also includes
    image_url data (base64 data URIs).

    Args:
        messages: List of message dicts in OpenAI format

    Returns:
        dict: {"text": int, "total": int} where "text" counts only text content
              and "total" also includes image_url data.
    """
    text_size = 0
    image_size = 0
    for msg in messages:
        if isinstance(msg, dict):
            content = msg.get("content", "")
            if isinstance(content, str):
                text_size += len(content)
            elif isinstance(content, list):
                for item in content:
                    if isinstance(item, dict):
                        if "text" in item:
                            text_size += len(item["text"])
                        elif item.get("type") == "image_url":
                            url = item.get("image_url", {}).get("url", "")
                            image_size += len(url)
    return {"text": text_size, "total": text_size + image_size}


# Prompt-cache billing multipliers per provider, applied to the input price:
# (cache READ multiplier, cache WRITE multiplier). Measured/documented:
#   openai    — cached input billed at 25% (sol canary step 2: 92% cached,
#               balance delta agreed with 0.25 within 1%); no write premium.
#   anthropic — cache reads 10%, 5-minute-TTL cache writes 125% (Anthropic
#               prompt-caching pricing; reads are 2.5% on claude-fable-5-1,
#               which is not a registered grader).
# Any other provider: cached tokens priced at the full input rate (no
# discount assumed until measured). OpenRouter/Gemini implicit caching shows
# up in the provider's exact per-call cost, not in these estimates.
_CACHE_PRICING = {
    "openai": (0.25, 1.0),
    "anthropic": (0.10, 1.25),
    # tensorblock (Forge) — measured 2026-09-15 from Forge's own usage log
    # over the 430-grading toy run 2 (1,769 Sol calls, $446.67): every
    # charge = $5/M uncached input + $0.50/M cached input + $30/M output,
    # i.e. cached input at 10%. The run's recorded estimates (no discount)
    # came to $1,144 against $447 actually billed.
    "tensorblock": (0.10, 1.0),
}


def calculate_cost(model, prompt_tokens, completion_tokens,
                   cached_tokens=0, cache_write_tokens=0, provider=None):
    """Estimate cost for an LLM call based on model and token counts.

    Args:
        model: Model identifier string (pricing key)
        prompt_tokens: TOTAL input tokens (uncached + cached + cache-write)
        completion_tokens: Number of output tokens
        cached_tokens: input tokens served from the prompt cache (subset of
            prompt_tokens), billed at the provider's read multiplier
        cache_write_tokens: input tokens written to the cache this call
            (subset of prompt_tokens), billed at the write multiplier
        provider: identity.provider, selects the cache multipliers

    Returns:
        dict: {"total_cost", "prompt_cost", "completion_cost",
               "cached_tokens", "cache_write_tokens", "cache_savings"}
    """
    input_price, output_price = _MODEL_PRICING.get(model, _DEFAULT_PRICING)
    cached_tokens = max(0, int(cached_tokens or 0))
    cache_write_tokens = max(0, int(cache_write_tokens or 0))
    read_mult, write_mult = _CACHE_PRICING.get(provider or "", (1.0, 1.0))
    uncached = max(0, prompt_tokens - cached_tokens - cache_write_tokens)
    per_tok = input_price / 1_000_000
    prompt_cost = (
        uncached * per_tok
        + cached_tokens * per_tok * read_mult
        + cache_write_tokens * per_tok * write_mult
    )
    completion_cost = (completion_tokens / 1_000_000) * output_price
    return {
        "total_cost": prompt_cost + completion_cost,
        "prompt_cost": prompt_cost,
        "completion_cost": completion_cost,
        "cached_tokens": cached_tokens,
        "cache_write_tokens": cache_write_tokens,
        # what the same call would have cost with no cache accounting
        "cache_savings": (prompt_tokens * per_tok) - prompt_cost,
    }


def usage_breakdown(usage) -> dict:
    """Normalise a response.usage object (chat-completions, the responses
    shim, or the native-Anthropic shim) into total / cached / cache-write
    input tokens plus output tokens. Missing fields read as 0."""
    if usage is None:
        return {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0,
                "cached_tokens": 0, "cache_write_tokens": 0}
    details = getattr(usage, "prompt_tokens_details", None)
    cached = getattr(details, "cached_tokens", None) if details is not None else None
    if cached is None and isinstance(details, dict):
        cached = details.get("cached_tokens")
    return {
        "prompt_tokens": getattr(usage, "prompt_tokens", 0) or 0,
        "completion_tokens": getattr(usage, "completion_tokens", 0) or 0,
        "total_tokens": getattr(usage, "total_tokens", 0) or 0,
        "cached_tokens": int(cached or 0),
        "cache_write_tokens": int(getattr(usage, "cache_write_tokens", 0) or 0),
    }


def backoff(attempt):
    """Calculate exponential backoff with full jitter for retry delays."""
    return min(MAX_DELAY, (BASE_DELAY * (2**attempt)) * random.uniform(0.5, 1.5))


def robust_send_message(
    client,
    messages,
    identity,
    system_instruction=None,
    response_format=None,
    reasoning_effort=None,
):
    """Send a message to the identity's provider with exponential backoff.

    Args:
        client: OpenAI-compatible client (get_client(identity))
        messages: List of message dicts
        identity: JudgeIdentity — supplies the wire model id and the
            provider-specific request shaping below
        system_instruction: Optional system message to prepend
        response_format: Optional response format dict (e.g., {"type": "json_object"})
        reasoning_effort: Optional reasoning effort level passed verbatim to the
            API (`"low"`, `"medium"`, `"high"`, `"minimal"`, or `"none"`).
            Models without thinking support may reject the kwarg.

    Returns:
        tuple: (response, metrics_dict) where metrics_dict contains:
            - message_size: Character count of text-only content
            - message_size_with_images: Character count including image_url data
            - prompt_tokens: Actual tokens used for prompt
            - completion_tokens: Actual tokens used for completion
            - total_tokens: Total tokens used
    """
    attempt = 0
    while True:
        try:
            api_messages = list(messages)

            if system_instruction:
                if not messages or (
                    isinstance(messages[0], dict)
                    and messages[0].get("role") != "system"
                ):
                    system_msg = {"role": "system", "content": system_instruction}
                    messages.insert(0, system_msg)
                    api_messages.insert(0, system_msg)

            size_info = calculate_message_size(api_messages)

            kwargs = {"model": identity.model, "messages": api_messages}
            if response_format:
                kwargs["response_format"] = response_format
            if reasoning_effort is not None:
                kwargs["reasoning_effort"] = reasoning_effort
            if identity.provider in ("openai", "tensorblock"):
                # OpenAI (direct, or via the Forge gateway) 400s on non-image MIME parts (emf/wmf/bmp/
                # tiff embedded in workbooks) exactly like Anthropic; the
                # OpenRouter route normalized these, the direct API doesn't.
                kwargs["messages"] = strip_unsupported_anthropic_images(
                    kwargs["messages"]
                )
            if identity.provider == "anthropic":
                # Anthropic's OpenAI-compat endpoint: the native API requires
                # max_tokens (compat default is small), and the judge's
                # reasoning_effort values ("minimal"/"none") aren't valid
                # Anthropic efforts — drop it and run the model's default,
                # which matches the Gemini path's minimal-thinking intent.
                # response_format {"type": "json_object"} is REJECTED (400:
                # only 'json_schema' is accepted) — drop it and rely on the
                # judge prompt's strict-JSON instruction plus
                # _extract_json_from_response, which handles fenced and
                # trailing-text responses.
                # 32000 (was 16000): on Claude Opus 5 thinking is always on and
                # counts against max_tokens, so 16k truncated the largest
                # category's JSON mid-array (14 of 23 parse failures were
                # Formatting, the longest form). 2026-08-12.
                kwargs["max_tokens"] = 32000
                kwargs.pop("reasoning_effort", None)
                kwargs.pop("response_format", None)
                kwargs["messages"] = strip_unsupported_anthropic_images(
                    kwargs["messages"]
                )

            response = client.chat.completions.create(**kwargs)

            metrics = {
                "message_size": size_info["text"],
                "message_size_with_images": size_info["total"],
                "prompt_tokens": 0,
                "completion_tokens": 0,
                "total_tokens": 0,
            }

            if hasattr(response, "usage") and response.usage:
                metrics["prompt_tokens"] = response.usage.prompt_tokens
                metrics["completion_tokens"] = response.usage.completion_tokens
                metrics["total_tokens"] = response.usage.total_tokens

            return response, metrics
        except Exception as e:
            error_msg = str(e)

            if "400" in error_msg or "invalid" in error_msg.lower():
                raise

            if attempt >= MAX_ATTEMPTS - 1:
                raise

            delay = backoff(attempt)

            if (
                "429" in error_msg
                or "rate limit" in error_msg.lower()
                or "quota" in error_msg.lower()
            ):
                delay = RATE_LIMIT_DELAY

            import traceback

            traceback.print_exc()
            logger.info(
                f"   Retry {attempt + 1}/{MAX_ATTEMPTS} after {delay:.2f}s due to: {type(e).__name__}"
            )
            time.sleep(delay)
            attempt += 1
