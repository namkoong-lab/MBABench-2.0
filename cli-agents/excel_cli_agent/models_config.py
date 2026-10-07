"""Centralized model configuration: pricing, defaults, and cost calculation."""

import json
import re
import urllib.request
from typing import Dict, Optional

DEFAULT_MODEL = "gpt-4o-mini"
DEFAULT_MAX_COMPLETION_TOKENS = 8000

# Gemini 3.8 Flash only (2026-09-22) flag. Allow custom tool calls setup and prompt
GEMINI_TOOL_CALL_MODELS = frozenset({"tensorblock/gemini-3.8-flash"})


def uses_gemini_tool_calls(model: Optional[str]) -> bool:
    """True only for the models in GEMINI_TOOL_CALL_MODELS (the exact model
    id the identity sends), so the Gemini path can never switch on by a
    substring of some other model's name."""
    return (model or "") in GEMINI_TOOL_CALL_MODELS


# Model pricing (per 1M tokens) - Updated January 2026
# FALLBACK ONLY: calculate_cost() prefers live prices fetched from
# OpenRouter's public model list at run start (see resolve_pricing); this
# table is used only when that fetch fails or the model isn't listed.
MODEL_PRICING = {
    # GPT-5 series (Flagship models - January 2025)
    "gpt-5.1": {"input": 1.250, "output": 10.00},
    "gpt-5-mini": {"input": 0.250, "output": 2.00},
    "gpt-5-nano": {"input": 0.050, "output": 0.400},
    "gpt-5-pro": {"input": 15.00, "output": 120.00},
    # GPT-4.1 series (Updated models with fine-tuning support)
    "gpt-4.1": {"input": 3.00, "output": 12.00},
    "gpt-4.1-mini": {"input": 0.80, "output": 3.20},
    "gpt-4.1-nano": {"input": 0.20, "output": 0.80},
    # GPT-4o family (Legacy - still available via API)
    "gpt-4o": {"input": 3.00, "output": 10.00},
    "gpt-4o-mini": {"input": 0.150, "output": 0.600},
    "gpt-4o-2024-11-20": {"input": 3.00, "output": 10.00},
    "gpt-4o-2024-08-06": {"input": 3.00, "output": 10.00},
    "gpt-4o-2024-05-13": {"input": 5.00, "output": 15.00},
    "gpt-4o-mini-2024-07-18": {"input": 0.150, "output": 0.600},
    # Realtime API models (Text pricing)
    "gpt-realtime": {"input": 4.00, "output": 16.00},
    "gpt-realtime-mini": {"input": 0.60, "output": 2.40},
    # GPT-4 Turbo family (Legacy)
    "gpt-4-turbo": {"input": 10.00, "output": 30.00},
    "gpt-4-turbo-2024-04-09": {"input": 10.00, "output": 30.00},
    "gpt-4-turbo-preview": {"input": 10.00, "output": 30.00},
    # GPT-4 family (Legacy)
    "gpt-4": {"input": 30.00, "output": 60.00},
    "gpt-4-0613": {"input": 30.00, "output": 60.00},
    "gpt-4-32k": {"input": 60.00, "output": 120.00},
    # GPT-3.5 Turbo family (Legacy)
    "gpt-3.5-turbo": {"input": 0.50, "output": 1.50},
    "gpt-3.5-turbo-0125": {"input": 0.50, "output": 1.50},
    "gpt-3.5-turbo-1106": {"input": 1.00, "output": 2.00},
    # O1/O4 family (Reasoning models)
    "o1": {"input": 15.00, "output": 60.00},
    "o1-mini": {"input": 3.00, "output": 12.00},
    "o1-preview": {"input": 15.00, "output": 60.00},
    "o4-mini": {"input": 4.00, "output": 16.00},
    # OpenRouter models (third-party via OpenRouter) - Updated Jan 2026
    "moonshotai/kimi-k2-thinking": {"input": 0.57, "output": 2.42},
    "deepseek/deepseek-r1": {"input": 0.55, "output": 2.19},
    "deepseek/deepseek-v3-0324": {"input": 0.50, "output": 1.50},
    "deepseek/deepseek-v3.2": {"input": 0.25, "output": 0.38},
    "openai/gpt-5.2": {"input": 2.50, "output": 10.00},
    "openai/gpt-5.2-codex": {"input": 1.75, "output": 14.00},
    "openai/gpt-5.3-codex": {"input": 1.75, "output": 14.00},
    "anthropic/claude-opus-4.5": {"input": 5.00, "output": 25.00},
    "anthropic/claude-opus-4.6": {"input": 5.00, "output": 25.00},
    "claude-opus-4-6": {"input": 5.00, "output": 25.00},  # Anthropic direct model ID
    "claude-opus-4-8": {"input": 5.00, "output": 25.00},  # Anthropic direct model ID
    "claude-fable-5": {
        "input": 25.00,
        "output": 50.00,
    },  # Anthropic direct model ID; see DIRECT_API_PRICING
    "gpt-5.6-sol": {"input": 5.00, "output": 30.00},  # OpenAI direct model ID
    # 2026-09-19 backstop for the 101-task rerun cohorts: the OpenRouter list
    # price on that day, used only when the live fetch fails (a failed fetch
    # used to record these runs at $0). NOT billing-verified - claude-fable-5
    # billed $25 input against a $10 listing; pin the verified rate in
    # DIRECT_API_PRICING after the first isolated run's credit diff.
    "claude-fable-5-1": {"input": 10.00, "output": 50.00},  # Anthropic direct model ID
    "gpt-6-astra": {"input": 10.00, "output": 50.00},  # OpenAI direct model ID
    # 2026-09-20 backstop for the Forge cohort tensorblock/grok-4.6 (looked up
    # by its bare id): xAI's list price that day, equal to OpenRouter's
    # x-ai/grok-4.6. NOT billing-verified - Forge publishes no prices; check
    # the first run against a Forge credit diff.
    "grok-4.6": {"input": 2.00, "output": 6.00},
    # tensorblock/Kimi-K3 and tensorblock/gemini-3.8-flash: see
    # DIRECT_API_PRICING (billed rates read from Forge's usage log).
    "google/gemini-3-pro-preview": {"input": 1.25, "output": 10.00},
    "z-ai/glm-4.7": {"input": 0.40, "output": 1.50},
    "x-ai/grok-4": {"input": 3.00, "output": 15.00},
    "mistralai/mistral-medium-3": {"input": 0.40, "output": 2.00},
    "mistralai/mistral-large-2512": {"input": 0.50, "output": 1.50},
    "qwen/qwen3-235b-a22b": {"input": 0.18, "output": 0.54},
    "qwen/qwen3-235b-a22b-2507": {"input": 0.071, "output": 0.463},
    "qwen/qwen3-235b-a22b-thinking-2507": {"input": 0.11, "output": 0.60},
    "allenai/olmo-3.1-32b-think:free": {"input": 0.0, "output": 0.0},
    "allenai/olmo-3.1-32b-think": {"input": 0.15, "output": 0.50},
    "moonshotai/kimi-k2.5": {"input": 0.60, "output": 2.50},
    "qwen/qwen3.5-397b-a17b": {"input": 0.20, "output": 0.60},
}

# Per-model defaults (only models that differ from standard)
MODEL_DEFAULTS: Dict[str, Dict] = {
    "openai/gpt-5.2": {
        "reasoning_effort": "none",
        "max_completion_tokens": 64000,
    },
    "anthropic/claude-opus-4.5": {
        "thinking_budget_tokens": 50000,
        "max_completion_tokens": 64000,
    },
    "claude-opus-4-6": {
        "thinking_budget_tokens": 50000,
        "max_completion_tokens": 64000,
    },
    # claude-fable-5: thinking always on; effort (not budget_tokens) controls it.
    # reasoning_effort maps to output_config.effort in the Anthropic direct path.
    "claude-fable-5": {
        "reasoning_effort": "max",
        "max_completion_tokens": 64000,
    },
    # gpt-5.6-sol: xhigh is the highest reasoning_effort the API accepts
    # (the ChatGPT UI "Ultra" label does not exist as an API value).
    "gpt-5.6-sol": {
        "reasoning_effort": "xhigh",
        "max_completion_tokens": 64000,
    },
    "google/gemini-3-pro-preview": {
        "reasoning_effort": "high",
        "max_completion_tokens": 64000,
    },
    "x-ai/grok-4": {
        "max_completion_tokens": 64000,
    },
    "z-ai/glm-4.7": {
        "max_completion_tokens": 64000,
    },
    "qwen/qwen3-235b-a22b-thinking-2507": {
        "max_completion_tokens": 64000,
    },
    "moonshotai/kimi-k2.5": {
        "max_completion_tokens": 64000,
    },
    "qwen/qwen3.5-397b-a17b": {
        "max_completion_tokens": 64000,
    },
}

# Per-call model timeout (seconds) by reasoning effort. 2026-09-19 (maintainer): the
# two top tiers get the same 60 minutes - "xhigh" was 300 s, a hard bound on
# the whole call, while "max" had 3600, so an OpenAI top-tier cohort could lose
# a long thinking turn to the clock that an Anthropic one never faced.
TIMEOUT_BY_REASONING: Dict[Optional[str], int] = {
    "max": 3600,
    "xhigh": 3600,
    "high": 240,
    None: 180,
}

# TensorBlock Forge only: longest silence (no bytes at all, response headers
# included) before a call is cut and retried. 2026-09-21: Forge left Grok
# requests unanswered - task 41 sat 60 min on one call, then 2.6 h on another
# (retry included), while a healthy call streams its thinking from the first
# seconds and finished in 4-5 min. The per-call budget stays
# api_timeout_seconds; this only splits it into shorter tries.
FORGE_STALL_TIMEOUT_SECONDS = 600

# 2026-09-22 (maintainer): gpt-6-astra through Forge sends NO byte - not even the
# response headers - until its thinking is done (probed 2026-09-22; Forge
# serves it from Azure OpenAI), so a long think looks exactly like a hang. Its
# longest genuine think on the direct OpenAI cohort (pv 1609) was ~14 min
# (57k tokens); the maintainer set 15 min. Every other Forge model streams its thinking
# from the first seconds and keeps the 10-minute limit.
#
# 2026-09-23: claude-opus-5 through Forge is silent the same way - probed at
# max effort, the first byte came after 114 s, when the 11.6k-token think was
# done. Its sibling claude-opus-5-5 ran single CLI steps of 67k-128k output
# tokens (up to ~21 min) through Forge that day, and a think can run to the
# 128000-token cap (~22 min at ~100 tokens/s). 30 min covers that with margin;
# a gateway that never answers waits 30 min, not 10.
FORGE_SILENT_THINKING_STALL_SECONDS = {"tensorblock/gpt-6-astra": 900,
                                       "tensorblock/claude-opus-5": 1800}


def resolve_stall_timeout(
    base_url: Optional[str], model: Optional[str] = None
) -> Optional[int]:
    """Seconds of total silence a call may show before it is retried, or None
    where no such limit applies (every endpoint but Forge: OpenAI does not
    stream thinking, so a long silence there is a healthy call)."""
    if base_url and "tensorblock" in base_url.lower():
        return FORGE_SILENT_THINKING_STALL_SECONDS.get(
            model or "", FORGE_STALL_TIMEOUT_SECONDS
        )
    return None


# 2026-09-24: Forge models whose prompt cache needs a routing key. This harness
# rebuilds every step's prompt (the fixed system prompt, then a user message
# that changes), so a step shares only its first ~10k tokens with the step
# before. Probed on GLM 5.3 (served by Fireworks) with a real step 1 and step 2
# of task 8, 20-30 s apart: without a key the second call read 0-13 cached
# tokens (4 of 4 trials), with a prompt_cache_key 10,418-10,428 of ~23,600
# (4 of 4) - Fireworks sends same-key calls to the replica that holds the
# prefix. Billing only: the model sees the same text. Qwen 3.8 is not listed:
# it only reuses a whole earlier prompt, so a key changed nothing (0 cached
# either way).
FORGE_PROMPT_CACHE_KEY_MODELS = {"tensorblock/glm-5.3"}


# Iteration cap (one model call per iteration) when a run config names none.
# 40 is what every v2 API cohort actually ran with; the old fallback of 30
# would have silently shortened a cohort whose config omitted the key.
DEFAULT_MAX_ITERATIONS = 40


def resolve_api_timeout(
    reasoning_effort: Optional[str], explicit: Optional[int] = None
) -> int:
    """Seconds one model call may take: the run config's api_timeout_seconds
    if set, else the effort tier's. The runner records the result on every
    attempt row (extra_configs.api_timeout_seconds)."""
    if explicit:
        return int(explicit)
    return TIMEOUT_BY_REASONING.get(reasoning_effort, TIMEOUT_BY_REASONING[None])


# --- Live pricing -----------------------------------------------------------
# OpenRouter publishes current per-token prices for every model it routes
# (including Anthropic/OpenAI list prices) at a public, unauthenticated
# endpoint. Fetched once per process so each run prices with current data
# instead of a hand-maintained table.
_LIVE_PRICING_URL = "https://openrouter.ai/api/v1/models"
_live_pricing: Optional[Dict[str, Dict[str, float]]] = None
_live_pricing_attempted = False
_pricing_warned: set = set()


def _fetch_live_pricing(timeout: int = 10) -> Optional[Dict[str, Dict[str, float]]]:
    """Fetch {model_id: {input, output}} in $/MTok. Cached for the process."""
    global _live_pricing, _live_pricing_attempted
    if _live_pricing_attempted:
        return _live_pricing
    _live_pricing_attempted = True
    try:
        with urllib.request.urlopen(_LIVE_PRICING_URL, timeout=timeout) as resp:
            data = json.load(resp)
        pricing = {}
        for entry in data.get("data", []):
            p = entry.get("pricing") or {}
            try:
                pricing[entry["id"]] = {
                    "input": float(p["prompt"]) * 1_000_000,
                    "output": float(p["completion"]) * 1_000_000,
                    "context": int(entry["context_length"] or 0) or None,
                }
            except (KeyError, TypeError, ValueError):
                continue
        _live_pricing = pricing or None
        if _live_pricing:
            print(f"💲 Live pricing loaded from OpenRouter ({len(pricing)} models)")
    except Exception as e:
        print(
            f"⚠️ Live pricing fetch failed ({e}); using static MODEL_PRICING fallback"
        )
        _live_pricing = None
    return _live_pricing


def _candidate_slugs(model: str) -> list:
    """Map a model name to the OpenRouter ids it may be listed under.

    Direct-API ids differ from OpenRouter slugs: "claude-fable-5" is listed
    as "anthropic/claude-fable-5", "claude-opus-4-8" as
    "anthropic/claude-opus-4.8", "gpt-5.6-sol" as "openai/gpt-5.6-sol",
    "grok-4.6" as "x-ai/grok-4.6", "Kimi-K3" as "moonshotai/kimi-k3".
    """
    candidates = [model]
    # TensorBlock Forge ids carry a "tensorblock/" prefix over the bare
    # direct-API name ("tensorblock/gpt-5.6-sol"); price and context come
    # from the bare id's usual sources. Forge publishes no prices, so this
    # is the OpenRouter/list rate, not necessarily what Forge credits charge.
    if model.startswith("tensorblock/"):
        bare = model[len("tensorblock/") :]
        candidates.append(bare)
        model = bare
    dotted = re.sub(r"-(\d+)-(\d+)$", r"-\1.\2", model)
    if dotted != model:
        candidates.append(dotted)
    for name in list(candidates):
        if "/" not in name:
            if name.startswith("claude"):
                candidates.append(f"anthropic/{name}")
            elif name.startswith(("gpt", "o1", "o3", "o4", "chatgpt")):
                candidates.append(f"openai/{name}")
            elif name.startswith("grok"):
                candidates.append(f"x-ai/{name}")
            elif name.lower().startswith("kimi"):
                # Forge spells it "Kimi-K3"; OpenRouter lists "moonshotai/kimi-k3".
                candidates.append(f"moonshotai/{name.lower()}")
    return candidates


# Prices verified against ACTUAL billing for direct-API runs. These outrank
# the OpenRouter live feed, which lists brokered-route prices that can
# differ from what the provider bills a direct API key. Keys are bare
# direct-API ids (no "/"), which never name an OpenRouter-routed run, so
# these overrides cannot touch OpenRouter cohorts (whose real billed cost
# is preferred via usage.include anyway).
#
# claude-fable-5: calibrated 2026-08-27 against a Console credit diff over
# a single isolated run (attempt 468: 3,214,460 in / 439,972 out drew
# $102.58; $25/$50 predicts $102.36, +0.2% residual; the $10/$50 the live
# feed lists for anthropic/claude-fable-5 predicts $54.14, -47%). Single
# calibration point — input/output split assumes output stayed at the
# listed $50; re-verify against the next isolated run's credit diff.
#
# Forge cohorts, 2026-09-21: Forge publishes no prices, but its usage log
# (GET /v1/statistic/usage/download with the API key) lists every call's
# tokens and charge. Keys are Forge's bare ids, looked up with the
# "tensorblock/" prefix stripped.
#   Kimi-K3      (served by Fireworks, echoes "FW-Kimi-K3"): all 79 billed
#                calls = $3.30 in / $16.50 out exactly (cached input $0.33) -
#                1.94x OpenRouter's moonshotai/kimi-k3 list ($1.70/$8.50),
#                the rate rows 1881 and 1944 were recorded at. Not in any
#                rate: Forge also charged $2.21 for each call it left
#                unanswered for 600 s (zero tokens; 14 of them on task 41).
#   gemini-3.8-flash  (upstream "models/gemini-3.8-flash"): all 34 billed
#                calls to 2026-09-21 = $0.75 in / $3.75 out exactly - Google's
#                introductory price, which it says runs to 2026-12-31 ($1.50 /
#                $7.50 from 2027-01-01). The live feed never matches the bare
#                id (OpenRouter lists google/gemini-3.8-flash), so without this
#                entry every row would cost $0.
#   qwen3.8-max  every Forge-billed probe call on 2026-09-21 (16 calls) =
#                $2.00 in / $6.00 out exactly, cached input $0.25. The live
#                feed never matches the bare id (OpenRouter lists
#                qwen/qwen3.8-max), so without this entry every row would
#                cost $0.
#   claude-opus-5  the 2026-09-23 probe calls = $5.00 in / $25.00 out exactly
#                (Anthropic's list price, also OpenRouter's
#                anthropic/claude-opus-5); no cached tokens reported.
#   glm-5.3      (served by Fireworks, echoes "FW-GLM-5.3"): every billed call
#                in the usage log to 2026-09-24 = $1.54 in / $4.84 out exactly,
#                cached input $0.286 (OpenRouter's z-ai/glm-5.3: $1.40 / $4.40).
#                The live feed never matches the bare id.
DIRECT_API_PRICING = {
    "claude-fable-5": {"input": 25.00, "output": 50.00},
    "Kimi-K3": {"input": 3.30, "output": 16.50},
    "gemini-3.8-flash": {"input": 0.75, "output": 3.75},
    "qwen3.8-max": {"input": 2.00, "output": 6.00},
    "claude-opus-5": {"input": 5.00, "output": 25.00},
    "glm-5.3": {"input": 1.54, "output": 4.84},
}


def resolve_pricing(model: str) -> Optional[Dict[str, float]]:
    """Return {input, output} $/MTok for a model, preferring live data.

    Order: DIRECT_API_PRICING (billing-verified direct-API rates) -> live
    OpenRouter prices (exact id, then direct-API id mapped to its
    OpenRouter slug) -> static MODEL_PRICING -> None. Falling back or
    failing to price is warned once per model so silent 0-cost rows can't
    accumulate unnoticed.
    """
    if model in DIRECT_API_PRICING:
        return DIRECT_API_PRICING[model]
    if (
        model.startswith("tensorblock/")
        and model[len("tensorblock/") :] in DIRECT_API_PRICING
    ):
        return DIRECT_API_PRICING[model[len("tensorblock/") :]]
    live = _fetch_live_pricing()
    if live:
        for slug in _candidate_slugs(model):
            if slug in live:
                return live[slug]
    # A Forge id falls back to its bare id's static price, as it does for the
    # direct-API and context-window tables.
    static_id = model
    if model not in MODEL_PRICING and model.startswith("tensorblock/"):
        static_id = model[len("tensorblock/") :]
    if static_id in MODEL_PRICING:
        if live and model not in _pricing_warned:
            _pricing_warned.add(model)
            print(f"⚠️ {model} not in live pricing; using static table price")
        return MODEL_PRICING[static_id]
    if model not in _pricing_warned:
        _pricing_warned.add(model)
        print(f"⚠️ No pricing found for {model}; cost will be logged as 0.0")
    return None


# FALLBACK ONLY: resolve_context_window() prefers the live context_length
# from the same OpenRouter fetch; these cover the direct-API ids in use when
# that fetch fails.
MODEL_CONTEXT_WINDOWS = {
    "claude-fable-5": 200_000,
    "claude-opus-4-8": 200_000,
    "claude-opus-4-6": 200_000,
    "gpt-5.6-sol": 400_000,
    # 2026-09-19: equal to the live feed's value on that day, so an attempt
    # whose fetch failed sees the same workbook context as one whose fetch
    # worked. Without an entry the 128k default minus the 128k output
    # allowance floors the context budget at 10k tokens - a crippled attempt.
    "claude-fable-5-1": 1_000_000,
    "gpt-6-astra": 1_050_000,
    # 2026-09-20: xAI's documented window, equal to the live feed's value for
    # x-ai/grok-4.6 (reached as tensorblock/grok-4.6; the prefix is stripped).
    "grok-4.6": 500_000,
    # 2026-09-20: the live feed's value for moonshotai/kimi-k3 (reached as
    # tensorblock/Kimi-K3). Fireworks' own limit is not published through
    # Forge; a provider context-length 400 still tightens the budget.
    "Kimi-K3": 1_048_576,
    # 2026-09-21: OpenRouter's value for google/gemini-3.8-flash (Google: "1M
    # token context"), reached as tensorblock/gemini-3.8-flash. The live feed
    # never matches this id; without the entry the 128k default would squeeze
    # the workbook context to 10k tokens.
    "gemini-3.8-flash": 1_048_576,
    # 2026-09-24: what TensorBlock accepts for tensorblock/qwen3.8-max, not the
    # published 1M (OpenRouter's value, used here from 2026-09-21): it took a
    # 255,886-token prompt and refused ~268k with its generic 400, which the
    # overflow rescue only reads as "too long" above half this window. The
    # two task-8 rows (2206, 3629) sent ~20-40k tokens, so this changes
    # nothing they saw. The live feed never matches this id.
    "qwen3.8-max": 255_000,
    # 2026-09-23: Anthropic's documented window for claude-opus-5, equal to the
    # live feed's value for anthropic/claude-opus-5 (reached as
    # tensorblock/claude-opus-5). Without the entry a failed fetch would squeeze
    # the workbook context to 10k tokens.
    "claude-opus-5": 1_000_000,
    # 2026-09-24: what TensorBlock accepts for tensorblock/glm-5.3 (served by
    # Fireworks): a 795,002-token prompt went through, ~894k was refused (429,
    # as for Kimi); published 1,048,576. The live feed never matches the id;
    # without the entry the 128k default would squeeze the workbook context
    # to 10k tokens.
    "glm-5.3": 795_000,
}
DEFAULT_CONTEXT_WINDOW = 128_000


def resolve_context_window(model: str) -> int:
    """Return the model's input context window in tokens.

    Order: live OpenRouter context_length -> static MODEL_CONTEXT_WINDOWS ->
    DEFAULT_CONTEXT_WINDOW (conservative). Warned once per model on fallback.
    """
    live = _fetch_live_pricing()
    if live:
        for slug in _candidate_slugs(model):
            ctx = (live.get(slug) or {}).get("context")
            if ctx:
                return ctx
    if model in MODEL_CONTEXT_WINDOWS:
        return MODEL_CONTEXT_WINDOWS[model]
    if (
        model.startswith("tensorblock/")
        and model[len("tensorblock/") :] in MODEL_CONTEXT_WINDOWS
    ):
        return MODEL_CONTEXT_WINDOWS[model[len("tensorblock/") :]]
    key = f"ctx:{model}"
    if key not in _pricing_warned:
        _pricing_warned.add(key)
        print(
            f"⚠️ No context window found for {model}; "
            f"assuming {DEFAULT_CONTEXT_WINDOW:,} tokens"
        )
    return DEFAULT_CONTEXT_WINDOW


def calculate_cost(model: str, prompt_tokens: int, completion_tokens: int) -> float:
    """Calculate cost in USD for an API call.

    Args:
        model: Model name (e.g., "claude-fable-5", "openai/gpt-5.2")
        prompt_tokens: Number of input tokens
        completion_tokens: Number of output tokens

    Returns:
        Cost in USD (0.0 when no price is known — warned by resolve_pricing)
    """
    pricing = resolve_pricing(model)
    if not pricing:
        return 0.0

    input_cost = (prompt_tokens / 1_000_000) * pricing["input"]
    output_cost = (completion_tokens / 1_000_000) * pricing["output"]

    return input_cost + output_cost


def get_model_defaults(model: str) -> Dict:
    """Get per-model default parameters.

    Args:
        model: Model name/slug

    Returns:
        Dict of default parameters for the model (empty if no special defaults)
    """
    return MODEL_DEFAULTS.get(model, {})
