# MBABench

MBABench asks an AI agent to do what a junior analyst does: take a business case
and a starting workbook, and build a complete, working financial model in Excel.
It has 101 tasks, four ways of running an agent on them, and an LLM judge that
grades every attempt against a golden solution with a 132-check rubric.

- **Tasks** are on Hugging Face: [namkoong-lab/MBABench](https://huggingface.co/datasets/namkoong-lab/MBABench)
  (starting workbook, golden solution and metadata per task).
- **Agents** reach a model through four surfaces: the vendors' chat products in a browser
  (`gui-agents/`), their add-ins inside Excel Online (`excel-agents/`), raw model APIs driving an
  Excel tool server (`cli-agents/`), and coding agents in a Docker sandbox (`coding-agents/`).
- **Judge** (`judge/`): one LLM conversation per attempt plus deterministic Python checks for the
  Questions-sheet answers and 20 formatting and structure checks.

```text
data/            the downloaded tasks            (scripts/download_dataset.py)
outputs/         attempts and gradings           (written by runs)
config/          config_default.yaml + your config.yaml
house_standards/ the modelling conventions every agent receives with its task
cli-agents/  coding-agents/  gui-agents/  excel-agents/  judge/
```

## Install

You need Python 3.12+, [uv](https://docs.astral.sh/uv/) and LibreOffice (`soffice`, for formula
recalculation). The coding pipeline needs Docker; the GUI and Excel pipelines need Google Chrome
and a paid account with the vendor.

```bash
git clone https://github.com/namkoong-lab/MBABench && cd MBABench
./setup.sh                                   # creates .venv, installs everything, creates config/config.yaml
hf auth login                                # once; the dataset is private
uv run python scripts/download_dataset.py    # -> data/, verified against MANIFEST.json
```

## Configure

`config/config.yaml` (gitignored) holds everything that is yours. Fill in only what you use:

```yaml
keys:
  anthropic_api_key: "${env:ANTHROPIC_API_KEY}"   # or paste the key; CLI + coding pipelines on Claude
  openai_api_key: "${env:OPENAI_API_KEY}"         # CLI + coding pipelines on OpenAI, and the judge
local:
  data_root: "data"       # where download_dataset.py put the tasks
  output_root: "outputs"  # where attempts and gradings go
libreoffice_path: null    # null = find soffice on PATH / in /Applications
```

The GUI and Excel pipelines need no key, only a Chrome signed in to the vendor; their
machine settings (Chrome port, profile, OneDrive folder) go in `<pipeline>/infra/configs/configs.yaml`.

## Run an agent

Every pipeline takes one run config that names the cohort by `agent_model_name` (an entry in the
pipeline's `agent_identities.yaml`, which pins model, effort and settings) and lists the task ids
to run (default: all 101). Runs resume: a task that already has a successful attempt for that
cohort is skipped. Attempts land under `outputs/attempts/<agent_model_name>/`.

**CLI** — a model API plus a local Excel tool server (MCP). Needs the provider key.

```bash
cd cli-agents
uv run excel-agent --batch-config examples/anthropic_claude_fable_5_1_max.yaml --dry-run
uv run excel-agent --batch-config examples/anthropic_claude_fable_5_1_max.yaml
```

**Coding** — Claude Code or Codex inside a Docker sandbox, one container per task. Needs Docker
and `ANTHROPIC_API_KEY` or `OPENAI_API_KEY`.

```bash
cd coding-agents
docker build -t mbabench-coding-agent:v2 docker/
uv run python -m coding_agent.run_sweep --dry-run --config run_configs/claude_code_fable_5_1_max.yaml
uv run python -m coding_agent.run_sweep           --config run_configs/claude_code_fable_5_1_max.yaml
uv run python -m coding_agent.run_sweep --task-ids 1,2 --config run_configs/codex_gpt_6_astra_xhigh.yaml
```

**GUI** — claude.ai or chatgpt.com driven through a real Chrome. Start Chrome with remote
debugging on the port in `infra/configs/configs.yaml`, sign in, then:

```bash
cd gui-agents
uv run python -m infra.run --dry-run --run-config infra/configs/run_configs/claude_cowork_fable_5_1_max.yaml
uv run python -m infra.run -y        --run-config infra/configs/run_configs/claude_cowork_fable_5_1_max.yaml
```

**Excel** — the Claude or ChatGPT add-in inside Excel Online. Chrome must be signed in to
Microsoft 365 with the add-in installed, and the task workbooks must sit in OneDrive:

```bash
cd excel-agents
scripts/setup_chrome.sh                                   # launches Chrome on the CDP port
uv run python scripts/provision_onedrive.py --dry-run     # then without --dry-run: uploads data/ tasks to OneDrive
uv run python -m infra.run --dry-run --run-config infra/configs/run_configs/claude_excel_fable_5_1.yaml
uv run python -m infra.run -y        --run-config infra/configs/run_configs/claude_excel_fable_5_1.yaml
```

`--dry-run` resolves the task, prompt, attachments and output path without touching a model or
a browser. Each example config has a header comment on what to change.

**Try one task first.** A full cohort is 101 attempts; start with one. CLI: set `task_ids: [1]`
in the batch config. Coding: `--task-ids 1`. GUI and Excel: `--task-id 1`. Then grade it with
`grade.py --all` (below) and look at `outputs/gradings/<id>/scores.json`.

## Grade attempts

The judge reads every `outputs/attempts/**/task_attempts.jsonl`, stages each attempt with its
task's golden solution and rubric annotation, and writes `outputs/gradings/gradings.jsonl` plus one
folder of evidence per grading. The default grader is `openai/gpt-5.6-sol` (about $3.5 and 2
minutes per attempt); `judge/judge_identities.yaml` lists the others.

```bash
uv run python judge/main_scripts/grade.py --all --dry-run                 # what would be graded
uv run python judge/main_scripts/grade.py --all                           # grade everything not yet graded
uv run python judge/main_scripts/grade.py --agent-model-name <label> --workers 4
uv run python judge/main_scripts/grade.py --attempt-ids <id> --regrade    # grade again
```

To grade one workbook outside the benchmark flow, assemble a folder with `ai_attempt.xlsx`,
`solution/<golden>.xlsx` and optionally `starting/<workbook>.xlsx`, and run
`uv run python judge/main_scripts/judge.py -f <folder>` (set `JUDGE_SKIP_SUITABILITY=1` for a
task outside the pool). Results land in `<folder>/judge_results/`.

A grading records a verdict per rubric check, a weighted score per category and a total out of
100 (`scores.json`), the Questions-sheet answer check, the deterministic checks, and the judge's
full conversation.

## Output layout

```text
outputs/
  attempts/<agent_model_name>/
    task_attempts.jsonl                   one row per attempt: task_id, files, timing, cost, agent_failed, ...
    task_id=<N>/<YYYYmmdd_HHMMSS>/        the delivered workbook first, then prompts, transcript, logs
  gradings/
    gradings.jsonl                        one row per grading: attempt_id, grader_model, scores, ...
    <grading_id>/                         scores.json, ai_judgement.json, det_checks.json, extracted CSVs, trajectory
```

## Tasks

Each `data/tasks/task_id=<N>/task.json` carries `task_name`, `difficulty` (Easy to Hard),
`difficulty_score`, `model_type`, a one-line `description`, `estimated_solve_time_hours`, and the
`starting_files` / `solution_files` paths. `data/tasks.jsonl` has the same rows in one file.
