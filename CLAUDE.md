# Project context for Claude Code

## What this is
A standalone local desktop app that runs multi-phase Excel transformations.
Deployed per-laptop, pulls logic updates from this GitHub repo on launch.
The main pipeline is **plain deterministic Python** — no server, no LLM
calls. An optional ad-hoc chat panel uses the Anthropic API (requires a
local `ANTHROPIC_API_KEY` in `.env`) but the core flow works without it.

## Architecture
1. **Streamlit UI** (`app.py`) — file upload, analysis parameters, run
   button, run-outputs list, optional ad-hoc chat panel. Runs at
   `localhost:8501`.
2. **Pipeline** (`excel_pipeline/pipeline.py`) — plain Python functions
   (e.g. `run_phase1`) that call skill functions in a fixed, deterministic
   sequence. **This is the main flow; no API key needed.**
3. **Skill** (`excel_pipeline/skill.py`) — plain Python functions that
   read/write `.xlsx`. Every Excel transformation lives here.

### Optional Claude layer (ad-hoc chat only)
`excel_pipeline/orchestrator.py` + `subagents/*.md` + `TOOL_DEFINITIONS` /
`_DISPATCH` in `skill.py` exist only to power the **ad-hoc chat panel**.
It runs a tool-use loop with Claude (model: `claude-sonnet-4-6`) so the
user can request one-off changes in natural language after a run. Fully
optional: the chat panel shows a disabled notice if `ANTHROPIC_API_KEY`
is missing, and the main pipeline is unaffected.

## The discipline rule (non-negotiable)
All Excel work lives in `skill.py` as pure Python functions with tests.
`pipeline.py` orchestrates them deterministically. The optional ad-hoc
agent picks them via tool-use. **Never** emit openpyxl/pandas code as a
string or improvise logic outside the skill.

Preserve this boundary when adding features.

## Adding a new transformation
1. Typed Python function in `skill.py` with a docstring.
2. `pytest` against a fixture workbook in `tests/`.
3. Call it from `pipeline.py` in the appropriate phase (deterministic flow).
4. *(Only if needed for ad-hoc chat)* add a `TOOL_DEFINITIONS` entry and
   a `_DISPATCH` entry, and mention it in the relevant subagent prompt.

## Run commands
- App: `uv run streamlit run app.py`
- Tests: `uv run pytest`
- Sync deps: `uv sync`
- Update to latest from GitHub: `git pull && uv sync`

## File layout
```
app.py                         Streamlit entry point
excel_pipeline/
  pipeline.py                  deterministic Python flow (main pipeline)
  skill.py                     Excel operations (all transformations here)
  orchestrator.py              OPTIONAL Claude runner — ad-hoc chat only
  subagents/
    phase1.md                  legacy / reference subagent prompt
    phase2.md                  legacy / reference subagent prompt
    ad_hoc.md                  ad-hoc chat system prompt
tests/                         pytest fixtures + tests
runs/                          per-run input/output folders (gitignored)
launch.bat / launch.sh         installer + updater + app launcher
pyproject.toml                 uv-managed deps
```

## Things to avoid
- Committing anything to `runs/` (user data — `.gitignore` handles this
  but don't fight it).
- Committing `.env` or any real API key.
- Putting Excel logic in `pipeline.py`, subagent prompts, or anywhere other
  than `skill.py`. Pipelines chain skill calls; skills do the work.
- Having the ad-hoc agent call raw openpyxl or pandas. If you see tool
  names like `run_python` or `execute_code`, something has gone wrong.
- Making transformations non-idempotent without a warning in the docstring.

## Model choice (ad-hoc only)
`claude-sonnet-4-6` by default in `orchestrator.py:MODEL`. Swap to
`claude-opus-4-7` only if ad-hoc requests routinely need more reasoning.
Haiku is usually too weak for tool-use orchestration.
