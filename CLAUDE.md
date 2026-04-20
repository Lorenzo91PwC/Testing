# Project context for Claude Code

## What this is
A standalone local desktop app that runs multi-phase Excel transformations.
Deployed per-laptop, pulls logic updates from this GitHub repo on launch,
calls the Anthropic API directly from the user's machine. No server, no
proxy — each user brings their own API key in a local `.env` file.

## Architecture (three layers)
1. **Streamlit UI** (`app.py`) — file upload, run button, per-phase progress,
   ad-hoc chat panel. Runs at `localhost:8501`.
2. **Orchestrator** (`excel_pipeline/orchestrator.py`) — loads a subagent
   prompt from `subagents/{phase}.md`, runs a tool-use loop with Claude
   (model: `claude-sonnet-4-6`), terminates on `end_turn`.
3. **Skill** (`excel_pipeline/skill.py`) — plain Python functions that
   read/write `.xlsx`. Every transformation lives here, with a matching
   entry in `TOOL_DEFINITIONS` and `_DISPATCH`.

## The discipline rule (non-negotiable)
**Subagents decide which function to call and with what arguments.
Skill functions do the Excel work.** The model never writes cell values
directly, never emits openpyxl code as a string, never improvises logic
that isn't in the skill. If the model wants to do something no skill
function supports, the correct answer is "please add a function for this"
— not freestyle.

Preserve this boundary when adding features.

## Adding a new transformation — the four coupled edits
1. Typed Python function in `skill.py` with a docstring.
2. Entry in `TOOL_DEFINITIONS` with a description that teaches Claude
   *when* to use it (and, if relevant, when not to — e.g. idempotency).
3. Entry in `_DISPATCH` mapping the tool name to the function.
4. Mention of the tool in the relevant subagent prompt.

Plus a `pytest` against a fixture workbook.

## Run commands
- App: `uv run streamlit run app.py`
- Tests: `uv run pytest`
- Sync deps: `uv sync`
- Update to latest from GitHub: `git pull && uv sync`

## File layout
```
app.py                         Streamlit entry point
excel_pipeline/
  orchestrator.py              subagent runner + tool-use loop
  skill.py                     Excel operations (all transformations here)
  subagents/
    phase1.md                  Phase 1 system prompt
    phase2.md                  Phase 2 system prompt
    ad_hoc.md                  Ad-hoc edit system prompt
runs/                          per-run input/output folders (gitignored)
launch.bat / launch.sh         installer + updater + app launcher
pyproject.toml                 uv-managed deps
```

## Things to avoid
- Committing anything to `runs/` (user data — `.gitignore` handles this
  but don't fight it).
- Committing `.env` or any real API key.
- Adding logic to subagent prompts that belongs in skill functions.
- Having the model call raw openpyxl or pandas. If you see tool names like
  `run_python` or `execute_code`, something has gone wrong.
- Making transformations non-idempotent without a warning in both the
  docstring and the tool description.

## Model choice
`claude-sonnet-4-6` by default. Swap to `claude-opus-4-7` in
`orchestrator.py:MODEL` if a specific phase needs more reasoning (complex
multi-sheet consolidation, ambiguous ad-hoc requests). Haiku is usually
too weak for tool-use orchestration here.
