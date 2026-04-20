# Excel Pipeline

A local Excel transformation tool powered by Claude. Runs entirely on your
laptop — input and output files never leave the machine.

## What it does

Walks an Excel file through a multi-phase pipeline (Phase 1 → Phase 2), then
lets you make ad-hoc edits via a chat panel. Claude decides **which**
transformations to apply; the actual Excel work is done by Python functions
you can read, test, and trust.

## Prerequisites

- **Git** installed ([git-scm.com](https://git-scm.com/downloads))
- **Anthropic API key** — get one at [console.anthropic.com](https://console.anthropic.com)
- **Windows 10+**, **macOS 12+**, or a recent Linux
- Network access to `api.anthropic.com` and `github.com`
  (ask your IT to allow these if you're behind a corporate firewall)

## First-time setup (5 minutes)

1. **Clone this repo** to a folder on your laptop:
   ```bash
   git clone https://github.com/Lorenzo91PwC/Testing.git excel-pipeline
   cd excel-pipeline
   ```

2. **Create your `.env`** file from the example and fill in your API key:
   ```bash
   cp .env.example .env
   ```
   Then open `.env` in a text editor and paste your key after `ANTHROPIC_API_KEY=`.

3. **Launch:**
   - **Windows:** double-click `launch.bat`
   - **macOS / Linux:** run `./launch.sh` in a terminal

   The first launch installs `uv` and Python dependencies (~1 minute). It
   then opens the app in your browser at http://localhost:8501.

## Daily use

Just double-click the launcher. It will:
- Pull the latest logic from GitHub (new transformations, prompt updates)
- Sync any new dependencies
- Open the app in your browser

In the app:
1. **Upload** your Excel file
2. Click **▶ Run pipeline**
3. **Download** the outputs from each phase (buttons appear below)
4. Use the **chat panel** on the right for ad-hoc changes

## Where files live

```
excel-pipeline/
├── runs/                          ← your inputs and outputs
│   └── 2026-04-20_143052/
│       ├── input.xlsx
│       ├── phase1_output.xlsx
│       └── phase2_output.xlsx
├── app.py
└── ...
```

Everything in `runs/` is gitignored — your data never leaves your laptop.

## Updating

The launcher pulls from GitHub on every start. To skip updates (e.g. for
offline work), launch while disconnected — the app runs with whatever's in
the local copy.

## Project layout

```
excel-pipeline/
├── app.py                         ← Streamlit UI
├── launch.bat / launch.sh         ← one-click launchers
├── pyproject.toml                 ← deps (managed by uv)
├── excel_pipeline/
│   ├── orchestrator.py            ← runs subagents in sequence
│   ├── skill.py                   ← Excel operations (add new ones here)
│   └── subagents/                 ← Claude prompts per phase
│       ├── phase1.md
│       ├── phase2.md
│       └── ad_hoc.md
└── runs/                          ← per-run folders (gitignored)
```

## Adding a new Excel transformation

1. Write a Python function in `excel_pipeline/skill.py` with typed args.
2. Add a schema entry in `TOOL_DEFINITIONS` so Claude knows it exists.
3. Add the function to `_DISPATCH` so the orchestrator can call it.
4. Mention it in the relevant subagent prompt so Claude knows when to use it.

The discipline throughout: **subagents decide, skill functions do**.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `uv: command not found` after install | Close and reopen your terminal |
| `Could not pull updates` | You're offline or don't have repo access — the app still runs |
| `ANTHROPIC_API_KEY not found` | Copy `.env.example` to `.env` and add your key |
| Browser doesn't open | Go to http://localhost:8501 manually |
| Corporate firewall blocks install | Ask IT to allow `api.anthropic.com`, `github.com`, `astral.sh`, and `pypi.org` |

## Contributing

Issues and suggestions welcome in the repo's Issues tab.
