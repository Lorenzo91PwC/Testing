"""Orchestrates subagent execution for each phase.

Each phase is a focused Claude call with:
- a system prompt loaded from `subagents/{phase}.md`
- a whitelist of tools (the Excel skill functions)
- a simple tool-use loop that terminates when Claude stops calling tools

No file edits happen in prompts — all changes route through skill functions.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from anthropic import Anthropic

from .skill import TOOL_DEFINITIONS, dispatch_tool

# Update if you want a different model. Sonnet 4.6 is a good default.
MODEL = "claude-sonnet-4-6"
MAX_TOKENS = 4096
MAX_TOOL_ITERATIONS = 30

SUBAGENT_DIR = Path(__file__).parent / "subagents"


def _load_prompt(name: str) -> str:
    """Load a subagent system prompt from its markdown file."""
    path = SUBAGENT_DIR / f"{name}.md"
    return path.read_text(encoding="utf-8")


def _extract_text(content_blocks: list) -> str:
    """Concatenate all text blocks from a Claude response."""
    parts = [b.text for b in content_blocks if getattr(b, "type", None) == "text"]
    return "\n".join(parts)


def _run_tool_loop(system_prompt: str, user_message: str) -> str:
    """Run Claude with tool access until it produces a final text reply.

    Returns the final text the model produced when it stopped calling tools.
    Raises if the loop exceeds MAX_TOOL_ITERATIONS (usually a sign of a bad
    prompt or a tool that errors in a way the model can't recover from).
    """
    client = Anthropic()
    messages: list[dict[str, Any]] = [
        {"role": "user", "content": user_message}
    ]

    for _ in range(MAX_TOOL_ITERATIONS):
        response = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=system_prompt,
            tools=TOOL_DEFINITIONS,
            messages=messages,
        )

        # Done: Claude has nothing more to do and returned final text.
        if response.stop_reason == "end_turn":
            return _extract_text(response.content)

        # Claude wants to call one or more tools. Execute them and feed back.
        if response.stop_reason == "tool_use":
            messages.append({"role": "assistant", "content": response.content})
            tool_results = []
            for block in response.content:
                if getattr(block, "type", None) != "tool_use":
                    continue
                try:
                    result = dispatch_tool(block.name, block.input)
                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": json.dumps(result, default=str),
                    })
                except Exception as e:
                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": f"Error: {type(e).__name__}: {e}",
                        "is_error": True,
                    })
            messages.append({"role": "user", "content": tool_results})
            continue

        # Any other stop reason is unexpected — return whatever text we have.
        return _extract_text(response.content) or f"Stopped: {response.stop_reason}"

    raise RuntimeError(
        f"Tool loop exceeded {MAX_TOOL_ITERATIONS} iterations. "
        "Check subagent prompt and tool error handling."
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def run_phase(phase: str, input_path: Path, run_dir: Path) -> Path:
    """Run a phase subagent and return the path to its output file.

    The subagent is expected to write `{phase}_output.xlsx` in the run dir.
    """
    system_prompt = _load_prompt(phase)
    user_message = (
        f"Input file: {input_path}\n"
        f"Run directory: {run_dir}\n"
        f"Phase: {phase}\n\n"
        f"Apply the transformations described in your instructions, then "
        f"save the output as {phase}_output.xlsx in the run directory."
    )
    _run_tool_loop(system_prompt, user_message)

    output = run_dir / f"{phase}_output.xlsx"
    if not output.exists():
        raise RuntimeError(
            f"{phase} did not produce the expected output at {output}. "
            "Check the subagent prompt and the skill functions it called."
        )
    return output


def run_ad_hoc(request: str, run_dir: Path) -> str:
    """Run the ad-hoc subagent against a completed run folder.

    Returns the assistant's reply describing what it did.
    """
    system_prompt = _load_prompt("ad_hoc")
    user_message = (
        f"Run directory: {run_dir}\n"
        f"User request: {request}\n\n"
        f"Apply the requested change using skill functions only. "
        f"Save as a new versioned file (e.g. phase2_output_v2.xlsx). "
        f"Never overwrite an existing file."
    )
    return _run_tool_loop(system_prompt, user_message)
