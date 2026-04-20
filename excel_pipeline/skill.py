"""Excel skill — deterministic, typed operations on workbooks.

Design rule for this module:
-----------------------------
    Subagents pick WHICH function to call and WITH WHAT ARGUMENTS.
    They never write cell values directly. Every transformation lives
    here as a Python function with tests.

Adding a new transformation (checklist):
  1. Write a plain Python function below. Type its args.
  2. Add a `pytest` in `tests/` that proves it works on a fixture.
  3. Add an entry in `TOOL_DEFINITIONS` with a clear description + schema.
  4. Add an entry in `_DISPATCH` mapping the tool name to the function.
  5. Mention the function in the relevant subagent prompt.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.workbook import Workbook


# ===========================================================================
# Low-level helpers (not exposed as tools — used by other skill functions)
# ===========================================================================
def load_workbook(path: str) -> Workbook:
    """Open an Excel file for reading or editing."""
    return openpyxl.load_workbook(path)


def save_workbook(wb: Workbook, path: str) -> None:
    """Save a workbook to disk, creating parent dirs if needed."""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def list_run_files(run_dir: Path) -> list[Path]:
    """List all .xlsx files in a run directory, sorted by name."""
    return sorted(run_dir.glob("*.xlsx"))


# ===========================================================================
# Tools exposed to Claude — inspection
# ===========================================================================
def inspect_workbook(path: str) -> dict[str, Any]:
    """Summarise a workbook: sheet names, sizes, and column headers.

    This is the first call the subagent makes on any file — it tells Claude
    what it's working with so it can pick the right transformations.
    """
    wb = load_workbook(path)
    sheets = []
    for name in wb.sheetnames:
        ws = wb[name]
        headers = [
            ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)
        ]
        sheets.append({
            "name": name,
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "headers": headers,
        })
    return {"path": path, "sheets": sheets}


def preview_rows(path: str, sheet: str, n: int = 5) -> dict[str, Any]:
    """Return the first N rows of a sheet as a list of dicts.

    Useful for the subagent to sanity-check data shape before transforming.
    """
    wb = load_workbook(path)
    ws = wb[sheet]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, min(2 + n, ws.max_row + 1)):
        row = {
            headers[c - 1]: ws.cell(row=r, column=c).value
            for c in range(1, ws.max_column + 1)
        }
        rows.append(row)
    return {"sheet": sheet, "headers": headers, "rows": rows}


# ===========================================================================
# TODO — Domain-specific transformations
# ===========================================================================
# Add the customer's actual Phase 1 / Phase 2 / ad-hoc operations here.
# Each should be a pure function with typed arguments, with a matching
# entry in TOOL_DEFINITIONS and _DISPATCH below.
#
# Template:
#
#   def my_transformation(
#       input_path: str,
#       output_path: str,
#       sheet: str,
#       # ... typed parameters ...
#   ) -> dict[str, Any]:
#       """One-line description of what this does."""
#       wb = load_workbook(input_path)
#       ws = wb[sheet]
#       # ... work ...
#       save_workbook(wb, output_path)
#       return {"rows_changed": n, "output_path": output_path}
#
# ===========================================================================


# ===========================================================================
# Tool schemas for Claude (Anthropic tool-use format)
# ===========================================================================
TOOL_DEFINITIONS: list[dict[str, Any]] = [
    {
        "name": "inspect_workbook",
        "description": (
            "Summarise an Excel file: sheet names, row/column counts, and "
            "column headers. Call this first to understand the file's shape."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "Absolute path to the .xlsx file.",
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "preview_rows",
        "description": (
            "Return the first N rows of a sheet as a list of dicts keyed by "
            "column header. Use to inspect data values before transforming."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "sheet": {"type": "string"},
                "n": {"type": "integer", "default": 5},
            },
            "required": ["path", "sheet"],
        },
    },
    # Add a schema entry for each domain-specific function above.
]


# ===========================================================================
# Dispatcher — maps tool names to Python functions
# ===========================================================================
_DISPATCH: dict[str, Any] = {
    "inspect_workbook": inspect_workbook,
    "preview_rows": preview_rows,
    # Add an entry for each domain-specific function above.
}


def dispatch_tool(name: str, arguments: dict[str, Any]) -> Any:
    """Route a tool call from Claude to the corresponding Python function."""
    if name not in _DISPATCH:
        raise ValueError(
            f"Unknown tool: {name}. Known: {sorted(_DISPATCH.keys())}"
        )
    return _DISPATCH[name](**arguments)
