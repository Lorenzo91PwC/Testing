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
import pandas as pd
from openpyxl.utils import column_index_from_string
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


def find_file_by_suffix(directory: str, suffix: str) -> str | None:
    """Return the path of the .xlsx file in `directory` whose stem ends with
    `suffix`, or None if no match is found.

    Used to locate a specific input file independently of the date prefix in
    its name (e.g. suffix "AAI_P&C_Ceded" matches "1.1_2025.12.31_AAI_P&C_Ceded.xlsx").
    """
    for f in Path(directory).glob("*.xlsx"):
        if f.stem.endswith(suffix):
            return str(f)
    return None


def get_unique_column_values(path: str, sheet: str, column: str) -> list[Any]:
    """Return the sorted unique, non-empty values from a single column.

    `column` may be either a column letter (e.g. "G") or a header name from
    row 1. None and empty-string values are excluded. Values are sorted by
    their string representation for stable output across heterogeneous types.
    """
    wb = load_workbook(path)
    ws = wb[sheet]
    try:
        col_idx = column_index_from_string(column)
    except ValueError:
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        if column not in headers:
            raise ValueError(f"Column '{column}' not found. Headers: {headers}")
        col_idx = headers.index(column) + 1

    values: set[Any] = set()
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=col_idx).value
        if v is None or (isinstance(v, str) and v.strip() == ""):
            continue
        values.add(v)
    return sorted(values, key=str)


def build_goc_year_summary(path: str, sheet: str) -> dict[str, list[dict[str, Any]]]:
    """Aggregate the Ceded sheet by (GoC, year) and return per-GoC tables.

    Expected input columns (by header, present in sheet `AAI_P&C_Ceded_H_NH`):
    - `GoC`              — column AA, grouping key
    - `anno`             — column AB, reporting year
    - `Sinistri`         — column AD, summed and returned as "Paid Claims"
    - `Riserve Sinistri` — column AE, summed and returned as "Claims Reserve"

    Non-numeric values in the sum columns are coerced to 0. Rows with an
    empty GoC are dropped. Output shape:

        {
          "<GoC>": [
            {"year": 2017, "Paid Claims": 123.45, "Claims Reserve": 0.0},
            ...
          ],
          ...
        }
    """
    df = pd.read_excel(path, sheet_name=sheet)

    required = ["GoC", "anno", "Sinistri", "Riserve Sinistri"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing columns {missing} in sheet '{sheet}'. "
            f"Got: {list(df.columns)}"
        )

    df = df.dropna(subset=["GoC"])
    df = df[df["GoC"].astype(str).str.strip() != ""]
    df["Sinistri"] = pd.to_numeric(df["Sinistri"], errors="coerce").fillna(0)
    df["Riserve Sinistri"] = pd.to_numeric(
        df["Riserve Sinistri"], errors="coerce"
    ).fillna(0)

    grouped = (
        df.groupby(["GoC", "anno"], dropna=False)[["Sinistri", "Riserve Sinistri"]]
        .sum()
        .reset_index()
        .rename(
            columns={
                "anno": "year",
                "Sinistri": "Paid Claims",
                "Riserve Sinistri": "Claims Reserve",
            }
        )
    )

    result: dict[str, list[dict[str, Any]]] = {}
    for goc in sorted(grouped["GoC"].unique(), key=str):
        sub = (
            grouped[grouped["GoC"] == goc]
            .drop(columns=["GoC"])
            .sort_values("year")
        )
        result[str(goc)] = sub.to_dict(orient="records")
    return result


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
    {
        "name": "find_file_by_suffix",
        "description": (
            "Find an .xlsx file in a directory whose stem (filename without "
            "extension) ends with a given suffix. Returns the absolute path "
            "or null. Use when a file has a variable date prefix but a stable "
            "suffix (e.g. suffix 'AAI_P&C_Ceded')."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "directory": {"type": "string"},
                "suffix": {"type": "string"},
            },
            "required": ["directory", "suffix"],
        },
    },
    {
        "name": "get_unique_column_values",
        "description": (
            "Return the sorted unique, non-empty values from a single column "
            "of a sheet. `column` can be either a column letter (e.g. 'G') "
            "or a header name from row 1."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "sheet": {"type": "string"},
                "column": {
                    "type": "string",
                    "description": "Column letter (e.g. 'G') or header name.",
                },
            },
            "required": ["path", "sheet", "column"],
        },
    },
    {
        "name": "build_goc_year_summary",
        "description": (
            "Aggregate the AAI P&C Ceded H/NH sheet by (GoC, year). Returns "
            "a mapping of GoC to a list of yearly totals with 'Paid Claims' "
            "(sum of 'Sinistri', column AD) and 'Claims Reserve' (sum of "
            "'Riserve Sinistri', column AE). Year comes from column AB "
            "('anno'); grouping key is column AA ('GoC')."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "sheet": {
                    "type": "string",
                    "description": "Usually 'AAI_P&C_Ceded_H_NH'.",
                },
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
    "find_file_by_suffix": find_file_by_suffix,
    "get_unique_column_values": get_unique_column_values,
    "build_goc_year_summary": build_goc_year_summary,
    # Add an entry for each domain-specific function above.
}


def dispatch_tool(name: str, arguments: dict[str, Any]) -> Any:
    """Route a tool call from Claude to the corresponding Python function."""
    if name not in _DISPATCH:
        raise ValueError(
            f"Unknown tool: {name}. Known: {sorted(_DISPATCH.keys())}"
        )
    return _DISPATCH[name](**arguments)
