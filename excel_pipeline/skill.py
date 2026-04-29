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

from datetime import date, datetime
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
# Astra — MP_GOC transformation
# ===========================================================================
_BUSINESS_TYPE_VALUES = {
    "Direct": "2_RE_ASSUMED",
    "Ceduto": "3_RE_CEDED_NON_RETRO",
}


def _parse_valuation_date(valuation_date: str) -> date:
    return datetime.strptime(valuation_date, "%Y-%m-%d").date()


def _semester_mmdd(val: date) -> str:
    return "0630" if val.month <= 6 else "1231"


def _inception_curve_id(cohort_year: int, mmdd: str) -> str:
    if cohort_year <= 2015:
        return f"{cohort_year}{mmdd}_ITA_LP100"
    if cohort_year <= 2021:
        return f"{cohort_year}{mmdd}_ITA_LP100_AVG"
    if cohort_year == 2022:
        return f"2022{mmdd}_ITA_LP100_FY22_AVG"
    yy = f"{cohort_year % 100:02d}"
    return f"{cohort_year}{mmdd}_EUR_LP100_FY{yy}_AVG"


def astra_transform_mp_goc(
    input_path: str,
    output_path: str,
    valuation_date: str,
    business_type: str,
) -> dict[str, Any]:
    """Apply the Astra MP_GOC column rules and save the result.

    Rules per row (data starts at row 2, single sheet, year is column C):
      - E (INCEPTION_CURVE_ID): year-bucketed string with MMDD = 0630
        if first semester else 1231.
      - F (TIMING_INCEPTION_CURVE): only when C year == 2025, set to
        "7_JULY" (first semester) or "13_YEAR_END" (second). Otherwise
        leave untouched.
      - L (GOC_DURATION): max(0, valuation_year - cohort_year) * 12.
      - P (GOC_TYPE_REINSURANCE): "2_RE_ASSUMED" if business_type=="Direct",
        "3_RE_CEDED_NON_RETRO" if "Ceduto".

    Idempotent: rerunning with the same inputs produces the same output.
    """
    if business_type not in _BUSINESS_TYPE_VALUES:
        raise ValueError(
            f"business_type must be one of {sorted(_BUSINESS_TYPE_VALUES)}, "
            f"got {business_type!r}"
        )
    val = _parse_valuation_date(valuation_date)
    mmdd = _semester_mmdd(val)
    f_value_for_2025 = "7_JULY" if val.month <= 6 else "13_YEAR_END"
    p_value = _BUSINESS_TYPE_VALUES[business_type]

    wb = load_workbook(input_path)
    ws = wb[wb.sheetnames[0]]
    rows_changed = 0
    for r in range(2, ws.max_row + 1):
        cohort_raw = ws.cell(row=r, column=3).value
        if cohort_raw is None or cohort_raw == "":
            continue
        cohort_year = int(cohort_raw)

        ws.cell(row=r, column=5).value = _inception_curve_id(cohort_year, mmdd)
        if cohort_year == 2025:
            ws.cell(row=r, column=6).value = f_value_for_2025
        ws.cell(row=r, column=12).value = max(0, val.year - cohort_year) * 12
        ws.cell(row=r, column=16).value = p_value
        rows_changed += 1

    save_workbook(wb, output_path)
    return {
        "rows_changed": rows_changed,
        "output_path": output_path,
        "valuation_date": valuation_date,
        "business_type": business_type,
    }


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
        "name": "astra_transform_mp_goc",
        "description": (
            "Apply the Astra MP_GOC column rules to a workbook and save the "
            "result. Operates on the only sheet, header in row 1, data from "
            "row 2. Modifies columns E, F, L, P per the MP_GOC spec, using "
            "the user-provided valuation_date (YYYY-MM-DD) and business_type "
            "('Direct' or 'Ceduto'). Idempotent — rerunning with the same "
            "inputs is safe and produces the same output."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "input_path": {
                    "type": "string",
                    "description": "Absolute path to the source MP_GOC .xlsx.",
                },
                "output_path": {
                    "type": "string",
                    "description": "Absolute path where the transformed .xlsx is saved.",
                },
                "valuation_date": {
                    "type": "string",
                    "description": "Valuation date in ISO format YYYY-MM-DD.",
                },
                "business_type": {
                    "type": "string",
                    "enum": ["Direct", "Ceduto"],
                    "description": "Direct → 2_RE_ASSUMED; Ceduto → 3_RE_CEDED_NON_RETRO.",
                },
            },
            "required": [
                "input_path",
                "output_path",
                "valuation_date",
                "business_type",
            ],
        },
    },
]


# ===========================================================================
# Dispatcher — maps tool names to Python functions
# ===========================================================================
_DISPATCH: dict[str, Any] = {
    "inspect_workbook": inspect_workbook,
    "preview_rows": preview_rows,
    "astra_transform_mp_goc": astra_transform_mp_goc,
}


def dispatch_tool(name: str, arguments: dict[str, Any]) -> Any:
    """Route a tool call from Claude to the corresponding Python function."""
    if name not in _DISPATCH:
        raise ValueError(
            f"Unknown tool: {name}. Known: {sorted(_DISPATCH.keys())}"
        )
    return _DISPATCH[name](**arguments)
