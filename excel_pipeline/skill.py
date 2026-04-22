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


def extract_unique_goc_names(
    path: str,
    sheet: str = "AAI_P&C_Ceded_H_NH",
    column: str = "AA",
    start_row: int = 3,
) -> dict[str, Any]:
    """Return the unique non-empty GoC names from a column of a Ceded workbook.

    Intended for the input file whose name ends with the fixed suffix
    ``AAI_P&C_Ceded`` (e.g. ``1.1_2025.12.31_AAI_P&C_Ceded.xlsx``). The
    default sheet, column and ``start_row`` (3, since rows 1-2 are header
    / sub-header) match that file's layout. Read-only and idempotent.
    Order follows first occurrence; whitespace is stripped and empty
    cells are skipped.
    """
    # data_only=True returns the cached computed value for formula cells
    # rather than the formula string — required for the Ceded file whose
    # column AA is typically populated by lookup formulas.
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    col_idx = column_index_from_string(column)
    seen: list[str] = []
    seen_set: set[str] = set()
    for r in range(start_row, ws.max_row + 1):
        v = ws.cell(row=r, column=col_idx).value
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        if s not in seen_set:
            seen_set.add(s)
            seen.append(s)
    return {
        "sheet": sheet,
        "column": column,
        "count": len(seen),
        "values": seen,
    }


def create_mp_lob(
    goc_names: list[str],
    entity_id: int,
    output_path: str,
) -> dict[str, Any]:
    """Create an ``MP_LoB`` workbook with two columns: ``GoC_ID`` and ``Entity_ID``.

    ``GoC_ID`` is filled with the supplied unique GoC names (typically the
    ``values`` result of ``extract_unique_goc_names``). ``Entity_ID`` is the
    selected entity code, repeated on every data row. Overwrites the
    output file if it already exists.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MP_LoB"
    ws.cell(row=1, column=1, value="GoC_ID")
    ws.cell(row=1, column=2, value="Entity_ID")
    for i, name in enumerate(goc_names, start=2):
        ws.cell(row=i, column=1, value=name)
        ws.cell(row=i, column=2, value=entity_id)
    save_workbook(wb, output_path)
    return {
        "output_path": output_path,
        "rows": len(goc_names),
        "columns": ["GoC_ID", "Entity_ID"],
    }


def create_mp_observation_year(
    goc_names: list[str],
    year: int,
    output_path: str,
) -> dict[str, Any]:
    """Create an ``MP_ObservationYear`` workbook.

    Two rows are written per GoC — an ``Opening`` row (year - 1) and a
    ``Closing`` row (year). Columns: ``ObservationID`` (``{goc}@Opening``
    or ``{goc}@Closing``), ``ObservationYear``, ``LoB_ID`` (the GoC),
    ``AdjULAEPagate`` (always ``0``), ``CY`` (always ``Yes``). Overwrites
    the output file if it already exists.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MP_ObservationYear"
    headers = ["ObservationID", "ObservationYear", "LoB_ID", "AdjULAEPagate", "CY"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=h)

    row = 2
    for name in goc_names:
        ws.cell(row=row, column=1, value=f"{name}@Opening")
        ws.cell(row=row, column=2, value=year - 1)
        ws.cell(row=row, column=3, value=name)
        ws.cell(row=row, column=4, value=0)
        ws.cell(row=row, column=5, value="Yes")
        row += 1

        ws.cell(row=row, column=1, value=f"{name}@Closing")
        ws.cell(row=row, column=2, value=year)
        ws.cell(row=row, column=3, value=name)
        ws.cell(row=row, column=4, value=0)
        ws.cell(row=row, column=5, value="Yes")
        row += 1

    save_workbook(wb, output_path)
    return {
        "output_path": output_path,
        "rows": 2 * len(goc_names),
        "columns": headers,
    }


def lookup_risk_adjustment_values(
    path: str,
    goc_names: list[str],
    year: int,
    semester: int,
    sheet: str = "ra_AAI_REINS",
    goc_column: str = "G",
    header_row: int = 1,
) -> dict[str, dict[str, Any]]:
    """Look up Opening/Closing Risk Adjustment values for each GoC.

    Opens a Payment_Patterns_&_Risk_Adjustments workbook with
    ``data_only=True``, locates the table that starts in ``goc_column``
    on ``sheet``, and picks the two year columns matching the selected
    period: ``{prefix}_{year}`` for Closing and ``{prefix}_{year-1}``
    for Opening, where ``prefix = 'HY' if semester == 1 else 'FY'``.

    Returns ``{goc: {"opening": value, "closing": value}}``. GoCs missing
    from the sheet map to ``{"opening": None, "closing": None}``. Raises
    ``KeyError`` if either year column is absent from the header row.
    """
    if semester not in (1, 2):
        raise ValueError(f"semester must be 1 or 2, got {semester}")
    prefix = "HY" if semester == 1 else "FY"
    closing_col_name = f"{prefix}_{year}"
    opening_col_name = f"{prefix}_{year - 1}"

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    goc_col_idx = column_index_from_string(goc_column)

    header_to_idx: dict[str, int] = {}
    for c in range(goc_col_idx + 1, ws.max_column + 1):
        h = ws.cell(row=header_row, column=c).value
        if h is not None:
            header_to_idx[str(h).strip()] = c

    for needed in (opening_col_name, closing_col_name):
        if needed not in header_to_idx:
            raise KeyError(
                f"Column '{needed}' not found in sheet '{sheet}'. "
                f"Available year columns: {sorted(header_to_idx.keys())}"
            )

    goc_to_row: dict[str, int] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=goc_col_idx).value
        if v is None:
            continue
        key = str(v).strip()
        if key and key not in goc_to_row:
            goc_to_row[key] = r

    result: dict[str, dict[str, Any]] = {}
    for goc in goc_names:
        row = goc_to_row.get(goc)
        if row is None:
            result[goc] = {"opening": None, "closing": None}
            continue
        result[goc] = {
            "opening": ws.cell(row=row, column=header_to_idx[opening_col_name]).value,
            "closing": ws.cell(row=row, column=header_to_idx[closing_col_name]).value,
        }
    return result


def create_risk_adjustment(
    goc_names: list[str],
    values: dict[str, dict[str, Any]],
    output_path: str,
) -> dict[str, Any]:
    """Create a ``Risk_Adjustment`` workbook with two columns.

    ``ObservationID`` follows the ``{goc}@Opening`` / ``{goc}@Closing``
    pattern; ``Risk_Adjustment`` pulls from ``values`` (the dict returned
    by ``lookup_risk_adjustment_values``). Missing values are written as
    empty cells. Overwrites the output file if it already exists.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Risk_Adjustment"
    ws.cell(row=1, column=1, value="ObservationID")
    ws.cell(row=1, column=2, value="Risk_Adjustment")

    row = 2
    for name in goc_names:
        vals = values.get(name, {"opening": None, "closing": None})
        ws.cell(row=row, column=1, value=f"{name}@Opening")
        ws.cell(row=row, column=2, value=vals.get("opening"))
        row += 1
        ws.cell(row=row, column=1, value=f"{name}@Closing")
        ws.cell(row=row, column=2, value=vals.get("closing"))
        row += 1

    save_workbook(wb, output_path)
    return {
        "output_path": output_path,
        "rows": 2 * len(goc_names),
        "columns": ["ObservationID", "Risk_Adjustment"],
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
        "name": "extract_unique_goc_names",
        "description": (
            "Return the unique non-empty values in column AA of sheet "
            "'AAI_P&C_Ceded_H_NH' — these are the GoC names. Use this on "
            "the input file whose name ends with the fixed suffix "
            "'AAI_P&C_Ceded' (e.g. '1.1_2025.12.31_AAI_P&C_Ceded.xlsx'). "
            "Default sheet and column match that file's layout; only "
            "override them if the file deviates. Read-only and idempotent."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "Absolute path to the Ceded .xlsx file.",
                },
                "sheet": {
                    "type": "string",
                    "default": "AAI_P&C_Ceded_H_NH",
                },
                "column": {
                    "type": "string",
                    "default": "AA",
                    "description": "Excel column letter to read GoC names from.",
                },
                "start_row": {
                    "type": "integer",
                    "default": 3,
                    "description": "First data row (rows 1-2 are header).",
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "create_mp_lob",
        "description": (
            "Create an 'MP_LoB' workbook with two columns: 'GoC_ID' "
            "(unique GoC names) and 'Entity_ID' (the analyzed company's "
            "code, repeated on every data row). Typically called after "
            "`extract_unique_goc_names`: pass its `values` list as "
            "`goc_names` and the session's entity code as `entity_id`. "
            "Overwrites the output file if it already exists."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "goc_names": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Unique GoC names to use as GoC_ID values.",
                },
                "entity_id": {
                    "type": "integer",
                    "description": "Entity code of the company being analyzed.",
                },
                "output_path": {
                    "type": "string",
                    "description": "Absolute path where MP_LoB.xlsx will be saved.",
                },
            },
            "required": ["goc_names", "entity_id", "output_path"],
        },
    },
    {
        "name": "create_mp_observation_year",
        "description": (
            "Create an 'MP_ObservationYear' workbook with five columns: "
            "ObservationID ('{goc}@Opening' or '{goc}@Closing'), "
            "ObservationYear (year-1 for Opening, year for Closing), "
            "LoB_ID (the GoC name), AdjULAEPagate (always 0), CY (always "
            "'Yes'). Two rows per GoC. Typically called after "
            "`extract_unique_goc_names`. Overwrites the output file."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "goc_names": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Unique GoC names.",
                },
                "year": {
                    "type": "integer",
                    "description": "Analysis year (used for Closing; Opening = year-1).",
                },
                "output_path": {
                    "type": "string",
                    "description": "Absolute path where MP_ObservationYear.xlsx will be saved.",
                },
            },
            "required": ["goc_names", "year", "output_path"],
        },
    },
    {
        "name": "lookup_risk_adjustment_values",
        "description": (
            "Read Opening/Closing Risk Adjustment values for each GoC "
            "from the Payment_Patterns_&_Risk_Adjustments workbook. The "
            "table starts at column G on sheet 'ra_AAI_REINS' (default). "
            "Year column names are chosen from semester: H1 -> 'HY_{year}', "
            "H2 -> 'FY_{year}'. Opening uses year-1, Closing uses year. "
            "Returns a dict keyed by GoC with 'opening' and 'closing' "
            "values. Read-only and idempotent."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "goc_names": {
                    "type": "array",
                    "items": {"type": "string"},
                },
                "year": {"type": "integer"},
                "semester": {"type": "integer", "enum": [1, 2]},
                "sheet": {"type": "string", "default": "ra_AAI_REINS"},
                "goc_column": {"type": "string", "default": "G"},
                "header_row": {"type": "integer", "default": 1},
            },
            "required": ["path", "goc_names", "year", "semester"],
        },
    },
    {
        "name": "create_risk_adjustment",
        "description": (
            "Create a 'Risk_Adjustment' workbook with two columns: "
            "ObservationID ('{goc}@Opening' or '{goc}@Closing') and "
            "Risk_Adjustment (values from `lookup_risk_adjustment_values`). "
            "Two rows per GoC. Missing lookups become empty cells. "
            "Overwrites the output file."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "goc_names": {
                    "type": "array",
                    "items": {"type": "string"},
                },
                "values": {
                    "type": "object",
                    "description": (
                        "Dict keyed by GoC with 'opening'/'closing' values, "
                        "as returned by lookup_risk_adjustment_values."
                    ),
                },
                "output_path": {"type": "string"},
            },
            "required": ["goc_names", "values", "output_path"],
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
    "extract_unique_goc_names": extract_unique_goc_names,
    "create_mp_lob": create_mp_lob,
    "create_mp_observation_year": create_mp_observation_year,
    "lookup_risk_adjustment_values": lookup_risk_adjustment_values,
    "create_risk_adjustment": create_risk_adjustment,
    # Add an entry for each domain-specific function above.
}


def dispatch_tool(name: str, arguments: dict[str, Any]) -> Any:
    """Route a tool call from Claude to the corresponding Python function."""
    if name not in _DISPATCH:
        raise ValueError(
            f"Unknown tool: {name}. Known: {sorted(_DISPATCH.keys())}"
        )
    return _DISPATCH[name](**arguments)
