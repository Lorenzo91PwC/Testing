"""Tests for excel_pipeline.skill."""
from __future__ import annotations

from pathlib import Path

import openpyxl

from excel_pipeline.skill import (
    create_mp_lob,
    create_mp_observation_year,
    extract_unique_goc_names,
)


def _build_ceded_fixture(path: Path) -> None:
    """Create a minimal workbook shaped like the AAI_P&C_Ceded input file.

    Matches the real file: rows 1-2 are header / sub-header, data from
    row 3 onwards.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AAI_P&C_Ceded_H_NH"
    ws.cell(row=1, column=27, value="GoC")  # column AA = 27
    ws.cell(row=2, column=27, value="Line of Business")  # sub-header
    values = [
        "Motor",
        "Property",
        "Motor",
        None,
        "  Property  ",
        "Liability",
        "",
        "Motor",
    ]
    for i, v in enumerate(values, start=3):
        ws.cell(row=i, column=27, value=v)
    wb.save(path)


def test_extract_unique_goc_names(tmp_path: Path) -> None:
    fixture = tmp_path / "1.1_2025.12.31_AAI_P&C_Ceded.xlsx"
    _build_ceded_fixture(fixture)

    result = extract_unique_goc_names(str(fixture))

    assert result["sheet"] == "AAI_P&C_Ceded_H_NH"
    assert result["column"] == "AA"
    assert result["values"] == ["Motor", "Property", "Liability"]
    assert result["count"] == 3


def test_extract_unique_goc_names_custom_column_and_start_row(tmp_path: Path) -> None:
    fixture = tmp_path / "custom.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AAI_P&C_Ceded_H_NH"
    ws.cell(row=1, column=2, value="GoC")
    for i, v in enumerate(["A", "B", "A"], start=2):
        ws.cell(row=i, column=2, value=v)
    wb.save(fixture)

    result = extract_unique_goc_names(str(fixture), column="B", start_row=2)

    assert result["values"] == ["A", "B"]


def test_create_mp_lob(tmp_path: Path) -> None:
    output = tmp_path / "MP_LoB.xlsx"
    goc_names = ["Motor", "Property", "Liability"]

    result = create_mp_lob(goc_names=goc_names, entity_id=6, output_path=str(output))

    assert result == {
        "output_path": str(output),
        "rows": 3,
        "columns": ["GoC_ID", "Entity_ID"],
    }
    assert output.exists()

    wb = openpyxl.load_workbook(output)
    ws = wb["MP_LoB"]
    rows = list(ws.iter_rows(values_only=True))
    assert rows == [
        ("GoC_ID", "Entity_ID"),
        ("Motor", 6),
        ("Property", 6),
        ("Liability", 6),
    ]


def test_create_mp_lob_end_to_end(tmp_path: Path) -> None:
    ceded = tmp_path / "1.1_2025.12.31_AAI_P&C_Ceded.xlsx"
    _build_ceded_fixture(ceded)
    output = tmp_path / "MP_LoB.xlsx"

    extracted = extract_unique_goc_names(str(ceded))
    create_mp_lob(
        goc_names=extracted["values"],
        entity_id=14,
        output_path=str(output),
    )

    wb = openpyxl.load_workbook(output)
    ws = wb["MP_LoB"]
    rows = list(ws.iter_rows(values_only=True))
    assert rows == [
        ("GoC_ID", "Entity_ID"),
        ("Motor", 14),
        ("Property", 14),
        ("Liability", 14),
    ]


def test_create_mp_observation_year(tmp_path: Path) -> None:
    output = tmp_path / "MP_ObservationYear.xlsx"
    goc_names = ["Motor", "Property"]

    result = create_mp_observation_year(
        goc_names=goc_names, year=2025, output_path=str(output),
    )

    assert result == {
        "output_path": str(output),
        "rows": 4,
        "columns": [
            "ObservationID",
            "ObservationYear",
            "LoB_ID",
            "AdjULAEPagate",
            "CY",
        ],
    }

    wb = openpyxl.load_workbook(output)
    ws = wb["MP_ObservationYear"]
    rows = list(ws.iter_rows(values_only=True))
    assert rows == [
        ("ObservationID", "ObservationYear", "LoB_ID", "AdjULAEPagate", "CY"),
        ("Motor@Opening", 2024, "Motor", 0, "Yes"),
        ("Motor@Closing", 2025, "Motor", 0, "Yes"),
        ("Property@Opening", 2024, "Property", 0, "Yes"),
        ("Property@Closing", 2025, "Property", 0, "Yes"),
    ]


def test_create_mp_observation_year_empty(tmp_path: Path) -> None:
    output = tmp_path / "MP_ObservationYear.xlsx"

    result = create_mp_observation_year(
        goc_names=[], year=2025, output_path=str(output),
    )

    assert result["rows"] == 0
    wb = openpyxl.load_workbook(output)
    ws = wb["MP_ObservationYear"]
    rows = list(ws.iter_rows(values_only=True))
    assert rows == [
        ("ObservationID", "ObservationYear", "LoB_ID", "AdjULAEPagate", "CY"),
    ]
