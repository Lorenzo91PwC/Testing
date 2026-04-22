"""Tests for excel_pipeline.skill."""
from __future__ import annotations

from pathlib import Path

import openpyxl

from excel_pipeline.skill import create_mp_lob, extract_unique_goc_names


def _build_ceded_fixture(path: Path) -> None:
    """Create a minimal workbook shaped like the AAI_P&C_Ceded input file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AAI_P&C_Ceded_H_NH"
    ws.cell(row=1, column=27, value="GoC")  # column AA = 27
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
    for i, v in enumerate(values, start=2):
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


def test_extract_unique_goc_names_custom_column(tmp_path: Path) -> None:
    fixture = tmp_path / "custom.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AAI_P&C_Ceded_H_NH"
    ws.cell(row=1, column=2, value="GoC")
    for i, v in enumerate(["A", "B", "A"], start=2):
        ws.cell(row=i, column=2, value=v)
    wb.save(fixture)

    result = extract_unique_goc_names(str(fixture), column="B")

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
