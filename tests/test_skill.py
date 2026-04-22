"""Tests for excel_pipeline.skill."""
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from excel_pipeline.skill import (
    create_mp_lob,
    create_mp_observation_year,
    create_risk_adjustment,
    extract_unique_goc_names,
    lookup_risk_adjustment_values,
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


def _build_payment_patterns_fixture(path: Path) -> None:
    """Fixture shaped like Payment_Patterns_&_Risk_Adjustments."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ra_AAI_REINS"
    # Column G = 7; header row 1.
    ws.cell(row=1, column=7, value="GoC")
    headers = ["HY_2024", "FY_2024", "HY_2025", "FY_2025"]
    for i, h in enumerate(headers, start=8):  # H, I, J, K
        ws.cell(row=1, column=i, value=h)
    data = [
        ("Motor", 100, 110, 120, 130),
        ("Property", 200, 210, 220, 230),
        ("Liability", 300, 310, 320, 330),
    ]
    for r, (goc, hy24, fy24, hy25, fy25) in enumerate(data, start=2):
        ws.cell(row=r, column=7, value=goc)
        ws.cell(row=r, column=8, value=hy24)
        ws.cell(row=r, column=9, value=fy24)
        ws.cell(row=r, column=10, value=hy25)
        ws.cell(row=r, column=11, value=fy25)
    wb.save(path)


def test_lookup_risk_adjustment_values_h2(tmp_path: Path) -> None:
    fixture = tmp_path / "pp.xlsx"
    _build_payment_patterns_fixture(fixture)

    values = lookup_risk_adjustment_values(
        path=str(fixture),
        goc_names=["Motor", "Property", "Liability"],
        year=2025,
        semester=2,
    )

    # H2 2025: Closing = FY_2025, Opening = FY_2024
    assert values == {
        "Motor": {"opening": 110, "closing": 130},
        "Property": {"opening": 210, "closing": 230},
        "Liability": {"opening": 310, "closing": 330},
    }


def test_lookup_risk_adjustment_values_h1(tmp_path: Path) -> None:
    fixture = tmp_path / "pp.xlsx"
    _build_payment_patterns_fixture(fixture)

    values = lookup_risk_adjustment_values(
        path=str(fixture),
        goc_names=["Motor"],
        year=2025,
        semester=1,
    )

    # H1 2025: Closing = HY_2025, Opening = HY_2024
    assert values == {"Motor": {"opening": 100, "closing": 120}}


def test_lookup_risk_adjustment_values_missing_goc(tmp_path: Path) -> None:
    fixture = tmp_path / "pp.xlsx"
    _build_payment_patterns_fixture(fixture)

    values = lookup_risk_adjustment_values(
        path=str(fixture),
        goc_names=["Motor", "UnknownGoC"],
        year=2025,
        semester=2,
    )

    assert values["UnknownGoC"] == {"opening": None, "closing": None}


def test_lookup_risk_adjustment_values_missing_year_column(tmp_path: Path) -> None:
    fixture = tmp_path / "pp.xlsx"
    _build_payment_patterns_fixture(fixture)

    with pytest.raises(KeyError, match=r"FY_20(29|30).*ra_AAI_REINS"):
        lookup_risk_adjustment_values(
            path=str(fixture),
            goc_names=["Motor"],
            year=2030,
            semester=2,
        )


def test_create_risk_adjustment(tmp_path: Path) -> None:
    output = tmp_path / "Risk_Adjustment.xlsx"
    goc_names = ["Motor", "Property"]
    values = {
        "Motor": {"opening": 110, "closing": 130},
        "Property": {"opening": 210, "closing": 230},
    }

    result = create_risk_adjustment(
        goc_names=goc_names, values=values, output_path=str(output),
    )

    assert result == {
        "output_path": str(output),
        "rows": 4,
        "columns": ["ObservationID", "Risk_Adjustment"],
    }

    wb = openpyxl.load_workbook(output)
    rows = list(wb["Risk_Adjustment"].iter_rows(values_only=True))
    assert rows == [
        ("ObservationID", "Risk_Adjustment"),
        ("Motor@Opening", 110),
        ("Motor@Closing", 130),
        ("Property@Opening", 210),
        ("Property@Closing", 230),
    ]


def test_create_risk_adjustment_missing_values_become_empty(tmp_path: Path) -> None:
    output = tmp_path / "Risk_Adjustment.xlsx"
    values = {"Motor": {"opening": None, "closing": None}}

    create_risk_adjustment(
        goc_names=["Motor"], values=values, output_path=str(output),
    )

    wb = openpyxl.load_workbook(output)
    rows = list(wb["Risk_Adjustment"].iter_rows(values_only=True))
    assert rows == [
        ("ObservationID", "Risk_Adjustment"),
        ("Motor@Opening", None),
        ("Motor@Closing", None),
    ]
