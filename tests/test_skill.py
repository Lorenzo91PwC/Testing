"""Tests for excel_pipeline.skill."""
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from excel_pipeline.skill import (
    MANDATORY_ACTUALS_VARIABLE_NAMES,
    create_coverage_unit,
    create_empty_workbook,
    create_mandatory_actuals,
    create_mp_lob,
    create_mp_observation_year,
    create_new_business_ppos,
    create_payment_pattern,
    create_reinsurance,
    create_risk_adjustment,
    extract_unique_goc_cohort_pairs,
    extract_unique_goc_names,
    lookup_payment_pattern_values,
    lookup_risk_adjustment_values,
    update_projection_parameters_entity,
)


def _pair(goc: str, year: int) -> dict:
    """Helper for tests — mirrors extract_unique_goc_cohort_pairs output."""
    return {"goc_id": f"{goc}{year}", "goc": goc, "year": year}


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


def _build_pp_sheet_fixture(path: Path) -> None:
    """Fixture of the pp_AAI_REINS sheet for Payment Pattern lookups."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "pp_AAI_REINS"
    # Header row: A and B empty, C=GoC, D=Year, E..AA = '0'..'22'
    ws.cell(row=1, column=3, value="GoC")
    ws.cell(row=1, column=4, value="Year")
    for i in range(23):
        ws.cell(row=1, column=5 + i, value=str(i))

    # Data rows — (goc, year_label, seed) -> values[i] = seed + i
    data = [
        ("Motor", "FY2024", 1000),
        ("Motor", "FY2025", 2000),
        ("Motor", "HY2024", 3000),
        ("Motor", "HY2025", 4000),
        ("Property", "FY2024", 5000),
        ("Property", "FY2025", 6000),
    ]
    for r, (goc, year_label, seed) in enumerate(data, start=2):
        ws.cell(row=r, column=3, value=goc)
        ws.cell(row=r, column=4, value=year_label)
        for i in range(23):
            ws.cell(row=r, column=5 + i, value=seed + i)
    wb.save(path)


def test_lookup_payment_pattern_values_h2(tmp_path: Path) -> None:
    fixture = tmp_path / "pp.xlsx"
    _build_pp_sheet_fixture(fixture)

    rows = lookup_payment_pattern_values(
        path=str(fixture),
        goc_names=["Motor"],
        year=2025,
        semester=2,
    )

    # H2 -> FY prefix. Order: reference year, then year-1.
    assert rows == [
        {"goc": "Motor", "year": 2025, "values": [2000 + i for i in range(23)]},
        {"goc": "Motor", "year": 2024, "values": [1000 + i for i in range(23)]},
    ]


def test_lookup_payment_pattern_values_h1(tmp_path: Path) -> None:
    fixture = tmp_path / "pp.xlsx"
    _build_pp_sheet_fixture(fixture)

    rows = lookup_payment_pattern_values(
        path=str(fixture),
        goc_names=["Motor"],
        year=2025,
        semester=1,
    )

    # H1 -> HY prefix
    assert rows[0]["values"][0] == 4000  # HY2025
    assert rows[1]["values"][0] == 3000  # HY2024


def test_lookup_payment_pattern_values_missing_row(tmp_path: Path) -> None:
    fixture = tmp_path / "pp.xlsx"
    _build_pp_sheet_fixture(fixture)

    rows = lookup_payment_pattern_values(
        path=str(fixture),
        goc_names=["Property"],
        year=2025,
        semester=1,  # HY2025 / HY2024 — not in fixture for Property
    )

    assert rows[0]["values"] == [None] * 23
    assert rows[1]["values"] == [None] * 23


def test_lookup_payment_pattern_values_fewer_than_23_columns(tmp_path: Path) -> None:
    fixture = tmp_path / "pp.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "pp_AAI_REINS"
    ws.cell(row=1, column=3, value="GoC")
    ws.cell(row=1, column=4, value="Year")
    # Only 5 data columns
    for i in range(5):
        ws.cell(row=1, column=5 + i, value=str(i))
    wb.save(fixture)

    with pytest.raises(KeyError, match="23 data columns"):
        lookup_payment_pattern_values(
            path=str(fixture),
            goc_names=["Motor"],
            year=2025,
            semester=2,
        )


def test_create_payment_pattern(tmp_path: Path) -> None:
    output = tmp_path / "Payment_pattern.xlsx"
    rows = [
        {"goc": "Motor", "year": 2025, "values": [i for i in range(23)]},
        {"goc": "Motor", "year": 2024, "values": [i * 10 for i in range(23)]},
    ]

    result = create_payment_pattern(rows=rows, output_path=str(output))

    expected_headers = ["GoC", "Year"] + [str(i) for i in range(23)]
    assert result == {
        "output_path": str(output),
        "rows": 2,
        "columns": expected_headers,
    }

    wb = openpyxl.load_workbook(output)
    ws = wb["Payment_pattern"]
    written = list(ws.iter_rows(values_only=True))
    assert written[0] == tuple(expected_headers)
    assert written[1] == ("Motor", 2025) + tuple(range(23))
    assert written[2] == ("Motor", 2024) + tuple(i * 10 for i in range(23))


def test_create_empty_workbook(tmp_path: Path) -> None:
    output = tmp_path / "Empty.xlsx"

    result = create_empty_workbook(str(output), sheet_name="Astra_Placeholder")

    assert result == {
        "output_path": str(output),
        "rows": 0,
        "columns": [],
    }
    assert output.exists()

    wb = openpyxl.load_workbook(output)
    assert wb.sheetnames == ["Astra_Placeholder"]
    ws = wb["Astra_Placeholder"]
    assert list(ws.iter_rows(values_only=True)) == []


def _build_ceded_with_year_fixture(path: Path, rows_data: list[tuple]) -> None:
    """Ceded fixture with GoC name in column AA and cohort year in column AB.

    `rows_data` is a list of (goc, year) tuples written from row 3 onward.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AAI_P&C_Ceded_H_NH"
    ws.cell(row=1, column=27, value="GoC")
    ws.cell(row=1, column=28, value="Cohort Year")
    ws.cell(row=2, column=27, value="Code")
    ws.cell(row=2, column=28, value="Year")
    for i, (goc, year) in enumerate(rows_data, start=3):
        ws.cell(row=i, column=27, value=goc)
        ws.cell(row=i, column=28, value=year)
    wb.save(path)


def test_extract_unique_goc_cohort_pairs(tmp_path: Path) -> None:
    fixture = tmp_path / "ceded.xlsx"
    _build_ceded_with_year_fixture(
        fixture,
        [
            ("IT05PABPPLE", 2024),
            ("IT05PABPPLE", 2023),
            ("IT05PABPPLE", 2024),  # duplicate -> dropped
            ("  IT05PABPPLE  ", 2022),  # stripped
            ("IT06ABCDE", 2024),
            (None, 2020),  # skipped (None GoC)
            ("IT06ABCDE", None),  # skipped (None year)
            ("", 2020),  # skipped (empty GoC)
        ],
    )

    result = extract_unique_goc_cohort_pairs(str(fixture))

    assert result["sheet"] == "AAI_P&C_Ceded_H_NH"
    assert result["count"] == 4
    assert result["pairs"] == [
        {"goc_id": "IT05PABPPLE2024", "goc": "IT05PABPPLE", "year": 2024},
        {"goc_id": "IT05PABPPLE2023", "goc": "IT05PABPPLE", "year": 2023},
        {"goc_id": "IT05PABPPLE2022", "goc": "IT05PABPPLE", "year": 2022},
        {"goc_id": "IT06ABCDE2024", "goc": "IT06ABCDE", "year": 2024},
    ]


def test_create_new_business_ppos(tmp_path: Path) -> None:
    output = tmp_path / "NEW_BUSINESS_PPOS.xlsx"
    pairs = [_pair("IT05PABPPLE", 2024), _pair("IT05PABPPLE", 2023), _pair("IT06ABCDE", 2024)]

    result = create_new_business_ppos(pairs=pairs, output_path=str(output))

    assert result == {
        "output_path": str(output),
        "rows": 3,
        "columns": ["GOC_ID", "VARIABLE_NAME", "1"],
    }

    wb = openpyxl.load_workbook(output)
    rows = list(wb["NEW_BUSINESS_PPOS"].iter_rows(values_only=True))
    assert rows == [
        ("GOC_ID", "VARIABLE_NAME", 1),
        ("IT05PABPPLE2024", "CROSS_SUB_FASSCHNG", 0),
        ("IT05PABPPLE2023", "CROSS_SUB_FASSCHNG", 0),
        ("IT06ABCDE2024", "CROSS_SUB_FASSCHNG", 0),
    ]


def test_create_new_business_ppos_empty_pairs(tmp_path: Path) -> None:
    output = tmp_path / "NEW_BUSINESS_PPOS.xlsx"

    result = create_new_business_ppos(pairs=[], output_path=str(output))

    assert result["rows"] == 0
    wb = openpyxl.load_workbook(output)
    rows = list(wb["NEW_BUSINESS_PPOS"].iter_rows(values_only=True))
    assert rows == [("GOC_ID", "VARIABLE_NAME", 1)]


def test_create_coverage_unit(tmp_path: Path) -> None:
    output = tmp_path / "COVERAGE_UNIT.xlsx"
    pairs = [_pair("IT05PABPPLE", 2024), _pair("IT06ABCDE", 2024)]

    result = create_coverage_unit(pairs=pairs, output_path=str(output))

    expected_columns = ["GOC_ID", "PROJECTION_PERIOD"] + [str(i) for i in range(1, 101)]
    assert result == {
        "output_path": str(output),
        "rows": 2,
        "columns": expected_columns,
    }

    wb = openpyxl.load_workbook(output)
    rows = list(wb["COVERAGE_UNIT"].iter_rows(values_only=True))
    assert rows[0] == ("GOC_ID", "PROJECTION_PERIOD") + tuple(range(1, 101))
    assert rows[1] == ("IT05PABPPLE2024", 1) + (0,) * 100
    assert rows[2] == ("IT06ABCDE2024", 1) + (0,) * 100
    assert len(rows) == 3


def test_create_reinsurance(tmp_path: Path) -> None:
    output = tmp_path / "REINSURANCE.xlsx"
    pairs = [
        _pair("IT05PABPPLE", 2024),
        _pair("IT05PABPPLE", 2023),
        _pair("IT06ABCDE", 2024),
    ]

    result = create_reinsurance(pairs=pairs, output_path=str(output))

    assert result == {
        "output_path": str(output),
        "rows": 3 * 2,
        "columns": ["GOC_ID", "VARIABLE_NAME", "1", "T"],
    }

    wb = openpyxl.load_workbook(output)
    rows = list(wb["REINSURANCE"].iter_rows(values_only=True))
    assert rows == [
        ("GOC_ID", "VARIABLE_NAME", 1, "T"),
        # First GoC: 2 IFE rows (one per pair of that GoC), then 2 CLOSING rows
        ("IT05PABPPLE2024", "LOSSRECO_IFE_ALLOCATION", 0, 2024),
        ("IT05PABPPLE2023", "LOSSRECO_IFE_ALLOCATION", 0, 2023),
        ("IT05PABPPLE2024", "LOSSRECO_CLOSING", 0, 2024),
        ("IT05PABPPLE2023", "LOSSRECO_CLOSING", 0, 2023),
        # Second GoC: 1 IFE then 1 CLOSING
        ("IT06ABCDE2024", "LOSSRECO_IFE_ALLOCATION", 0, 2024),
        ("IT06ABCDE2024", "LOSSRECO_CLOSING", 0, 2024),
    ]


def test_create_reinsurance_empty_pairs(tmp_path: Path) -> None:
    output = tmp_path / "REINSURANCE.xlsx"

    result = create_reinsurance(pairs=[], output_path=str(output))

    assert result["rows"] == 0
    wb = openpyxl.load_workbook(output)
    rows = list(wb["REINSURANCE"].iter_rows(values_only=True))
    assert rows == [("GOC_ID", "VARIABLE_NAME", 1, "T")]


def test_create_mandatory_actuals(tmp_path: Path) -> None:
    output = tmp_path / "MANDATORY_ACTUALS.xlsx"
    assert len(MANDATORY_ACTUALS_VARIABLE_NAMES) == 16  # guard the spec
    pairs = [
        _pair("IT05PABPPLE", 2024),
        _pair("IT05PABPPLE", 2023),
        _pair("IT06ABCDE", 2024),
    ]

    result = create_mandatory_actuals(pairs=pairs, output_path=str(output))

    assert result == {
        "output_path": str(output),
        "rows": 3 * 16,
        "columns": ["GOC_ID", "VARIABLE_NAME", "1"],
    }

    wb = openpyxl.load_workbook(output)
    rows = list(wb["MANDATORY_ACTUALS"].iter_rows(values_only=True))
    assert rows[0] == ("GOC_ID", "VARIABLE_NAME", 1)
    assert len(rows) == 1 + 3 * 16

    # First pair (IT05PABPPLE2024): 16 variables in order
    expected_first = [
        ("IT05PABPPLE2024", v, 0) for v in MANDATORY_ACTUALS_VARIABLE_NAMES
    ]
    assert rows[1:17] == expected_first

    # Second pair (IT05PABPPLE2023): 16 variables in order
    expected_second = [
        ("IT05PABPPLE2023", v, 0) for v in MANDATORY_ACTUALS_VARIABLE_NAMES
    ]
    assert rows[17:33] == expected_second

    # Third pair (IT06ABCDE2024): 16 variables in order
    expected_third = [
        ("IT06ABCDE2024", v, 0) for v in MANDATORY_ACTUALS_VARIABLE_NAMES
    ]
    assert rows[33:49] == expected_third


def test_create_mandatory_actuals_empty_pairs(tmp_path: Path) -> None:
    output = tmp_path / "MANDATORY_ACTUALS.xlsx"

    result = create_mandatory_actuals(pairs=[], output_path=str(output))

    assert result["rows"] == 0
    wb = openpyxl.load_workbook(output)
    rows = list(wb["MANDATORY_ACTUALS"].iter_rows(values_only=True))
    assert rows == [("GOC_ID", "VARIABLE_NAME", 1)]


PROJECTION_PARAMS_INPUT_ROWS = [
    ("PROJECTED_PERIODS", 110),
    ("SCENARIO_TIMESTEP", "YEARLY"),
    ("CF_TIMESTEP", "SEMESTRIAL"),
    ("REPORTING_MONTH", "12_DECEMBER"),
    ("PROJ_PERIOD_TIMESTEP", "YEARLY"),
    ("COVERAGE_UNIT_FORMAT", "COVERAGE_UNIT"),
    ("OCI_OPTION_APPROACH", "INDIRECT"),
    ("OCI_OPTION_RA_PROXY", "1_PROP_PVFC"),
    ("SCENARIO_MANAGEMENT", "CENTRAL"),
    ("FX_OPENING_DATE", "1M25"),
    ("FX_OPENING_RATE_TYPE", "YTD_VALUE"),
    ("FX_AVERAGE_DATE", "HY25"),
    ("FX_AVERAGE_RATE_TYPE", "YTD_VALUE"),
    ("FX_CLOSING_DATE", "FY25"),
    ("FX_CLOSING_RATE_TYPE", "YTD_VALUE"),
    ("FX_REPORTING_DATE", 20251231),
    ("FX_MANAGEMENT", "CENTRAL"),
    ("ACTUARIAL_AOM_ACQ_CF", "SEPARATE_INPUT"),
    ("OCI_OPTION_CF", "USE_EXISTING_INPUT"),
]


def _build_projection_parameters_fixture(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="PARAMETER")
    ws.cell(row=1, column=2, value="VALUE")
    for i, (param, val) in enumerate(PROJECTION_PARAMS_INPUT_ROWS, start=2):
        ws.cell(row=i, column=1, value=param)
        ws.cell(row=i, column=2, value=val)
    wb.save(path)


def test_update_projection_parameters_entity_h2_2025(tmp_path: Path) -> None:
    fixture = tmp_path / "PROJECTION_PARAMETERS_ENTITY.xlsx"
    output = tmp_path / "out.xlsx"
    _build_projection_parameters_fixture(fixture)

    result = update_projection_parameters_entity(
        input_path=str(fixture),
        output_path=str(output),
        year=2025,
        semester=2,
    )

    assert result["rows_updated"] == 6
    assert set(result["parameters_updated"]) == {
        "CF_TIMESTEP",
        "REPORTING_MONTH",
        "FX_OPENING_DATE",
        "FX_AVERAGE_DATE",
        "FX_CLOSING_DATE",
        "FX_REPORTING_DATE",
    }

    wb = openpyxl.load_workbook(output)
    ws = wb.active
    by_param = {
        ws.cell(row=r, column=1).value: ws.cell(row=r, column=2).value
        for r in range(2, ws.max_row + 1)
    }
    # Edits applied
    assert by_param["CF_TIMESTEP"] == "YEARLY"
    assert by_param["REPORTING_MONTH"] == "12_DECEMBER"
    assert by_param["FX_OPENING_DATE"] == "1M25"
    assert by_param["FX_AVERAGE_DATE"] == "HY25"
    assert by_param["FX_CLOSING_DATE"] == "FY25"
    assert by_param["FX_REPORTING_DATE"] == "20251231"
    # Unchanged values
    assert by_param["PROJECTED_PERIODS"] == 110
    assert by_param["SCENARIO_TIMESTEP"] == "YEARLY"
    assert by_param["FX_MANAGEMENT"] == "CENTRAL"
    assert by_param["OCI_OPTION_CF"] == "USE_EXISTING_INPUT"
    # Header preserved
    assert ws.cell(row=1, column=1).value == "PARAMETER"
    assert ws.cell(row=1, column=2).value == "VALUE"


def test_update_projection_parameters_entity_h1_2025(tmp_path: Path) -> None:
    fixture = tmp_path / "PROJECTION_PARAMETERS_ENTITY.xlsx"
    output = tmp_path / "out.xlsx"
    _build_projection_parameters_fixture(fixture)

    update_projection_parameters_entity(
        input_path=str(fixture),
        output_path=str(output),
        year=2025,
        semester=1,
    )

    wb = openpyxl.load_workbook(output)
    ws = wb.active
    by_param = {
        ws.cell(row=r, column=1).value: ws.cell(row=r, column=2).value
        for r in range(2, ws.max_row + 1)
    }
    assert by_param["REPORTING_MONTH"] == "6_JUNE"
    assert by_param["FX_AVERAGE_DATE"] == "Q125"
    assert by_param["FX_CLOSING_DATE"] == "HY25"
    assert by_param["FX_REPORTING_DATE"] == "20250630"
    assert by_param["FX_OPENING_DATE"] == "1M25"
    assert by_param["CF_TIMESTEP"] == "YEARLY"


def test_update_projection_parameters_entity_drops_extra_columns(tmp_path: Path) -> None:
    """Even if the input has a stray third column, the output is 2-column."""
    fixture = tmp_path / "input.xlsx"
    output = tmp_path / "out.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="PARAMETER")
    ws.cell(row=1, column=2, value="VALUE")
    ws.cell(row=1, column=3, value="Notes")  # stray column to be dropped
    ws.cell(row=2, column=1, value="CF_TIMESTEP")
    ws.cell(row=2, column=2, value="SEMESTRIAL")
    ws.cell(row=2, column=3, value="should disappear")
    wb.save(fixture)

    update_projection_parameters_entity(
        input_path=str(fixture),
        output_path=str(output),
        year=2025,
        semester=2,
    )

    out_ws = openpyxl.load_workbook(output).active
    assert out_ws.max_column == 2
    assert out_ws.cell(row=1, column=1).value == "PARAMETER"
    assert out_ws.cell(row=1, column=2).value == "VALUE"
    assert out_ws.cell(row=2, column=2).value == "YEARLY"


def test_update_projection_parameters_entity_invalid_semester(tmp_path: Path) -> None:
    fixture = tmp_path / "PROJECTION_PARAMETERS_ENTITY.xlsx"
    _build_projection_parameters_fixture(fixture)

    with pytest.raises(ValueError, match="semester must be 1 or 2"):
        update_projection_parameters_entity(
            input_path=str(fixture),
            output_path=str(tmp_path / "out.xlsx"),
            year=2025,
            semester=3,
        )
