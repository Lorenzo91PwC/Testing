"""Tests for excel_pipeline.pipeline (deterministic, no API calls)."""
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from excel_pipeline.pipeline import run_astra_phase1, run_phase1


def _build_ceded_fixture(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AAI_P&C_Ceded_H_NH"
    ws.cell(row=1, column=27, value="GoC")
    ws.cell(row=2, column=27, value="Line of Business")
    for i, v in enumerate(
        ["Motor", "Property", "Motor", "Liability"], start=3,
    ):
        ws.cell(row=i, column=27, value=v)
    wb.save(path)


def _build_payment_patterns_fixture(path: Path) -> None:
    """Fixture with both the ra_AAI_REINS and pp_AAI_REINS sheets."""
    wb = openpyxl.Workbook()

    # --- ra_AAI_REINS ---
    ra = wb.active
    ra.title = "ra_AAI_REINS"
    ra.cell(row=1, column=7, value="GoC")
    for i, h in enumerate(["HY_2024", "FY_2024", "HY_2025", "FY_2025"], start=8):
        ra.cell(row=1, column=i, value=h)
    ra_data = [
        ("Motor", 100, 110, 120, 130),
        ("Property", 200, 210, 220, 230),
        ("Liability", 300, 310, 320, 330),
    ]
    for r, (goc, *vals) in enumerate(ra_data, start=2):
        ra.cell(row=r, column=7, value=goc)
        for j, v in enumerate(vals):
            ra.cell(row=r, column=8 + j, value=v)

    # --- pp_AAI_REINS ---
    pp = wb.create_sheet("pp_AAI_REINS")
    pp.cell(row=1, column=3, value="GoC")
    pp.cell(row=1, column=4, value="Year")
    for i in range(23):
        pp.cell(row=1, column=5 + i, value=str(i))
    pp_data = [
        ("Motor", "FY2024", 1000),
        ("Motor", "FY2025", 2000),
        ("Property", "FY2024", 5000),
        ("Property", "FY2025", 6000),
        ("Liability", "FY2024", 7000),
        ("Liability", "FY2025", 8000),
    ]
    for r, (goc, yr, seed) in enumerate(pp_data, start=2):
        pp.cell(row=r, column=3, value=goc)
        pp.cell(row=r, column=4, value=yr)
        for i in range(23):
            pp.cell(row=r, column=5 + i, value=seed + i)

    wb.save(path)


def test_run_phase1_happy_path(tmp_path: Path) -> None:
    inputs_dir = tmp_path / "inputs"
    inputs_dir.mkdir()
    ceded = inputs_dir / "1.1_2025.12.31_AAI_P&C_Ceded.xlsx"
    _build_ceded_fixture(ceded)
    payments = inputs_dir / "1.2_2025.12.31_Payment_Patterns_&_Risk_Adjustments.xlsx"
    _build_payment_patterns_fixture(payments)

    result = run_phase1(
        input_paths=[ceded, payments],
        run_dir=tmp_path,
        entity_id=6,
        entity_name="AAI",
        year=2025,
        semester=2,
    )

    assert result["goc_names"] == ["Motor", "Property", "Liability"]
    assert result["year"] == 2025
    assert result["semester"] == 2

    outputs = result["outputs"]
    assert outputs == [
        tmp_path / "MP_LoB.xlsx",
        tmp_path / "MP_ObservationYear.xlsx",
        tmp_path / "Risk_Adjustment.xlsx",
        tmp_path / "Payment_pattern.xlsx",
    ]
    for p in outputs:
        assert p.exists()

    mp_lob = list(openpyxl.load_workbook(outputs[0])["MP_LoB"].iter_rows(values_only=True))
    assert mp_lob == [
        ("GoC_ID", "Entity_ID"),
        ("Motor", 6),
        ("Property", 6),
        ("Liability", 6),
    ]

    mp_obs = list(
        openpyxl.load_workbook(outputs[1])["MP_ObservationYear"].iter_rows(values_only=True)
    )
    assert mp_obs == [
        ("ObservationID", "ObservationYear", "LoB_ID", "AdjULAEPagate", "CY"),
        ("Motor@Opening", 2024, "Motor", 0, "Yes"),
        ("Motor@Closing", 2025, "Motor", 0, "Yes"),
        ("Property@Opening", 2024, "Property", 0, "Yes"),
        ("Property@Closing", 2025, "Property", 0, "Yes"),
        ("Liability@Opening", 2024, "Liability", 0, "Yes"),
        ("Liability@Closing", 2025, "Liability", 0, "Yes"),
    ]

    # H2 2025: Closing = FY_2025, Opening = FY_2024
    ra = list(
        openpyxl.load_workbook(outputs[2])["Risk_Adjustment"].iter_rows(values_only=True)
    )
    assert ra == [
        ("ObservationID", "Risk_Adjustment"),
        ("Motor@Opening", 110),
        ("Motor@Closing", 130),
        ("Property@Opening", 210),
        ("Property@Closing", 230),
        ("Liability@Opening", 310),
        ("Liability@Closing", 330),
    ]

    pp = list(
        openpyxl.load_workbook(outputs[3])["Payment_pattern"].iter_rows(values_only=True)
    )
    expected_headers = ("GoC", "Year") + tuple(str(i) for i in range(23))
    assert pp[0] == expected_headers
    # For each GoC, reference year first then year-1.
    # seed: Motor FY2025=2000, FY2024=1000; Property FY2025=6000, FY2024=5000; Liability FY2025=8000, FY2024=7000.
    assert pp[1] == ("Motor", 2025) + tuple(2000 + i for i in range(23))
    assert pp[2] == ("Motor", 2024) + tuple(1000 + i for i in range(23))
    assert pp[3] == ("Property", 2025) + tuple(6000 + i for i in range(23))
    assert pp[4] == ("Property", 2024) + tuple(5000 + i for i in range(23))
    assert pp[5] == ("Liability", 2025) + tuple(8000 + i for i in range(23))
    assert pp[6] == ("Liability", 2024) + tuple(7000 + i for i in range(23))


def test_run_phase1_missing_ceded(tmp_path: Path) -> None:
    other = tmp_path / "something_else.xlsx"
    wb = openpyxl.Workbook()
    wb.save(other)

    with pytest.raises(FileNotFoundError, match="AAI_P&C_Ceded"):
        run_phase1(
            input_paths=[other],
            run_dir=tmp_path,
            entity_id=6,
            entity_name="AAI",
            year=2025,
            semester=2,
        )


def test_run_phase1_picks_matching_files(tmp_path: Path) -> None:
    """With extra unrelated files mixed in, the right inputs are picked."""
    inputs_dir = tmp_path / "inputs"
    inputs_dir.mkdir()
    ceded = inputs_dir / "1.1_2025.12.31_AAI_P&C_Ceded.xlsx"
    _build_ceded_fixture(ceded)
    payments = inputs_dir / "1.2_2025.12.31_Payment_Patterns_&_Risk_Adjustments.xlsx"
    _build_payment_patterns_fixture(payments)
    unrelated = inputs_dir / "notes.xlsx"
    openpyxl.Workbook().save(unrelated)

    result = run_phase1(
        input_paths=[unrelated, ceded, payments],
        run_dir=tmp_path,
        entity_id=14,
        entity_name="MPS",
        year=2025,
        semester=1,
    )
    outputs = result["outputs"]

    mp_lob_rows = list(
        openpyxl.load_workbook(outputs[0])["MP_LoB"].iter_rows(values_only=True)
    )
    assert mp_lob_rows[1:] == [("Motor", 14), ("Property", 14), ("Liability", 14)]

    # H1 2025: Closing = HY_2025, Opening = HY_2024
    ra_rows = list(
        openpyxl.load_workbook(outputs[2])["Risk_Adjustment"].iter_rows(values_only=True)
    )
    assert ra_rows[1:] == [
        ("Motor@Opening", 100),
        ("Motor@Closing", 120),
        ("Property@Opening", 200),
        ("Property@Closing", 220),
        ("Liability@Opening", 300),
        ("Liability@Closing", 320),
    ]


def test_run_astra_phase1_produces_six_workbooks(tmp_path: Path) -> None:
    outputs = run_astra_phase1(
        input_paths=[],
        run_dir=tmp_path,
        goc_names=["IT05PABPPLE", "IT06ABCDE"],
        year=2024,
    )

    assert outputs == [
        tmp_path / "NEW_BUSINESS_PPOS.xlsx",
        tmp_path / "COVERAGE_UNIT.xlsx",
        tmp_path / "REINSURANCE.xlsx",
        tmp_path / "MANDATORY_ACTUALS.xlsx",
        tmp_path / "OCI_OPTION_CF_CLOSING.xlsx",
        tmp_path / "OCI_OPTION_CF_OPENING.xlsx",
    ]
    for p in outputs:
        assert p.exists()

    # NEW_BUSINESS_PPOS: 16 rows per GoC + 1 header row
    nb_rows = list(
        openpyxl.load_workbook(outputs[0])["NEW_BUSINESS_PPOS"].iter_rows(values_only=True)
    )
    assert nb_rows[0] == ("GOC_ID", "VARIABLE_NAME", 1)
    assert len(nb_rows) == 1 + 2 * 16
    assert nb_rows[1] == ("IT05PABPPLE2024", "CROSS_SUB_FASSCHNG", 0)
    assert nb_rows[16] == ("IT05PABPPLE2009", "CROSS_SUB_FASSCHNG", 0)
    assert nb_rows[17] == ("IT06ABCDE2024", "CROSS_SUB_FASSCHNG", 0)

    # COVERAGE_UNIT: 102 columns, same row layout
    cu_rows = list(
        openpyxl.load_workbook(outputs[1])["COVERAGE_UNIT"].iter_rows(values_only=True)
    )
    assert cu_rows[0] == ("GOC_ID", "PROJECTION_PERIOD") + tuple(range(1, 101))
    assert len(cu_rows) == 1 + 2 * 16
    assert cu_rows[1] == ("IT05PABPPLE2024", 1) + (0,) * 100
    assert cu_rows[17] == ("IT06ABCDE2024", 1) + (0,) * 100

    # REINSURANCE: 32 rows per GoC + header. Per GoC: 16 IFE then 16 CLOSING.
    rein_rows = list(
        openpyxl.load_workbook(outputs[2])["REINSURANCE"].iter_rows(values_only=True)
    )
    assert rein_rows[0] == ("GOC_ID", "VARIABLE_NAME", 1, "T")
    assert len(rein_rows) == 1 + 2 * 32
    assert rein_rows[1] == ("IT05PABPPLE2024", "LOSSRECO_IFE_ALLOCATION", 0, 2024)
    assert rein_rows[17] == ("IT05PABPPLE2024", "LOSSRECO_CLOSING", 0, 2024)
    assert rein_rows[33] == ("IT06ABCDE2024", "LOSSRECO_IFE_ALLOCATION", 0, 2024)

    # MANDATORY_ACTUALS: 256 rows per GoC + header.
    ma_rows = list(
        openpyxl.load_workbook(outputs[3])["MANDATORY_ACTUALS"].iter_rows(values_only=True)
    )
    assert ma_rows[0] == ("GOC_ID", "VARIABLE_NAME", 1)
    assert len(ma_rows) == 1 + 2 * 16 * 16
    assert ma_rows[1] == ("IT05PABPPLE2024", "ACTUAL_PREMIUM_CF_PAST_SERVICE", 0)
    # Last variable of first cohort year is at row index 16 (1 header + 16 vars)
    assert ma_rows[16] == ("IT05PABPPLE2024", "THEORETICAL_PREMIUM_DERECOGNITION_LIC", 0)
    # Second cohort year starts at row index 17
    assert ma_rows[17] == ("IT05PABPPLE2023", "ACTUAL_PREMIUM_CF_PAST_SERVICE", 0)

    # OCI files are still empty placeholders
    for p in outputs[4:]:
        wb = openpyxl.load_workbook(p)
        assert len(wb.sheetnames) == 1
        assert list(wb[wb.sheetnames[0]].iter_rows(values_only=True)) == []
