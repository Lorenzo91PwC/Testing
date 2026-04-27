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


def _build_ceded_with_pairs_fixture(
    path: Path, rows_data: list[tuple[str | None, int | None]]
) -> None:
    """Ceded fixture with GoC names in column AA and cohort years in column AB."""
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


def _build_projection_parameters_fixture(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="PARAMETER")
    ws.cell(row=1, column=2, value="VALUE")
    rows = [
        ("PROJECTED_PERIODS", 110),
        ("CF_TIMESTEP", "SEMESTRIAL"),
        ("REPORTING_MONTH", "12_DECEMBER"),
        ("FX_OPENING_DATE", "1M25"),
        ("FX_AVERAGE_DATE", "HY25"),
        ("FX_CLOSING_DATE", "FY25"),
        ("FX_REPORTING_DATE", 20251231),
    ]
    for i, (param, val) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=param)
        ws.cell(row=i, column=2, value=val)
    wb.save(path)


def _build_mp_goc_seg_fixture(path: Path, rows: list[tuple]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="GOC_SEG_ID")
    ws.cell(row=1, column=2, value="GOC_ID")
    ws.cell(row=1, column=3, value="SEG_ID")
    ws.cell(row=1, column=4, value="ALLOCATION_RATIO")
    for i, (a, b, c, d) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=a)
        ws.cell(row=i, column=2, value=b)
        ws.cell(row=i, column=3, value=c)
        ws.cell(row=i, column=4, value=d)
    wb.save(path)


def test_run_astra_phase1_uses_pairs_from_ceded(tmp_path: Path) -> None:
    inputs_dir = tmp_path / "inputs"
    inputs_dir.mkdir()
    ceded = inputs_dir / "1.1_2024.12.31_AAI_P&C_Ceded.xlsx"
    _build_ceded_with_pairs_fixture(
        ceded,
        [
            ("IT05PABPPLE", 2024),
            ("IT05PABPPLE", 2023),
            ("IT06ABCDE", 2024),
        ],
    )
    pp_params = inputs_dir / "1.4_2024.12.31_PROJECTION_PARAMETERS_ENTITY.xlsx"
    _build_projection_parameters_fixture(pp_params)
    mp_goc_seg = inputs_dir / "1.5_2024.12.31_MP_GOC_SEG.xlsx"
    _build_mp_goc_seg_fixture(
        mp_goc_seg,
        [
            ("IT05PABPPLE2024_02_P&C", "IT05PABPPLE2024", "02_P&C", 1),
            ("IT05RRIEEBB2024_02_P&C", "IT05RRIEEBB2024", "02_P&C", 1),
        ],
    )

    outputs = run_astra_phase1(
        input_paths=[ceded, pp_params, mp_goc_seg],
        run_dir=tmp_path,
        entity_id=6,
        entity_name="AAI",
        year=2024,
        semester=2,
        health_perimeter_gocs=["IT05RRIEEBB"],
    )

    assert outputs == [
        tmp_path / "NEW_BUSINESS_PPOS.xlsx",
        tmp_path / "COVERAGE_UNIT.xlsx",
        tmp_path / "REINSURANCE.xlsx",
        tmp_path / "MANDATORY_ACTUALS.xlsx",
        tmp_path / "PROJECTION_PARAMETERS_ENTITY.xlsx",
        tmp_path / "MP_GOC_SEG.xlsx",
        tmp_path / "OCI_OPTION_CF_CLOSING.xlsx",
        tmp_path / "OCI_OPTION_CF_OPENING.xlsx",
    ]
    for p in outputs:
        assert p.exists()

    # NEW_BUSINESS_PPOS: one row per pair (3 pairs in fixture)
    nb_rows = list(
        openpyxl.load_workbook(outputs[0])["NEW_BUSINESS_PPOS"].iter_rows(values_only=True)
    )
    assert nb_rows == [
        ("GOC_ID", "VARIABLE_NAME", 1),
        ("IT05PABPPLE2024", "CROSS_SUB_FASSCHNG", 0),
        ("IT05PABPPLE2023", "CROSS_SUB_FASSCHNG", 0),
        ("IT06ABCDE2024", "CROSS_SUB_FASSCHNG", 0),
    ]

    # COVERAGE_UNIT: 1 header + 3 pair rows
    cu_rows = list(
        openpyxl.load_workbook(outputs[1])["COVERAGE_UNIT"].iter_rows(values_only=True)
    )
    assert len(cu_rows) == 1 + 3
    assert cu_rows[1] == ("IT05PABPPLE2024", 1) + (0,) * 100
    assert cu_rows[3] == ("IT06ABCDE2024", 1) + (0,) * 100

    # REINSURANCE: per GoC, all IFE rows then all CLOSING rows
    rein_rows = list(
        openpyxl.load_workbook(outputs[2])["REINSURANCE"].iter_rows(values_only=True)
    )
    assert rein_rows == [
        ("GOC_ID", "VARIABLE_NAME", 1, "T"),
        ("IT05PABPPLE2024", "LOSSRECO_IFE_ALLOCATION", 0, 2024),
        ("IT05PABPPLE2023", "LOSSRECO_IFE_ALLOCATION", 0, 2023),
        ("IT05PABPPLE2024", "LOSSRECO_CLOSING", 0, 2024),
        ("IT05PABPPLE2023", "LOSSRECO_CLOSING", 0, 2023),
        ("IT06ABCDE2024", "LOSSRECO_IFE_ALLOCATION", 0, 2024),
        ("IT06ABCDE2024", "LOSSRECO_CLOSING", 0, 2024),
    ]

    # MANDATORY_ACTUALS: 1 header + (3 pairs * 16 variables) = 49 rows
    ma_rows = list(
        openpyxl.load_workbook(outputs[3])["MANDATORY_ACTUALS"].iter_rows(values_only=True)
    )
    assert len(ma_rows) == 1 + 3 * 16
    assert ma_rows[1] == ("IT05PABPPLE2024", "ACTUAL_PREMIUM_CF_PAST_SERVICE", 0)
    # Second pair starts at row 17
    assert ma_rows[17] == ("IT05PABPPLE2023", "ACTUAL_PREMIUM_CF_PAST_SERVICE", 0)
    assert ma_rows[33] == ("IT06ABCDE2024", "ACTUAL_PREMIUM_CF_PAST_SERVICE", 0)

    # PROJECTION_PARAMETERS_ENTITY: edits applied for H2 / 2024
    pp_ws = openpyxl.load_workbook(outputs[4]).active
    by_param = {
        pp_ws.cell(row=r, column=1).value: pp_ws.cell(row=r, column=2).value
        for r in range(2, pp_ws.max_row + 1)
    }
    assert by_param["CF_TIMESTEP"] == "YEARLY"
    assert by_param["REPORTING_MONTH"] == "12_DECEMBER"
    assert by_param["FX_OPENING_DATE"] == "1M24"
    assert by_param["FX_AVERAGE_DATE"] == "HY24"
    assert by_param["FX_CLOSING_DATE"] == "FY24"
    assert by_param["FX_REPORTING_DATE"] == "20241231"
    # Untouched
    assert by_param["PROJECTED_PERIODS"] == 110

    # MP_GOC_SEG: only the IT05RRIEEBB row gets P&C -> HLTH_PC
    seg_rows = list(openpyxl.load_workbook(outputs[5]).active.iter_rows(values_only=True))
    assert seg_rows[1] == ("IT05PABPPLE2024_02_P&C", "IT05PABPPLE2024", "02_P&C", 1)
    assert seg_rows[2] == (
        "IT05RRIEEBB2024_02_HLTH_PC", "IT05RRIEEBB2024", "02_HLTH_PC", 1,
    )


def test_run_astra_phase1_filters_pairs_outside_window(tmp_path: Path) -> None:
    """Pairs outside [year-15, year] are dropped before generation."""
    inputs_dir = tmp_path / "inputs"
    inputs_dir.mkdir()
    ceded = inputs_dir / "1.1_2024.12.31_AAI_P&C_Ceded.xlsx"
    _build_ceded_with_pairs_fixture(
        ceded,
        [
            ("IT05PABPPLE", 2024),  # kept (= year)
            ("IT05PABPPLE", 2009),  # kept (= year - 15)
            ("IT05PABPPLE", 2008),  # dropped (year - 16)
            ("IT05PABPPLE", 2030),  # dropped (> year)
            ("IT06ABCDE", 2020),    # kept (within window)
        ],
    )
    pp_params = inputs_dir / "1.4_PROJECTION_PARAMETERS_ENTITY.xlsx"
    _build_projection_parameters_fixture(pp_params)
    mp_goc_seg = inputs_dir / "1.5_MP_GOC_SEG.xlsx"
    _build_mp_goc_seg_fixture(mp_goc_seg, [])

    outputs = run_astra_phase1(
        input_paths=[ceded, pp_params, mp_goc_seg],
        run_dir=tmp_path,
        entity_id=6,
        entity_name="AAI",
        year=2024,
        semester=2,
        health_perimeter_gocs=[],
    )

    # Only 3 of the 5 pairs survive the filter
    nb_rows = list(
        openpyxl.load_workbook(outputs[0])["NEW_BUSINESS_PPOS"].iter_rows(values_only=True)
    )
    assert nb_rows == [
        ("GOC_ID", "VARIABLE_NAME", 1),
        ("IT05PABPPLE2024", "CROSS_SUB_FASSCHNG", 0),
        ("IT05PABPPLE2009", "CROSS_SUB_FASSCHNG", 0),
        ("IT06ABCDE2020", "CROSS_SUB_FASSCHNG", 0),
    ]


def test_run_astra_phase1_missing_ceded(tmp_path: Path) -> None:
    other = tmp_path / "irrelevant.xlsx"
    openpyxl.Workbook().save(other)

    with pytest.raises(FileNotFoundError, match="AAI_P&C_Ceded"):
        run_astra_phase1(
            input_paths=[other],
            run_dir=tmp_path,
            entity_id=6,
            entity_name="AAI",
            year=2024,
            semester=2,
            health_perimeter_gocs=[],
        )
