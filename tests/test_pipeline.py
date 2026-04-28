"""Tests for excel_pipeline.pipeline (deterministic, no API calls)."""
from __future__ import annotations

import csv
from pathlib import Path

import openpyxl
import pytest

from excel_pipeline.pipeline import run_astra_phase1, run_phase1


def _read_csv(path: Path) -> list[tuple]:
    """Read CSV rows; coerce numeric cells to int/float, '' to None."""
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows: list[tuple] = []
        for raw in reader:
            cells = []
            for cell in raw:
                if cell == "":
                    cells.append(None)
                    continue
                try:
                    cells.append(int(cell))
                    continue
                except ValueError:
                    pass
                try:
                    cells.append(float(cell))
                    continue
                except ValueError:
                    pass
                cells.append(cell)
            rows.append(tuple(cells))
        return rows


def _write_csv(path: Path, rows: list[tuple]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        for row in rows:
            writer.writerow(["" if v is None else v for v in row])


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
        tmp_path / "MP_LoB.csv",
        tmp_path / "MP_ObservationYear.csv",
        tmp_path / "Risk_Adjustment.csv",
        tmp_path / "Payment_pattern.csv",
    ]
    for p in outputs:
        assert p.exists()

    mp_lob = _read_csv(outputs[0])
    assert mp_lob == [
        ("GoC_ID", "Entity_ID"),
        ("Motor", 6),
        ("Property", 6),
        ("Liability", 6),
    ]

    mp_obs = _read_csv(outputs[1])
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
    ra = _read_csv(outputs[2])
    assert ra == [
        ("ObservationID", "Risk_Adjustment"),
        ("Motor@Opening", 110),
        ("Motor@Closing", 130),
        ("Property@Opening", 210),
        ("Property@Closing", 230),
        ("Liability@Opening", 310),
        ("Liability@Closing", 330),
    ]

    pp = _read_csv(outputs[3])
    # Header columns "0".."22" coerce to ints under _read_csv; the header
    # check is on the textual prefix only.
    assert pp[0][:2] == ("GoC", "Year")
    assert pp[0][2:] == tuple(range(23))
    # For each GoC, reference year first then year-1.
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

    mp_lob_rows = _read_csv(outputs[0])
    assert mp_lob_rows[1:] == [("Motor", 14), ("Property", 14), ("Liability", 14)]

    # H1 2025: Closing = HY_2025, Opening = HY_2024
    ra_rows = _read_csv(outputs[2])
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
    rows: list[tuple] = [
        ("PARAMETER", "VALUE"),
        ("PROJECTED_PERIODS", 110),
        ("CF_TIMESTEP", "SEMESTRIAL"),
        ("REPORTING_MONTH", "12_DECEMBER"),
        ("FX_OPENING_DATE", "1M25"),
        ("FX_AVERAGE_DATE", "HY25"),
        ("FX_CLOSING_DATE", "FY25"),
        ("FX_REPORTING_DATE", 20251231),
    ]
    _write_csv(path, rows)


def _build_mp_goc_seg_fixture(path: Path, rows: list[tuple]) -> None:
    out: list[tuple] = [("GOC_SEG_ID", "GOC_ID", "SEG_ID", "ALLOCATION_RATIO")]
    out.extend(rows)
    _write_csv(path, out)


def _build_aom_impact_fixture(path: Path, rows: list[tuple]) -> None:
    out: list[tuple] = [("GOC_ID", "STEP_ID", 1)]
    out.extend(rows)
    _write_csv(path, out)


def _build_curve_id_param_fixture(path: Path, rows: list[tuple]) -> None:
    out: list[tuple] = [("GOC_ID", "VARIABLE_NAME", 1)]
    out.extend(rows)
    _write_csv(path, out)


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
    pp_params = inputs_dir / "1.4_2024.12.31_PROJECTION_PARAMETERS_ENTITY.csv"
    _build_projection_parameters_fixture(pp_params)
    mp_goc_seg = inputs_dir / "1.5_2024.12.31_MP_GOC_SEG.csv"
    _build_mp_goc_seg_fixture(
        mp_goc_seg,
        [
            ("IT05PABPPLE2024_02_P&C", "IT05PABPPLE2024", "02_P&C", 1),
            ("IT05RRIEEBB2024_02_P&C", "IT05RRIEEBB2024", "02_P&C", 1),
        ],
    )
    aom_impact = inputs_dir / "1.6_2024.12.31_ACTUARIAL_AOM_IMPACT.csv"
    _build_aom_impact_fixture(
        aom_impact,
        [("IT05PABPPLE2024", "PVFC_LIC_UNWIND", 0)],  # historical entry
    )
    curve_id_param = inputs_dir / "1.7_2024.12.31_CURVE_ID_PARAM.csv"
    _build_curve_id_param_fixture(
        curve_id_param,
        [
            ("IT05PABPPLE2024", "CLOSING_CURVE_ID", None),
            ("IT05PABPPLE2024", "OPENING_CURVE_ID", None),
            ("IT05PABPPLE2024", "CREDITED_RATE_CURVE_ID", None),
        ],
    )

    outputs = run_astra_phase1(
        input_paths=[ceded, pp_params, mp_goc_seg, aom_impact, curve_id_param],
        run_dir=tmp_path,
        entity_id=6,
        entity_name="AAI",
        year=2024,
        semester=2,
        health_perimeter_gocs=["IT05RRIEEBB"],
        actuarial_aom_impact_pairs=[("DA_LIC_OP", 0), ("DA_LIC_CLO", 0)],
        closing_curve_name="pippo",
        opening_curve_name="carlo",
    )

    assert outputs == [
        tmp_path / "NEW_BUSINESS_PPOS.csv",
        tmp_path / "COVERAGE_UNIT.csv",
        tmp_path / "REINSURANCE.csv",
        tmp_path / "MANDATORY_ACTUALS.csv",
        tmp_path / "PROJECTION_PARAMETERS_ENTITY.csv",
        tmp_path / "MP_GOC_SEG.csv",
        tmp_path / "ACTUARIAL_AOM_IMPACT.csv",
        tmp_path / "CURVE_ID_PARAM.csv",
        tmp_path / "OCI_OPTION_CF_CLOSING.csv",
        tmp_path / "OCI_OPTION_CF_OPENING.csv",
    ]
    for p in outputs:
        assert p.exists()

    # NEW_BUSINESS_PPOS: one row per pair (3 pairs in fixture)
    nb_rows = _read_csv(outputs[0])
    assert nb_rows == [
        ("GOC_ID", "VARIABLE_NAME", 1),
        ("IT05PABPPLE2024", "CROSS_SUB_FASSCHNG", 0),
        ("IT05PABPPLE2023", "CROSS_SUB_FASSCHNG", 0),
        ("IT06ABCDE2024", "CROSS_SUB_FASSCHNG", 0),
    ]

    cu_rows = _read_csv(outputs[1])
    assert len(cu_rows) == 1 + 3
    assert cu_rows[1] == ("IT05PABPPLE2024", 1) + (0,) * 100
    assert cu_rows[3] == ("IT06ABCDE2024", 1) + (0,) * 100

    rein_rows = _read_csv(outputs[2])
    assert rein_rows == [
        ("GOC_ID", "VARIABLE_NAME", 1, "T"),
        ("IT05PABPPLE2024", "LOSSRECO_IFE_ALLOCATION", 0, 2024),
        ("IT05PABPPLE2023", "LOSSRECO_IFE_ALLOCATION", 0, 2023),
        ("IT05PABPPLE2024", "LOSSRECO_CLOSING", 0, 2024),
        ("IT05PABPPLE2023", "LOSSRECO_CLOSING", 0, 2023),
        ("IT06ABCDE2024", "LOSSRECO_IFE_ALLOCATION", 0, 2024),
        ("IT06ABCDE2024", "LOSSRECO_CLOSING", 0, 2024),
    ]

    ma_rows = _read_csv(outputs[3])
    assert len(ma_rows) == 1 + 3 * 16
    assert ma_rows[1] == ("IT05PABPPLE2024", "ACTUAL_PREMIUM_CF_PAST_SERVICE", 0)
    assert ma_rows[17] == ("IT05PABPPLE2023", "ACTUAL_PREMIUM_CF_PAST_SERVICE", 0)
    assert ma_rows[33] == ("IT06ABCDE2024", "ACTUAL_PREMIUM_CF_PAST_SERVICE", 0)

    pp_rows = _read_csv(outputs[4])
    by_param = {row[0]: row[1] for row in pp_rows[1:]}
    assert by_param["CF_TIMESTEP"] == "YEARLY"
    assert by_param["REPORTING_MONTH"] == "12_DECEMBER"
    assert by_param["FX_OPENING_DATE"] == "1M24"
    assert by_param["FX_AVERAGE_DATE"] == "HY24"
    assert by_param["FX_CLOSING_DATE"] == "FY24"
    assert by_param["FX_REPORTING_DATE"] == 20241231
    assert by_param["PROJECTED_PERIODS"] == 110

    seg_rows = _read_csv(outputs[5])
    assert seg_rows[1] == ("IT05PABPPLE2024_02_P&C", "IT05PABPPLE2024", "02_P&C", 1)
    assert seg_rows[2] == (
        "IT05RRIEEBB2024_02_HLTH_PC", "IT05RRIEEBB2024", "02_HLTH_PC", 1,
    )

    aom_rows = _read_csv(outputs[6])
    assert aom_rows[0] == ("GOC_ID", "STEP_ID", 1)
    assert aom_rows[1:] == [
        ("IT05PABPPLE2023", "DA_LIC_CLO", 0),
        ("IT05PABPPLE2023", "DA_LIC_OP", 0),
        ("IT05PABPPLE2024", "DA_LIC_CLO", 0),
        ("IT05PABPPLE2024", "DA_LIC_OP", 0),
        ("IT05PABPPLE2024", "PVFC_LIC_UNWIND", 0),
        ("IT06ABCDE2024", "DA_LIC_CLO", 0),
        ("IT06ABCDE2024", "DA_LIC_OP", 0),
    ]

    curve_rows = _read_csv(outputs[7])
    assert curve_rows == [
        ("GOC_ID", "VARIABLE_NAME", 1),
        ("IT05PABPPLE2024", "CLOSING_CURVE_ID", "pippo"),
        ("IT05PABPPLE2024", "OPENING_CURVE_ID", "carlo"),
        ("IT05PABPPLE2024", "CREDITED_RATE_CURVE_ID", "IT05PABPPLE2024"),
    ]


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
    pp_params = inputs_dir / "1.4_PROJECTION_PARAMETERS_ENTITY.csv"
    _build_projection_parameters_fixture(pp_params)
    mp_goc_seg = inputs_dir / "1.5_MP_GOC_SEG.csv"
    _build_mp_goc_seg_fixture(mp_goc_seg, [])
    aom_impact = inputs_dir / "1.6_ACTUARIAL_AOM_IMPACT.csv"
    _build_aom_impact_fixture(aom_impact, [])
    curve_id_param = inputs_dir / "1.7_CURVE_ID_PARAM.csv"
    _build_curve_id_param_fixture(curve_id_param, [])

    outputs = run_astra_phase1(
        input_paths=[ceded, pp_params, mp_goc_seg, aom_impact, curve_id_param],
        run_dir=tmp_path,
        entity_id=6,
        entity_name="AAI",
        year=2024,
        semester=2,
        health_perimeter_gocs=[],
        actuarial_aom_impact_pairs=[],
        closing_curve_name="",
        opening_curve_name="",
    )

    # Only 3 of the 5 pairs survive the filter
    nb_rows = _read_csv(outputs[0])
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
            actuarial_aom_impact_pairs=[],
            closing_curve_name="",
            opening_curve_name="",
        )
