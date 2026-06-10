"""Tests for excel_pipeline.pipeline (deterministic, no API calls)."""
from __future__ import annotations

import csv
from pathlib import Path

import openpyxl
import pytest

from excel_pipeline.pipeline import (
    run_astra_phase1,
    run_phase1,
    validate_sunrise_inputs,
)


def _read_csv(path: Path) -> list[tuple]:
    """Read CSV rows; coerce numeric cells to int/float, '' to None.

    Uses ``;`` field separator and converts ``,`` decimal back to ``.``
    so floats round-trip correctly (matches ``_write_csv_rows``).
    """
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=";")
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
                    cells.append(float(cell.replace(",", ".")))
                    continue
                except ValueError:
                    pass
                cells.append(cell)
            rows.append(tuple(cells))
        return rows


def _write_csv(path: Path, rows: list[tuple]) -> None:
    """Write a CSV fixture using the project convention (``;`` separator)."""
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
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


def _build_input_sunrise_workbook(
    path: Path,
    gocs: list[str | None],
    *,
    year: int = 2025,
    sinistri: float = 0.0,
    riserva: float = 0.0,
) -> None:
    """Workbook with the new Input_Sunrise sheet.

    By default each GoC is paired with the same ``year`` and zero
    SINISTRI / RISERVA — enough to exercise the row-count path.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Input_Sunrise"
    ws.cell(row=1, column=1, value="GoC")
    ws.cell(row=1, column=2, value="Year")
    ws.cell(row=1, column=3, value="Perimetro")
    ws.cell(row=1, column=4, value="Sinistri")
    ws.cell(row=1, column=5, value="Riserva")
    for i, g in enumerate(gocs, start=2):
        ws.cell(row=i, column=1, value=g)
        ws.cell(row=i, column=2, value=year)
        ws.cell(row=i, column=4, value=sinistri)
        ws.cell(row=i, column=5, value=riserva)
    wb.save(path)


def test_run_phase1_emits_only_mp_model_point(tmp_path: Path) -> None:
    """Phase 1 currently runs only the MP_ModelPoint step."""
    inputs_dir = tmp_path / "inputs"
    inputs_dir.mkdir()
    ceded_curr = inputs_dir / "1.1_2025.12.31_AAI_Ceded.xlsx"
    _build_input_sunrise_workbook(
        ceded_curr, ["Motor", "Property"], year=2025, sinistri=100.0, riserva=50.0,
    )
    ceded_prev = inputs_dir / "1.2_2024.12.31_AAI_Ceded.xlsx"
    _build_input_sunrise_workbook(
        ceded_prev, ["Property", "Liability"], year=2024, sinistri=80.0, riserva=40.0,
    )
    transcodifica = inputs_dir / "1.3_Transcodifica_aggregazione_GOC_H_NH.csv"
    _write_csv(
        transcodifica,
        [
            ("GOC", "Aggregation1", "Aggregation2"),
            ("Motor", "Agg1_Motor", "Agg2_Motor"),
            ("Property", "Agg1_Prop", "Agg2_Prop"),
        ],
    )
    # Payment_Patterns may still be uploaded but is not consumed here.
    payments = inputs_dir / "1.4_2025.12.31_Payment_Patterns_&_Risk_Adjustments.xlsx"
    _build_payment_patterns_fixture(payments)

    result = run_phase1(
        input_paths=[ceded_curr, ceded_prev, transcodifica, payments],
        run_dir=tmp_path,
        entities=[(6, "AAI"), (14, "MPS")],
        year=2025,
        semester=2,
    )

    assert result["entities"] == [(6, "AAI"), (14, "MPS")]
    assert result["year"] == 2025
    assert result["semester"] == 2

    outputs = result["outputs"]
    assert outputs == [
        tmp_path / "MP_ModelPoint.csv",
        tmp_path / "MP_LoB.csv",
        tmp_path / "MP_ObservationYear.csv",
        tmp_path / "Risk_Adjustment.csv",
        tmp_path / "Payment_pattern.csv",
    ]
    for p in outputs:
        assert p.exists()

    # MP_ModelPoint has data rows from both source files.
    rows = _read_csv(outputs[0])
    assert rows[0][0] == "Primary_Key"
    assert len(rows) > 1

    # MP_LoB: one row per (GoC, entity). Filtered to non-zero GoCs
    # (Motor, Property and Liability all have non-zero sinistri/riserva
    # in the fixture). 3 GoCs * 2 entities = 6 rows.
    mp_lob_rows = _read_csv(outputs[1])
    assert mp_lob_rows[0] == ("GoC_ID", "Entity_ID")
    assert len(mp_lob_rows) == 1 + 3 * 2

    # MP_ObservationYear: two rows per GoC (@Opening + @Closing).
    mp_obs_rows = _read_csv(outputs[2])
    assert mp_obs_rows[0] == (
        "ObservationID", "ObservationYear", "LoB_ID", "AdjULAEPagate", "CY",
    )
    assert len(mp_obs_rows) == 1 + 3 * 2

    # Risk_Adjustment: two rows per GoC (@Opening + @Closing).
    ra_rows = _read_csv(outputs[3])
    assert ra_rows[0] == ("ObservationID", "Risk_Adjustment")
    assert len(ra_rows) == 1 + 3 * 2

    # Payment_pattern: 25 columns (GoC, Year, 0..22); 2 rows per GoC.
    pp_rows = _read_csv(outputs[4])
    assert pp_rows[0][:2] == ("GoC", "Year")
    assert pp_rows[0][2:] == tuple(range(23))
    assert len(pp_rows) == 1 + 3 * 2


def test_run_phase1_raises_without_any_ceded_or_assumed(tmp_path: Path) -> None:
    payments = tmp_path / "1.2_2025.12.31_Payment_Patterns_&_Risk_Adjustments.xlsx"
    _build_payment_patterns_fixture(payments)

    with pytest.raises(FileNotFoundError, match="_Ceded|_Assumed"):
        run_phase1(
            input_paths=[payments],
            run_dir=tmp_path,
            entities=[(6, "AAI")],
            year=2025,
            semester=2,
        )


def test_validate_sunrise_inputs_missing_transcodifica(tmp_path: Path) -> None:
    ceded = tmp_path / "1.1_2025.12.31_AAI_Ceded.xlsx"
    _build_input_sunrise_workbook(ceded, ["Motor"])
    ceded_prev = tmp_path / "1.2_2024.12.31_AAI_Ceded.xlsx"
    _build_input_sunrise_workbook(ceded_prev, ["Motor"])

    errors = validate_sunrise_inputs(
        [ceded, ceded_prev], year=2025, semester=2,
    )
    assert any("Transcodifica" in e for e in errors)


def test_validate_sunrise_inputs_missing_ceded_for_one_date(tmp_path: Path) -> None:
    ceded_curr = tmp_path / "1.1_2025.12.31_AAI_Ceded.xlsx"
    _build_input_sunrise_workbook(ceded_curr, ["Motor"])
    transcodifica = tmp_path / "1.3_Transcodifica_aggregazione_GOC_H_NH.csv"
    transcodifica.write_text("dummy\n", encoding="utf-8-sig")

    errors = validate_sunrise_inputs(
        [ceded_curr, transcodifica], year=2025, semester=2,
    )
    # Current-year date is covered; previous-year date is missing
    assert any("2024.12.31" in e for e in errors)
    assert not any("2025.12.31" in e for e in errors)


def test_validate_sunrise_inputs_all_present_no_errors(tmp_path: Path) -> None:
    ceded_curr = tmp_path / "1.1_2025.12.31_AAI_Ceded.xlsx"
    _build_input_sunrise_workbook(ceded_curr, ["Motor"])
    ceded_prev = tmp_path / "1.2_2024.12.31_AAI_Ceded.xlsx"
    _build_input_sunrise_workbook(ceded_prev, ["Motor"])
    transcodifica = tmp_path / "1.3_Transcodifica_aggregazione_GOC_H_NH.csv"
    transcodifica.write_text("dummy\n", encoding="utf-8-sig")

    errors = validate_sunrise_inputs(
        [ceded_curr, ceded_prev, transcodifica], year=2025, semester=2,
    )
    assert errors == []


def test_validate_sunrise_inputs_hy_semester_uses_june_30(tmp_path: Path) -> None:
    ceded_curr = tmp_path / "1.1_2025.06.30_AAI_Ceded.xlsx"
    _build_input_sunrise_workbook(ceded_curr, ["Motor"])
    ceded_prev = tmp_path / "1.2_2024.06.30_AAI_Ceded.xlsx"
    _build_input_sunrise_workbook(ceded_prev, ["Motor"])
    transcodifica = tmp_path / "1.3_Transcodifica_aggregazione_GOC_H_NH.csv"
    transcodifica.write_text("dummy\n", encoding="utf-8-sig")

    errors = validate_sunrise_inputs(
        [ceded_curr, ceded_prev, transcodifica], year=2025, semester=1,
    )
    assert errors == []


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


_MP_GOC_HEADERS_PIPE = (
    "GOC_ID", "MEASUREMENT_MODEL", "ANNUAL_COHORT", "AOM_ID",
    "INCEPTION_CURVE_ID", "TIMING_INCEPTION_CURVE",
    "CSM_RELEASE_RATIO_CURVE_ID", "SHARE_TECH_EXP_ENTITY_SHARE",
    "OCI_OPTION", "OCI_OPTION_LRC", "OCI_OPTION_LIC", "GOC_DURATION",
    "GOC_TYPE_IF_NB", "GOC_CURRENCY", "REPORTING_CURRENCY",
    "GOC_TYPE_REINSURANCE", "AGGREG_1_ID", "AGGREG_2_ID", "AGGREG_3_ID",
    "AGGREG_4_ID", "AGGREG_5_ID",
)


def _build_mp_goc_fixture(path: Path, cohort_years: list[int]) -> None:
    rows: list[tuple] = [_MP_GOC_HEADERS_PIPE]
    for y in cohort_years:
        row = [f"IT{y}", "PAA", y, 20] + [None] * (len(_MP_GOC_HEADERS_PIPE) - 4)
        rows.append(tuple(row))
    _write_csv(path, rows)


def _build_curve_id_param_fixture(path: Path, rows: list[tuple]) -> None:
    out: list[tuple] = [("GOC_ID", "VARIABLE_NAME", 1)]
    out.extend(rows)
    _write_csv(path, out)


def test_run_astra_phase1_uses_pairs_from_sunrise(tmp_path: Path) -> None:
    """Astra now receives (GoC, accident_year) pairs from Sunrise instead of
    reading the legacy AAI_P&C_Ceded file."""
    inputs_dir = tmp_path / "inputs"
    inputs_dir.mkdir()
    sunrise_pairs = [
        {"goc_id": "IT05PABPPLE2024", "goc": "IT05PABPPLE", "year": 2024},
        {"goc_id": "IT05PABPPLE2023", "goc": "IT05PABPPLE", "year": 2023},
        {"goc_id": "IT06ABCDE2024", "goc": "IT06ABCDE", "year": 2024},
    ]
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
    mp_goc = inputs_dir / "1.8_2024.12.31_MP_GOC.csv"
    _build_mp_goc_fixture(mp_goc, [2014, 2024])

    outputs = run_astra_phase1(
        input_paths=[pp_params, mp_goc_seg, mp_goc, aom_impact, curve_id_param],
        run_dir=tmp_path,
        entities=[(6, "AAI")],
        year=2024,
        semester=2,
        business_type="Diretto",
        health_perimeter_gocs=["IT05RRIEEBB"],
        actuarial_aom_impact_pairs=[("DA_LIC_OP", 0), ("DA_LIC_CLO", 0)],
        closing_curve_name="pippo",
        opening_curve_name="carlo",
        goc_cohort_pairs=sunrise_pairs,
    )

    assert outputs == [
        tmp_path / "NEW_BUSINESS_PPOS.csv",
        tmp_path / "COVERAGE_UNIT.csv",
        tmp_path / "REINSURANCE.csv",
        tmp_path / "MANDATORY_ACTUALS.csv",
        tmp_path / "PROJECTION_PARAMETERS_ENTITY.csv",
        tmp_path / "MP_GOC_SEG.csv",
        tmp_path / "MP_GOC.csv",
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
    assert by_param["CF_TIMESTEP"] == "SEMESTRIAL"
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

    mp_goc_rows = _read_csv(outputs[6])
    by_cohort = {row[2]: row for row in mp_goc_rows[1:]}
    # E (idx 4), F (idx 5), L (idx 11), P (idx 15)
    assert by_cohort[2014][4] == "20141231_ITA_LP100"
    assert by_cohort[2024][4] == "20241231_EUR_LP100_FY24_AVG"
    assert by_cohort[2014][11] == (2024 - 2014) * 12
    assert by_cohort[2024][11] == 0
    for cohort in (2014, 2024):
        assert by_cohort[cohort][15] == "2_RE_ASSUMED"

    aom_rows = _read_csv(outputs[7])
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

    curve_rows = _read_csv(outputs[8])
    assert curve_rows == [
        ("GOC_ID", "VARIABLE_NAME", 1),
        ("IT05PABPPLE2024", "CLOSING_CURVE_ID", "pippo"),
        ("IT05PABPPLE2024", "OPENING_CURVE_ID", "carlo"),
        ("IT05PABPPLE2024", "CREDITED_RATE_CURVE_ID", "CR_IT05PABPPLE2024"),
    ]


def test_run_astra_phase1_filters_pairs_outside_window(tmp_path: Path) -> None:
    """Pairs outside [year-15, year] are dropped before generation."""
    inputs_dir = tmp_path / "inputs"
    inputs_dir.mkdir()
    sunrise_pairs = [
        {"goc_id": "IT05PABPPLE2024", "goc": "IT05PABPPLE", "year": 2024},  # kept
        {"goc_id": "IT05PABPPLE2009", "goc": "IT05PABPPLE", "year": 2009},  # kept
        {"goc_id": "IT05PABPPLE2008", "goc": "IT05PABPPLE", "year": 2008},  # dropped
        {"goc_id": "IT05PABPPLE2030", "goc": "IT05PABPPLE", "year": 2030},  # dropped
        {"goc_id": "IT06ABCDE2020", "goc": "IT06ABCDE", "year": 2020},      # kept
    ]
    pp_params = inputs_dir / "1.4_PROJECTION_PARAMETERS_ENTITY.csv"
    _build_projection_parameters_fixture(pp_params)
    mp_goc_seg = inputs_dir / "1.5_MP_GOC_SEG.csv"
    _build_mp_goc_seg_fixture(mp_goc_seg, [])
    aom_impact = inputs_dir / "1.6_ACTUARIAL_AOM_IMPACT.csv"
    _build_aom_impact_fixture(aom_impact, [])
    curve_id_param = inputs_dir / "1.7_CURVE_ID_PARAM.csv"
    _build_curve_id_param_fixture(curve_id_param, [])
    mp_goc = inputs_dir / "1.8_MP_GOC.csv"
    _build_mp_goc_fixture(mp_goc, [])

    outputs = run_astra_phase1(
        input_paths=[pp_params, mp_goc_seg, mp_goc, aom_impact, curve_id_param],
        run_dir=tmp_path,
        entities=[(6, "AAI")],
        year=2024,
        semester=2,
        business_type="Diretto",
        health_perimeter_gocs=[],
        actuarial_aom_impact_pairs=[],
        closing_curve_name="",
        opening_curve_name="",
        goc_cohort_pairs=sunrise_pairs,
    )

    # Only 3 of the 5 pairs survive the filter
    nb_rows = _read_csv(outputs[0])
    assert nb_rows == [
        ("GOC_ID", "VARIABLE_NAME", 1),
        ("IT05PABPPLE2024", "CROSS_SUB_FASSCHNG", 0),
        ("IT05PABPPLE2009", "CROSS_SUB_FASSCHNG", 0),
        ("IT06ABCDE2020", "CROSS_SUB_FASSCHNG", 0),
    ]


def test_run_astra_phase1_missing_required_input(tmp_path: Path) -> None:
    """PROJECTION_PARAMETERS_ENTITY (still mandatory) is missing -> error."""
    other = tmp_path / "irrelevant.xlsx"
    openpyxl.Workbook().save(other)

    with pytest.raises(FileNotFoundError, match="PROJECTION_PARAMETERS_ENTITY"):
        run_astra_phase1(
            input_paths=[other],
            run_dir=tmp_path,
            entities=[(6, "AAI")],
            year=2024,
            semester=2,
            business_type="Diretto",
            health_perimeter_gocs=[],
            actuarial_aom_impact_pairs=[],
            closing_curve_name="",
            opening_curve_name="",
            goc_cohort_pairs=[],
        )
