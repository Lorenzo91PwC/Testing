"""Tests for excel_pipeline.pipeline (deterministic, no API calls)."""
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from excel_pipeline.pipeline import run_phase1


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
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ra_AAI_REINS"
    ws.cell(row=1, column=7, value="GoC")
    for i, h in enumerate(["HY_2024", "FY_2024", "HY_2025", "FY_2025"], start=8):
        ws.cell(row=1, column=i, value=h)
    data = [
        ("Motor", 100, 110, 120, 130),
        ("Property", 200, 210, 220, 230),
        ("Liability", 300, 310, 320, 330),
    ]
    for r, (goc, *vals) in enumerate(data, start=2):
        ws.cell(row=r, column=7, value=goc)
        for j, v in enumerate(vals):
            ws.cell(row=r, column=8 + j, value=v)
    wb.save(path)


def test_run_phase1_happy_path(tmp_path: Path) -> None:
    inputs_dir = tmp_path / "inputs"
    inputs_dir.mkdir()
    ceded = inputs_dir / "1.1_2025.12.31_AAI_P&C_Ceded.xlsx"
    _build_ceded_fixture(ceded)
    payments = inputs_dir / "1.2_2025.12.31_Payment_Patterns_&_Risk_Adjustments.xlsx"
    _build_payment_patterns_fixture(payments)

    outputs = run_phase1(
        input_paths=[ceded, payments],
        run_dir=tmp_path,
        entity_id=6,
        entity_name="AAI",
        year=2025,
        semester=2,
    )

    assert outputs == [
        tmp_path / "MP_LoB.xlsx",
        tmp_path / "MP_ObservationYear.xlsx",
        tmp_path / "Risk_Adjustment.xlsx",
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

    outputs = run_phase1(
        input_paths=[unrelated, ceded, payments],
        run_dir=tmp_path,
        entity_id=14,
        entity_name="MPS",
        year=2025,
        semester=1,
    )

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
