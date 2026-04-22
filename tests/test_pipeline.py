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


def test_run_phase1_happy_path(tmp_path: Path) -> None:
    inputs_dir = tmp_path / "inputs"
    inputs_dir.mkdir()
    ceded = inputs_dir / "1.1_2025.12.31_AAI_P&C_Ceded.xlsx"
    _build_ceded_fixture(ceded)

    outputs = run_phase1(
        input_paths=[ceded],
        run_dir=tmp_path,
        entity_id=6,
        entity_name="AAI",
        year=2025,
        semester=2,
    )

    assert outputs == [
        tmp_path / "MP_LoB.xlsx",
        tmp_path / "MP_ObservationYear.xlsx",
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


def test_run_phase1_picks_matching_file(tmp_path: Path) -> None:
    """With extra unrelated files, only the Ceded one is used."""
    inputs_dir = tmp_path / "inputs"
    inputs_dir.mkdir()
    ceded = inputs_dir / "1.1_2025.12.31_AAI_P&C_Ceded.xlsx"
    _build_ceded_fixture(ceded)
    unrelated = inputs_dir / "notes.xlsx"
    openpyxl.Workbook().save(unrelated)

    outputs = run_phase1(
        input_paths=[unrelated, ceded],
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
