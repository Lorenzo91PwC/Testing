"""Deterministic pipeline — plain Python, no LLM calls.

The main processing flow is defined here as ordinary Python functions that
call skill functions in a fixed sequence. No Anthropic API key is required
to run it; no agent picks the steps.

Keep the function signatures aligned with `app.py` so the UI can invoke
them directly. The Claude-powered orchestrator in ``orchestrator.py`` is
only used for the optional ad-hoc chat panel.
"""
from __future__ import annotations

from pathlib import Path

from .skill import (
    create_mp_lob,
    create_mp_observation_year,
    create_risk_adjustment,
    extract_unique_goc_names,
    lookup_risk_adjustment_values,
)

CEDED_SUFFIX = "AAI_P&C_Ceded"
PAYMENT_PATTERNS_SUFFIX = "Payment_Patterns_&_Risk_Adjustments"


def _find_file_by_suffix(paths: list[Path], suffix: str) -> Path:
    """Return the single uploaded file whose stem ends with ``suffix``."""
    matches = [p for p in paths if p.stem.endswith(suffix)]
    if not matches:
        raise FileNotFoundError(
            f"No input file matches the expected suffix '{suffix}'. "
            f"Got: {[p.name for p in paths]}"
        )
    if len(matches) > 1:
        raise ValueError(
            f"Multiple files match the suffix '{suffix}': "
            f"{[p.name for p in matches]}"
        )
    return matches[0]


def run_phase1(
    input_paths: list[Path],
    run_dir: Path,
    entity_id: int,
    entity_name: str,
    year: int,
    semester: int,
) -> list[Path]:
    """Phase 1: Ceded + Payment_Patterns -> MP_LoB, MP_ObservationYear, Risk_Adjustment.

    Picks the Ceded input file and extracts the unique GoC names from
    column AA of sheet ``AAI_P&C_Ceded_H_NH``. Picks the
    Payment_Patterns_&_Risk_Adjustments file and reads the Opening/Closing
    Risk Adjustment values for each GoC (sheet ``ra_AAI_REINS``, year
    columns selected by ``year`` and ``semester``). Writes:

    - ``{run_dir}/MP_LoB.xlsx`` — columns ``GoC_ID``, ``Entity_ID``.
    - ``{run_dir}/MP_ObservationYear.xlsx`` — two rows per GoC
      (``@Opening`` with ``year-1`` and ``@Closing`` with ``year``).
    - ``{run_dir}/Risk_Adjustment.xlsx`` — two rows per GoC with the
      Opening/Closing Risk Adjustment values.

    Returns the list of output paths in the order they were produced.
    ``entity_name`` is accepted for symmetry with the UI call site but is
    not used by the current transformations.
    """
    del entity_name  # reserved for future phases

    ceded_path = _find_file_by_suffix(input_paths, CEDED_SUFFIX)
    goc_names = extract_unique_goc_names(str(ceded_path))["values"]

    mp_lob_path = run_dir / "MP_LoB.xlsx"
    create_mp_lob(
        goc_names=goc_names,
        entity_id=entity_id,
        output_path=str(mp_lob_path),
    )

    mp_obs_path = run_dir / "MP_ObservationYear.xlsx"
    create_mp_observation_year(
        goc_names=goc_names,
        year=year,
        output_path=str(mp_obs_path),
    )

    pp_path = _find_file_by_suffix(input_paths, PAYMENT_PATTERNS_SUFFIX)
    ra_values = lookup_risk_adjustment_values(
        path=str(pp_path),
        goc_names=goc_names,
        year=year,
        semester=semester,
    )
    ra_path = run_dir / "Risk_Adjustment.xlsx"
    create_risk_adjustment(
        goc_names=goc_names,
        values=ra_values,
        output_path=str(ra_path),
    )

    return [mp_lob_path, mp_obs_path, ra_path]
