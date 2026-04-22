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
    extract_unique_goc_names,
)

CEDED_SUFFIX = "AAI_P&C_Ceded"


def _find_ceded_file(paths: list[Path]) -> Path:
    """Return the uploaded file whose stem ends with the Ceded suffix."""
    matches = [p for p in paths if p.stem.endswith(CEDED_SUFFIX)]
    if not matches:
        raise FileNotFoundError(
            f"No input file matches the expected suffix '{CEDED_SUFFIX}'. "
            f"Got: {[p.name for p in paths]}"
        )
    if len(matches) > 1:
        raise ValueError(
            f"Multiple files match the suffix '{CEDED_SUFFIX}': "
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
    """Phase 1: Ceded → MP_LoB + MP_ObservationYear.

    Picks the input file ending in the Ceded suffix and extracts the unique
    GoC names from column AA of sheet ``AAI_P&C_Ceded_H_NH``. Then writes:

    - ``{run_dir}/MP_LoB.xlsx`` — columns ``GoC_ID``, ``Entity_ID``.
    - ``{run_dir}/MP_ObservationYear.xlsx`` — two rows per GoC
      (``@Opening`` with ``year-1`` and ``@Closing`` with ``year``), columns
      ``ObservationID``, ``ObservationYear``, ``LoB_ID``, ``AdjULAEPagate``,
      ``CY``.

    Returns the list of output paths in the order they were produced.
    ``entity_name`` and ``semester`` are accepted for symmetry with the UI
    call site but are not used by the current transformations.
    """
    del entity_name, semester  # reserved for future phases

    ceded_path = _find_ceded_file(input_paths)
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

    return [mp_lob_path, mp_obs_path]
