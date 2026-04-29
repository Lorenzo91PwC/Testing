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

from typing import Any

from .skill import (
    ASTRA_COHORT_YEAR_SPAN,
    append_actuarial_aom_impact,
    create_coverage_unit,
    create_empty_csv,
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
    update_curve_id_param,
    update_mp_goc,
    update_mp_goc_seg,
    update_projection_parameters_entity,
)

CEDED_SUFFIX = "AAI_P&C_Ceded"
PAYMENT_PATTERNS_SUFFIX = "Payment_Patterns_&_Risk_Adjustments"
PROJECTION_PARAMETERS_SUFFIX = "PROJECTION_PARAMETERS_ENTITY"
MP_GOC_SEG_SUFFIX = "MP_GOC_SEG"
MP_GOC_SUFFIX = "MP_GOC"
ACTUARIAL_AOM_IMPACT_SUFFIX = "ACTUARIAL_AOM_IMPACT"
CURVE_ID_PARAM_SUFFIX = "CURVE_ID_PARAM"


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
) -> dict[str, Any]:
    """Phase 1: Ceded + Payment_Patterns -> four output CSVs.

    Picks the Ceded input file and extracts the unique GoC names from
    column AA of sheet ``AAI_P&C_Ceded_H_NH``. Picks the
    Payment_Patterns_&_Risk_Adjustments file and reads Risk Adjustment
    and Payment Pattern values (sheets ``ra_AAI_REINS`` and
    ``pp_AAI_REINS``) using ``year`` and ``semester``. Writes:

    - ``{run_dir}/MP_LoB.csv`` — columns ``GoC_ID``, ``Entity_ID``.
    - ``{run_dir}/MP_ObservationYear.csv`` — two rows per GoC
      (``@Opening`` with ``year-1`` and ``@Closing`` with ``year``).
    - ``{run_dir}/Risk_Adjustment.csv`` — two rows per GoC with the
      Opening/Closing Risk Adjustment values.
    - ``{run_dir}/Payment_pattern.csv`` — 25 columns (``GoC``, ``Year``,
      ``0`` .. ``22``), two rows per GoC (``year`` and ``year-1``).

    Returns a dict with:
    - ``outputs``: the produced files in order;
    - ``goc_names``: the unique GoC list extracted from the Ceded file;
    - ``year`` and ``semester``: echoed for downstream consumers (e.g.
      the Astra page reads them from session state to avoid a duplicate
      input form).

    ``entity_name`` is accepted for symmetry with the UI call site but is
    not used by the current transformations.
    """
    del entity_name  # reserved for future phases

    ceded_path = _find_file_by_suffix(input_paths, CEDED_SUFFIX)
    goc_names = extract_unique_goc_names(str(ceded_path))["values"]

    mp_lob_path = run_dir / "MP_LoB.csv"
    create_mp_lob(
        goc_names=goc_names,
        entity_id=entity_id,
        output_path=str(mp_lob_path),
    )

    mp_obs_path = run_dir / "MP_ObservationYear.csv"
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
    ra_path = run_dir / "Risk_Adjustment.csv"
    create_risk_adjustment(
        goc_names=goc_names,
        values=ra_values,
        output_path=str(ra_path),
    )

    pp_rows = lookup_payment_pattern_values(
        path=str(pp_path),
        goc_names=goc_names,
        year=year,
        semester=semester,
    )
    payment_pattern_path = run_dir / "Payment_pattern.csv"
    create_payment_pattern(
        rows=pp_rows,
        output_path=str(payment_pattern_path),
    )

    return {
        "outputs": [mp_lob_path, mp_obs_path, ra_path, payment_pattern_path],
        "goc_names": goc_names,
        "year": year,
        "semester": semester,
    }


# ---------------------------------------------------------------------------
# Astra
# ---------------------------------------------------------------------------
def run_astra_phase1(
    input_paths: list[Path],
    run_dir: Path,
    entity_id: int,
    entity_name: str,
    year: int,
    semester: int,
    business_type: str,
    health_perimeter_gocs: list[str],
    actuarial_aom_impact_pairs: list[tuple[str, Any]],
    closing_curve_name: str,
    opening_curve_name: str,
) -> list[Path]:
    """Astra Phase 1.

    Reads the Ceded input file (suffix ``AAI_P&C_Ceded``): from sheet
    ``AAI_P&C_Ceded_H_NH``, columns AA (GoC code) and AB (cohort year)
    are joined into ``GOC_ID`` strings (e.g. ``IT05PABPPLE2024``) and
    deduped. The resulting list is then capped to the 16-year window
    ``[year - 15, year]`` — pairs outside that window are dropped.

    Writes:

    - ``{run_dir}/NEW_BUSINESS_PPOS.csv`` — one row per kept pair.
    - ``{run_dir}/COVERAGE_UNIT.csv`` — one row per kept pair.
    - ``{run_dir}/REINSURANCE.csv`` — two variables × N pairs per GoC.
    - ``{run_dir}/MANDATORY_ACTUALS.csv`` — 16 variables × N pairs per GoC.
    - ``{run_dir}/PROJECTION_PARAMETERS_ENTITY.csv`` — uploaded copy of
      the Projection Parameters file with rule-based VALUE edits applied
      based on ``year`` and ``semester``.
    - ``{run_dir}/MP_GOC_SEG.csv`` — uploaded copy of the MP_GOC_SEG
      file with ``P&C`` rewritten to ``HLTH_PC`` in columns A and C for
      rows whose GoC name (first 11 chars of column B) is in
      ``health_perimeter_gocs``.
    - ``{run_dir}/MP_GOC.csv`` — uploaded copy of the MP_GOC file with
      columns E, F, L, P rewritten per ``year``, ``semester`` and
      ``business_type``. ``business_type`` is ``"Diretto"`` or ``"Ceduto"``.
    - ``{run_dir}/ACTUARIAL_AOM_IMPACT.csv`` — historical file with new
      rows appended (one per ``(goc_id, step_id, value)`` combination
      from the kept pairs and ``actuarial_aom_impact_pairs``); sorted
      by columns A and B.
    - ``{run_dir}/CURVE_ID_PARAM.csv`` — uploaded copy with column C
      filled per VARIABLE_NAME: ``CLOSING_CURVE_ID`` ->
      ``closing_curve_name``, ``OPENING_CURVE_ID`` ->
      ``opening_curve_name``, ``CREDITED_RATE_CURVE_ID`` -> column A.
    - ``{run_dir}/OCI_OPTION_CF_CLOSING.csv`` — empty placeholder.
    - ``{run_dir}/OCI_OPTION_CF_OPENING.csv`` — empty placeholder.

    ``entity_id`` and ``entity_name`` are accepted for symmetry with the
    Astra UI form but are not used by the current transformations.
    """
    del entity_id, entity_name  # reserved for future skills

    ceded_path = _find_file_by_suffix(input_paths, CEDED_SUFFIX)
    raw_pairs = extract_unique_goc_cohort_pairs(str(ceded_path))["pairs"]

    # Cap at the 16-year window around the analysis year.
    min_year = year - (ASTRA_COHORT_YEAR_SPAN - 1)
    pairs = [p for p in raw_pairs if min_year <= p["year"] <= year]

    new_business_path = run_dir / "NEW_BUSINESS_PPOS.csv"
    create_new_business_ppos(pairs=pairs, output_path=str(new_business_path))

    coverage_unit_path = run_dir / "COVERAGE_UNIT.csv"
    create_coverage_unit(pairs=pairs, output_path=str(coverage_unit_path))

    reinsurance_path = run_dir / "REINSURANCE.csv"
    create_reinsurance(pairs=pairs, output_path=str(reinsurance_path))

    mandatory_actuals_path = run_dir / "MANDATORY_ACTUALS.csv"
    create_mandatory_actuals(pairs=pairs, output_path=str(mandatory_actuals_path))

    projection_params_in = _find_file_by_suffix(
        input_paths, PROJECTION_PARAMETERS_SUFFIX
    )
    projection_params_path = run_dir / "PROJECTION_PARAMETERS_ENTITY.csv"
    update_projection_parameters_entity(
        input_path=str(projection_params_in),
        output_path=str(projection_params_path),
        year=year,
        semester=semester,
    )

    mp_goc_seg_in = _find_file_by_suffix(input_paths, MP_GOC_SEG_SUFFIX)
    mp_goc_seg_path = run_dir / "MP_GOC_SEG.csv"
    update_mp_goc_seg(
        input_path=str(mp_goc_seg_in),
        output_path=str(mp_goc_seg_path),
        health_perimeter_gocs=health_perimeter_gocs,
    )

    mp_goc_in = _find_file_by_suffix(input_paths, MP_GOC_SUFFIX)
    mp_goc_path = run_dir / "MP_GOC.csv"
    update_mp_goc(
        input_path=str(mp_goc_in),
        output_path=str(mp_goc_path),
        year=year,
        semester=semester,
        business_type=business_type,
    )

    aom_impact_in = _find_file_by_suffix(input_paths, ACTUARIAL_AOM_IMPACT_SUFFIX)
    aom_impact_path = run_dir / "ACTUARIAL_AOM_IMPACT.csv"
    append_actuarial_aom_impact(
        input_path=str(aom_impact_in),
        output_path=str(aom_impact_path),
        pairs=pairs,
        step_id_value_pairs=actuarial_aom_impact_pairs,
    )

    curve_id_param_in = _find_file_by_suffix(input_paths, CURVE_ID_PARAM_SUFFIX)
    curve_id_param_path = run_dir / "CURVE_ID_PARAM.csv"
    update_curve_id_param(
        input_path=str(curve_id_param_in),
        output_path=str(curve_id_param_path),
        closing_curve_name=closing_curve_name,
        opening_curve_name=opening_curve_name,
    )

    closing_path = run_dir / "OCI_OPTION_CF_CLOSING.csv"
    opening_path = run_dir / "OCI_OPTION_CF_OPENING.csv"
    create_empty_csv(str(closing_path))
    create_empty_csv(str(opening_path))

    return [
        new_business_path,
        coverage_unit_path,
        reinsurance_path,
        mandatory_actuals_path,
        projection_params_path,
        mp_goc_seg_path,
        mp_goc_path,
        aom_impact_path,
        curve_id_param_path,
        closing_path,
        opening_path,
    ]
