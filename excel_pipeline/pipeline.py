"""Deterministic pipeline — plain Python, no LLM calls.

The main processing flow is defined here as ordinary Python functions that
call skill functions in a fixed sequence. No Anthropic API key is required
to run it; no agent picks the steps.

Keep the function signatures aligned with `app.py` so the UI can invoke
them directly. The Claude-powered orchestrator in ``orchestrator.py`` is
only used for the optional ad-hoc chat panel.
"""
from __future__ import annotations

import re
from pathlib import Path

from typing import Any

from .skill import (
    ASTRA_COHORT_YEAR_SPAN,
    append_actuarial_aom_impact,
    create_coverage_unit,
    create_empty_csv,
    create_mandatory_actuals,
    create_mp_lob,
    create_mp_model_point,
    create_mp_observation_year,
    create_new_business_ppos,
    create_payment_pattern,
    create_reinsurance,
    create_risk_adjustment,
    extract_input_sunrise_goc_names,
    extract_unique_goc_cohort_pairs,
    extract_unique_goc_names,
    lookup_payment_pattern_values,
    lookup_risk_adjustment_values,
    update_curve_id_param,
    update_mp_goc,
    update_mp_goc_seg,
    update_projection_parameters_entity,
)

# Astra still uses the legacy Ceded file with the AAI_P&C_Ceded_H_NH sheet.
ASTRA_CEDED_SUFFIX = "AAI_P&C_Ceded"
PAYMENT_PATTERNS_SUFFIX = "Payment_Patterns_&_Risk_Adjustments"
PROJECTION_PARAMETERS_SUFFIX = "PROJECTION_PARAMETERS_ENTITY"
MP_GOC_SEG_SUFFIX = "MP_GOC_SEG"
MP_GOC_SUFFIX = "MP_GOC"
ACTUARIAL_AOM_IMPACT_SUFFIX = "ACTUARIAL_AOM_IMPACT"
CURVE_ID_PARAM_SUFFIX = "CURVE_ID_PARAM"

# New Sunrise input model: many _Ceded / _Assumed files (one Input_Sunrise
# sheet each), plus a single Transcodifica master list.
SUNRISE_CEDED_SUFFIX = "_Ceded"
SUNRISE_ASSUMED_SUFFIX = "_Assumed"
TRANSCODIFICA_SUFFIX = "Transcodifica_aggregazione_GOC_H_NH"

_DATE_RE = re.compile(r"(\d{4})\.(\d{2})\.(\d{2})")


def _expected_sunrise_dates(year: int, semester: int) -> tuple[str, str]:
    """Return the two ``YYYY.MM.DD`` strings the Sunrise inputs must cover.

    Semester convention: ``1 = HY = June 30``, ``2 = FY = December 31``.
    The first string is the analysis date, the second is the previous year.
    """
    if semester == 1:
        month_day = "06.30"
    elif semester == 2:
        month_day = "12.31"
    else:
        raise ValueError(f"semester must be 1 or 2, got {semester}")
    return (f"{year}.{month_day}", f"{year - 1}.{month_day}")


def _files_ending_with(paths: list[Path], suffix: str) -> list[Path]:
    """All uploaded files whose stem ends with ``suffix`` (zero or more)."""
    return [p for p in paths if p.stem.endswith(suffix)]


def _files_with_date(paths: list[Path], date_str: str) -> list[Path]:
    """All uploaded files whose name contains ``date_str``."""
    return [p for p in paths if date_str in p.name]


def validate_sunrise_inputs(
    input_paths: list[Path], year: int, semester: int
) -> list[str]:
    """Validate the Sunrise input upload set.

    Required:
    - at least one file with suffix ``Transcodifica_aggregazione_GOC_H_NH``;
    - for each expected date (analysis date and previous year), at least one
      file with suffix ``_Ceded`` and the date in its name.

    ``_Assumed`` files are optional (used for the GoC list if present).

    Returns a list of human-readable error messages; an empty list means
    the upload is OK.
    """
    errors: list[str] = []

    transcodifica_files = _files_ending_with(input_paths, TRANSCODIFICA_SUFFIX)
    if not transcodifica_files:
        errors.append(
            f"Missing the master list file: no upload ends with "
            f"'{TRANSCODIFICA_SUFFIX}'."
        )

    expected_dates = _expected_sunrise_dates(year, semester)
    ceded_files = _files_ending_with(input_paths, SUNRISE_CEDED_SUFFIX)
    for date_str in expected_dates:
        ceded_for_date = [p for p in ceded_files if date_str in p.name]
        if not ceded_for_date:
            errors.append(
                f"Missing a _Ceded file for date {date_str}. Upload at "
                f"least one file ending with '_Ceded' that contains "
                f"{date_str} in its name."
            )

    return errors


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


# Backwards-compat alias for legacy callers/tests still using CEDED_SUFFIX.
CEDED_SUFFIX = ASTRA_CEDED_SUFFIX


def run_phase1(
    input_paths: list[Path],
    run_dir: Path,
    entities: list[tuple[int, str]],
    year: int,
    semester: int,
) -> dict[str, Any]:
    """Phase 1 — currently runs only the MP_ModelPoint step.

    The legacy MP_LoB / MP_ObservationYear / Risk_Adjustment /
    Payment_pattern steps are temporarily skipped while we focus on the
    MP_ModelPoint output. The matching skill functions remain available
    in ``skill.py`` and can be re-enabled here at any time without other
    changes.

    Inputs consumed:
    - the union of ``_Ceded`` / ``_Assumed`` files (split internally by
      analysis date: current-year vs previous-year), used to build
      ``MP_ModelPoint.csv``;
    - the ``Transcodifica_aggregazione_GOC_H_NH`` master list, used to
      populate Aggregation1 / Aggregation2 on each MP_ModelPoint row.

    The Payment_Patterns_&_Risk_Adjustments workbook may still be
    uploaded but is not read in this reduced flow.

    Returns a dict with:
    - ``outputs``: ``[MP_ModelPoint.csv]``;
    - ``entities``, ``year``, ``semester``: echoed for downstream
      consumers (e.g. the Astra page reads them from session state).

    Callers should run ``validate_sunrise_inputs`` first to surface a
    clear UI error when an expected file is missing.
    """
    sunrise_paths = [
        p for p in input_paths
        if p.stem.endswith(SUNRISE_CEDED_SUFFIX)
        or p.stem.endswith(SUNRISE_ASSUMED_SUFFIX)
    ]
    if not sunrise_paths:
        raise FileNotFoundError(
            "No file with suffix '_Ceded' or '_Assumed' was provided. "
            f"Got: {[p.name for p in input_paths]}"
        )

    # For each Sunrise file extract ANNO_RIFERIMENTO from the date in its
    # filename (validated upstream by validate_sunrise_inputs). Files whose
    # name carries no recognizable date are skipped.
    sources: list[tuple[str, int]] = []
    for p in sunrise_paths:
        m = _DATE_RE.search(p.name)
        if m is None:
            continue
        sources.append((str(p), int(m.group(1))))

    transcodifica_path = _find_file_by_suffix(input_paths, TRANSCODIFICA_SUFFIX)
    mp_model_point_path = run_dir / "MP_ModelPoint.csv"
    create_mp_model_point(
        sources=sources,
        transcodifica_path=str(transcodifica_path),
        output_path=str(mp_model_point_path),
        year=year,
    )

    return {
        "outputs": [mp_model_point_path],
        "entities": entities,
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
