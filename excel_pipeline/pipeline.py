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
    extract_health_perimeter_gocs,
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
    gocs_to_exclude: list[str] | None = None,
) -> dict[str, Any]:
    """Phase 1 — Sunrise main outputs.

    Inputs consumed:
    - the union of ``_Ceded`` / ``_Assumed`` files, used to build the
      GoC list and ``MP_ModelPoint.csv``;
    - the ``Transcodifica_aggregazione_GOC_H_NH`` master list, used for
      MP_ModelPoint's Aggregation1 / Aggregation2 columns;
    - the ``Payment_Patterns_&_Risk_Adjustments`` workbook, used for
      ``Risk_Adjustment.csv`` and ``Payment_pattern.csv`` (sheets
      ``ra_AAI_REINS`` and ``pp_AAI_REINS``).

    Outputs produced:
    - ``MP_ModelPoint.csv`` (driven by SINISTRI / RISERVA_SINISTRI; GoCs
      whose values are all zero are excluded);
    - ``MP_LoB.csv`` (one row per ``(GoC, entity)`` pair);
    - ``MP_ObservationYear.csv`` (two rows per GoC: ``@Opening`` with
      ``year - 1`` and ``@Closing`` with ``year``);
    - ``Risk_Adjustment.csv`` (two rows per GoC with the Opening/Closing
      Risk Adjustment values from the Payment_Patterns workbook);
    - ``Payment_pattern.csv`` (25 columns ``GoC, Year, 0..22``; two rows
      per GoC for ``year`` and ``year - 1``).

    The GoC list shared by every downstream output is whatever
    ``create_mp_model_point`` returns — i.e. it follows the same
    all-zero-drop rule as MP_ModelPoint, so all files report the same
    GoC universe.

    Returns a dict with:
    - ``outputs``: the produced files in order;
    - ``entities``, ``year``, ``semester``: echoed.

    Callers should run ``input_validation.validate_sunrise_inputs``
    first to surface a clear UI error when an expected file is missing.
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
    mp_result = create_mp_model_point(
        sources=sources,
        transcodifica_path=str(transcodifica_path),
        output_path=str(mp_model_point_path),
        year=year,
        gocs_to_exclude=gocs_to_exclude,
    )
    goc_list = mp_result["goc_list"]

    mp_lob_path = run_dir / "MP_LoB.csv"
    create_mp_lob(
        goc_names=goc_list,
        entities=entities,
        output_path=str(mp_lob_path),
    )

    mp_obs_path = run_dir / "MP_ObservationYear.csv"
    create_mp_observation_year(
        goc_names=goc_list,
        year=year,
        output_path=str(mp_obs_path),
    )

    pp_path = _find_file_by_suffix(input_paths, PAYMENT_PATTERNS_SUFFIX)
    ra_values = lookup_risk_adjustment_values(
        path=str(pp_path),
        goc_names=goc_list,
        year=year,
        semester=semester,
    )
    ra_path = run_dir / "Risk_Adjustment.csv"
    create_risk_adjustment(
        goc_names=goc_list,
        values=ra_values,
        output_path=str(ra_path),
    )

    pp_rows = lookup_payment_pattern_values(
        path=str(pp_path),
        goc_names=goc_list,
        year=year,
        semester=semester,
    )
    payment_pattern_path = run_dir / "Payment_pattern.csv"
    create_payment_pattern(
        rows=pp_rows,
        output_path=str(payment_pattern_path),
    )

    return {
        "outputs": [
            mp_model_point_path,
            mp_lob_path,
            mp_obs_path,
            ra_path,
            payment_pattern_path,
        ],
        "goc_cohort_pairs": mp_result["goc_cohort_pairs"],
        "health_perimeter_gocs": extract_health_perimeter_gocs(
            str(transcodifica_path)
        ),
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
    entities: list[tuple[int, str]],
    year: int,
    semester: int,
    business_type: str,
    health_perimeter_gocs: list[str],
    actuarial_aom_impact_pairs: list[tuple[str, Any]],
    closing_curve_name: str,
    opening_curve_name: str,
    goc_cohort_pairs: list[dict[str, Any]],
    gocs_to_exclude: list[str] | None = None,
) -> list[Path]:
    """Astra Phase 1.

    The ``(GoC, accident_year)`` pairs that drive NEW_BUSINESS_PPOS,
    COVERAGE_UNIT, REINSURANCE and MANDATORY_ACTUALS come from the
    Sunrise run (``goc_cohort_pairs`` parameter — typically read from
    session state). The legacy ``AAI_P&C_Ceded`` upload is no longer
    consumed by Astra. The supplied pairs are filtered to the 16-year
    window ``[year - 15, year]`` before driving the outputs.

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
    - ``OCI_OPTION_CF_CLOSING.csv`` and ``OCI_OPTION_CF_OPENING.csv``
      are currently **not emitted** — their population rules are still
      TODO. The disabled call sites are commented in the pipeline body
      and can be re-enabled in two edits when the spec lands.

    ``goc_cohort_pairs`` is the list of ``{"goc_id", "goc", "year"}``
    dicts produced by the Sunrise run (typically read from session
    state by the Astra page). It replaces the legacy
    ``AAI_P&C_Ceded`` upload that this pipeline used to read directly.

    ``entities`` is accepted for symmetry with the Astra UI form (which
    now mirrors the Sunrise multiselect with ``X - name`` free-form
    entries) but is not used by the current transformations.
    """
    del entities  # reserved for future skills

    # Drop excluded GoCs (matched on the 11-char ``goc`` prefix, so every
    # cohort year for that GoC is removed in one shot) BEFORE the
    # 16-year window filter — order doesn't change the result but it
    # keeps the trace clear when debugging.
    exclude = {g.strip() for g in (gocs_to_exclude or []) if g and g.strip()}
    filtered_pairs = [
        p for p in goc_cohort_pairs if p["goc"] not in exclude
    ]

    # Cap the Sunrise-produced pairs at the 16-year window around the
    # Astra analysis year.
    min_year = year - (ASTRA_COHORT_YEAR_SPAN - 1)
    pairs = [
        p for p in filtered_pairs if min_year <= p["year"] <= year
    ]

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

    # OCI placeholders disabled until the population rules are defined.
    # Re-enable by uncommenting the four lines below and adding
    # `closing_path` / `opening_path` back to the return list.
    # closing_path = run_dir / "OCI_OPTION_CF_CLOSING.csv"
    # opening_path = run_dir / "OCI_OPTION_CF_OPENING.csv"
    # create_empty_csv(str(closing_path))
    # create_empty_csv(str(opening_path))

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
        # closing_path,
        # opening_path,
    ]
