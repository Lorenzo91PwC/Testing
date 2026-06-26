"""Astra Input Builder page.

Prepares the Astra input workbooks. Independent of the Sunrise page —
the user uploads the Ceded file and sets parameters here, and the Astra
pipeline runs on those.
"""
from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from excel_pipeline.input_validation import (
    ValidationReport,
    validate_astra_inputs,
)
from excel_pipeline.pipeline import run_astra_phase1
from excel_pipeline.skill import list_run_files
from excel_pipeline.user_prefs import load_pref, save_pref


def _render_validation_report(report: ValidationReport) -> None:
    """Show errors with ``st.error`` and warnings with ``st.warning``."""
    if report.errors:
        bullets = "\n".join(
            f"- **{i.file}** — {i.location} — {i.message}"
            for i in report.errors
        )
        st.error(
            "❌ **Cannot run the pipeline — input validation failed:**\n\n"
            f"{bullets}\n\n"
            "Fix the file(s) and click Run again.",
            icon="🚫",
        )
    if report.warnings:
        bullets = "\n".join(
            f"- **{i.file}** — {i.location} — {i.message}"
            for i in report.warnings
        )
        st.warning(
            "⚠️ **Input validation warnings** (run continues):\n\n"
            f"{bullets}",
            icon="⚠️",
        )

_ENTITY_FREEFORM_RE = re.compile(r"^\s*(\d+)\s*-\s*(.+?)\s*$")


def _parse_entity_selection(item: object) -> tuple[int, str] | None:
    """Normalize a multiselect entry to a ``(id, name)`` tuple.

    Accepts:
    - the original ``(id, name)`` tuple from the preset list,
    - a free-form string the user typed in the format ``"X - name"`` where
      ``X`` is a positive integer and ``name`` is non-empty.

    Returns ``None`` for anything that does not match.
    """
    if isinstance(item, tuple) and len(item) == 2:
        return item
    if isinstance(item, str):
        m = _ENTITY_FREEFORM_RE.match(item)
        if m:
            return (int(m.group(1)), m.group(2).strip())
    return None

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
ROOT = Path(__file__).parent.parent
RUNS_DIR = ROOT / "runs"
RUNS_DIR.mkdir(exist_ok=True)

CSV_MIME = "text/csv"

ENTITIES: list[tuple[int, str]] = [
    (14, "MPS"),
    (6, "AAI"),
    (11, "DIRECT ITALY"),
    (19, "NOBIS"),
]

ASTRA_DEFAULT_AOM_IMPACT_PAIRS: list[tuple[str, int]] = [
    ("DA_LIC_OP", 0),
    ("DA_LIC_INCLAIM_INCEXP", 0),
    ("DA_LIC_CHG", 0),
    ("DA_LIC_CLO", 0),
]

st.set_page_config(
    page_title="Astra Input Builder",
    page_icon="🚀",
    layout="wide",
)

st.session_state.setdefault("astra_run_id", None)

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("🚀 Astra Input Builder")
st.caption("Local pipeline that prepares Astra input files — data never leaves this machine.")

st.warning(
    "⚠️ **Pending specifications** — the following outputs still need their "
    "population rules:\n\n"
    "- `OCI_OPTION_CF_CLOSING.csv` and `OCI_OPTION_CF_OPENING.csv` — "
    "rules not yet defined; the placeholder emission is currently "
    "disabled in the pipeline.\n"
    "- `INCEPTION_FORWARD_RATES.csv` — the source input file (suffix and "
    "sheet) and the specific rates row to extract are still to be defined; "
    "the output is not produced yet.\n\n"
    "`MP_GOC_SEG.csv` is wired and rewrites `P&C` to `HLTH_PC` in "
    "columns A and C for the GoCs whose H-NH flag is `H` in the "
    "Transcodifica file uploaded on the Sunrise page. "
    "`MP_GOC.csv` is wired and rewrites columns E, F, L, P based on "
    "year, semester and business type.",
    icon="🚧",
)

st.subheader("1. Upload input files")
uploaded = st.file_uploader(
    "Excel or CSV files",
    type=["xlsx", "xlsm", "csv"],
    accept_multiple_files=True,
)

st.subheader("2. Analysis parameters")
col_year, col_sem, col_type = st.columns(3)
with col_year:
    year = st.number_input(
        "Year",
        min_value=2000,
        max_value=2100,
        value=datetime.now().year,
        step=1,
    )
with col_sem:
    semester = st.radio(
        "Semester",
        options=[1, 2],
        horizontal=True,
        format_func=lambda s: "HY" if s == 1 else "FY",
    )
with col_type:
    business_type = st.radio(
        "Business type",
        options=["Diretto", "Ceduto"],
        horizontal=True,
    )
st.caption(
    "Entities to analyze — pick from the list or type a new one in "
    "the format `X - name` (e.g. `99 - CUSTOM`) and press Enter."
)
raw_entity_selection = st.multiselect(
    "Entities to analyze",
    options=ENTITIES,
    default=[ENTITIES[1]],
    format_func=lambda e: f"{e[0]} — {e[1]}",
    accept_new_options=True,
    label_visibility="collapsed",
)
parsed_entities: list[tuple[int, str]] = []
invalid_entries: list[str] = []
for item in raw_entity_selection:
    parsed = _parse_entity_selection(item)
    if parsed is None:
        invalid_entries.append(str(item))
    else:
        parsed_entities.append(parsed)
if invalid_entries:
    st.warning(
        "These entries don't match the `X - name` format and were "
        f"ignored: {invalid_entries}",
        icon="⚠️",
    )

col_close_curve, col_open_curve = st.columns(2)
with col_close_curve:
    closing_curve_name = st.text_input(
        "Closing curve name",
        value="20251231_EUR_LP100_FY25",
        help="Goes into column C of CURVE_ID_PARAM.csv for rows where "
        "VARIABLE_NAME == 'CLOSING_CURVE_ID'.",
    )
with col_open_curve:
    opening_curve_name = st.text_input(
        "Opening curve name",
        value="20241231_EUR_LP100_FY24",
        help="Goes into column C of CURVE_ID_PARAM.csv for rows where "
        "VARIABLE_NAME == 'OPENING_CURVE_ID'.",
    )

health_perimeter_gocs = list(
    st.session_state.get("sunrise_health_perimeter_gocs") or []
)
if health_perimeter_gocs:
    with st.expander(
        f"🏥 Health perimeter — {len(health_perimeter_gocs)} GoC(s) from Transcodifica",
        expanded=False,
    ):
        st.write(health_perimeter_gocs)
else:
    st.warning(
        "⚠️ Health perimeter is empty — either Sunrise has not been run "
        "yet in this session, or the Transcodifica file does not contain "
        "any row with column D (H-NH) set to **H**. MP_GOC_SEG will be "
        "produced without any P&C → HLTH_PC substitution.",
        icon="🏥",
    )

st.info(
    "💡 **AOM Impact rows** — for every (GoC, year) found in the Ceded "
    "file, the rows below are appended to `ACTUARIAL_AOM_IMPACT.csv`. "
    "Use the `+` button at the bottom of the table to add a row, or "
    "the row's checkbox + Delete key to remove it.",
    icon="ℹ️",
)
aom_impact_df = st.data_editor(
    pd.DataFrame(
        [{"STEP_ID": s, "Value": v} for s, v in ASTRA_DEFAULT_AOM_IMPACT_PAIRS]
    ),
    num_rows="dynamic",
    hide_index=True,
    column_config={
        "STEP_ID": st.column_config.TextColumn(
            "STEP_ID",
            required=True,
            help="Goes into column B of ACTUARIAL_AOM_IMPACT.csv.",
        ),
        "Value": st.column_config.NumberColumn(
            "Value",
            required=True,
            format="%.10f",
            step=1e-10,
            help="Goes into column C of ACTUARIAL_AOM_IMPACT.csv. "
            "Supports up to 10 decimal places.",
        ),
    },
    key="astra_aom_impact_editor",
)
aom_impact_pairs: list[tuple[str, object]] = []
for _, row in aom_impact_df.iterrows():
    step = row.get("STEP_ID")
    val = row.get("Value")
    if step is None or pd.isna(step):
        continue
    step_str = str(step).strip()
    if not step_str:
        continue
    if val is None or pd.isna(val):
        continue
    aom_impact_pairs.append((step_str, val))

# The (GoC, accident_year) pairs come from the Sunrise run via
# session_state — replacing the legacy AAI_P&C_Ceded upload that used
# to drive Astra. If Sunrise has not been run yet in this session,
# block the pipeline with a clear message.
sunrise_pairs = st.session_state.get("sunrise_goc_cohort_pairs")
sunrise_ready = bool(sunrise_pairs)
if not sunrise_ready:
    st.error(
        "❌ Astra needs the GoC list from the Sunrise run. Open the "
        "**Sunrise** page, click Run there, then come back.",
        icon="🚫",
    )

run_clicked = st.button(
    "▶ Run pipeline",
    type="primary",
    disabled=not (uploaded and parsed_entities and sunrise_ready),
)
if run_clicked and uploaded and parsed_entities and sunrise_ready:
    run_id = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    run_dir = RUNS_DIR / run_id
    inputs_dir = run_dir / "inputs"
    inputs_dir.mkdir(parents=True)
    input_paths: list[Path] = []
    for f in uploaded:
        p = inputs_dir / f.name
        p.write_bytes(f.getvalue())
        input_paths.append(p)

    report = validate_astra_inputs(input_paths)
    _render_validation_report(report)
    report.save(run_dir / "validation.json")

    if report.is_blocking:
        # Errors already rendered above; do not start the pipeline.
        pass
    else:
        st.session_state.astra_run_id = run_id
        with st.status("Running Astra pipeline...", expanded=True) as status:
            try:
                outputs = run_astra_phase1(
                    input_paths=input_paths,
                    run_dir=run_dir,
                    entities=parsed_entities,
                    year=int(year),
                    semester=int(semester),
                    business_type=business_type,
                    health_perimeter_gocs=health_perimeter_gocs,
                    actuarial_aom_impact_pairs=aom_impact_pairs,
                    closing_curve_name=closing_curve_name.strip(),
                    opening_curve_name=opening_curve_name.strip(),
                    goc_cohort_pairs=sunrise_pairs,
                )
                for out in outputs:
                    st.write(f"✅ → `{out.name}`")
                status.update(label="Pipeline complete", state="complete")
            except Exception as e:
                status.update(label=f"Failed: {e}", state="error")
                st.exception(e)

# Show files produced in the current run
if st.session_state.astra_run_id:
    st.subheader("3. Run outputs")
    run_dir = RUNS_DIR / st.session_state.astra_run_id
    files = list_run_files(run_dir)
    if not files:
        st.info("No output files yet.")
    for f in files:
        st.download_button(
            label=f"⬇ {f.name}",
            data=f.read_bytes(),
            file_name=f.name,
            mime=CSV_MIME,
            key=str(f),
        )

    st.subheader("4. Save all outputs to folder")
    saved_folder = load_pref("astra_output_folder", "")
    output_folder = st.text_input(
        "Output folder path (remembered between sessions)",
        value=saved_folder,
        placeholder=r"e.g. C:\Users\loren\Astra_outputs",
        key="astra_output_folder_input",
    )
    if output_folder != saved_folder:
        save_pref("astra_output_folder", output_folder)

    save_all = st.button(
        "📥 Save all Astra outputs to that folder",
        disabled=not (output_folder and files),
        key="astra_save_all",
    )
    if save_all and output_folder and files:
        try:
            dest = Path(output_folder).expanduser()
            dest.mkdir(parents=True, exist_ok=True)
            copied: list[str] = []
            for f in files:
                shutil.copy2(f, dest / f.name)
                copied.append(f.name)
            st.success(
                f"✅ Saved {len(copied)} file(s) to `{dest}`."
            )
        except Exception as e:  # noqa: BLE001
            st.error(f"❌ Failed to save outputs: {e}")
