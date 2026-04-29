"""Astra Input Builder page.

Prepares the Astra input workbooks. Independent of the Sunrise page —
the user uploads the Ceded file and sets parameters here, and the Astra
pipeline runs on those.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from excel_pipeline.pipeline import run_astra_phase1
from excel_pipeline.skill import list_run_files

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

ASTRA_DEFAULT_GOC_SEG_LIST: list[str] = [
    "IT05RRIEEBB",
    "IT05RRIHEBB",
    "IT05RRIHMAF",
    "IT05RRIHVIC",
    "IT05RRIHAST",
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
    "rules not yet defined; currently produced as empty placeholders.\n"
    "- `INCEPTION_FORWARD_RATES.csv` — the source input file (suffix and "
    "sheet) and the specific rates row to extract are still to be defined; "
    "the output is not produced yet.\n\n"
    "`MP_GOC_SEG.csv` is wired and rewrites `P&C` to `HLTH_PC` in "
    "columns A and C for the GoCs in the Health perimeter list below. "
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
entity = st.selectbox(
    "Entity to analyze",
    options=ENTITIES,
    format_func=lambda e: f"{e[0]} — {e[1]}",
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

st.info(
    "💡 **Health perimeter GoC** — five default codes are pre-loaded. "
    "**Remove** a code by clicking the × on its chip. **Add** a new "
    "code by typing it in the box and pressing Enter.",
    icon="ℹ️",
)
goc_seg_list = st.multiselect(
    "Health perimeter GoC",
    options=ASTRA_DEFAULT_GOC_SEG_LIST,
    default=ASTRA_DEFAULT_GOC_SEG_LIST,
    accept_new_options=True,
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

run_clicked = st.button(
    "▶ Run pipeline",
    type="primary",
    disabled=not (uploaded and entity),
)
if run_clicked and uploaded and entity:
    entity_id, entity_name = entity
    run_id = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    run_dir = RUNS_DIR / run_id
    inputs_dir = run_dir / "inputs"
    inputs_dir.mkdir(parents=True)
    input_paths: list[Path] = []
    for f in uploaded:
        p = inputs_dir / f.name
        p.write_bytes(f.getvalue())
        input_paths.append(p)
    st.session_state.astra_run_id = run_id

    with st.status("Running Astra pipeline...", expanded=True) as status:
        try:
            outputs = run_astra_phase1(
                input_paths=input_paths,
                run_dir=run_dir,
                entity_id=entity_id,
                entity_name=entity_name,
                year=int(year),
                semester=int(semester),
                business_type=business_type,
                health_perimeter_gocs=list(goc_seg_list),
                actuarial_aom_impact_pairs=aom_impact_pairs,
                closing_curve_name=closing_curve_name.strip(),
                opening_curve_name=opening_curve_name.strip(),
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
