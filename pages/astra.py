"""Astra Input Builder page.

Prepares the Astra input workbooks. Independent of the Sunrise page —
the user uploads the Ceded file and sets parameters here, and the Astra
pipeline runs on those.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import streamlit as st

from excel_pipeline.pipeline import run_astra_phase1
from excel_pipeline.skill import list_run_files

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
ROOT = Path(__file__).parent.parent
RUNS_DIR = ROOT / "runs"
RUNS_DIR.mkdir(exist_ok=True)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

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
    "- `OCI_OPTION_CF_CLOSING.xlsx` and `OCI_OPTION_CF_OPENING.xlsx` — "
    "rules not yet defined; currently produced as empty placeholders.\n"
    "- `INCEPTION_FORWARD_RATES.xlsx` — the source input file (suffix and "
    "sheet) and the specific rates row to extract are still to be defined; "
    "the output is not produced yet.",
    icon="🚧",
)

st.subheader("1. Upload input files")
uploaded = st.file_uploader(
    "Excel files",
    type=["xlsx", "xlsm"],
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
        format_func=lambda s: f"H{s}",
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

goc_seg_list = st.multiselect(
    "GoC list for MP_GOC_SEG",
    options=ASTRA_DEFAULT_GOC_SEG_LIST,
    default=ASTRA_DEFAULT_GOC_SEG_LIST,
    accept_new_options=True,
    help=(
        "Used to update MP_GOC_SEG (skill not yet wired). The five "
        "defaults are pre-loaded; remove with the × on each chip or "
        "type a new code and press Enter to add it."
    ),
)

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
            mime=XLSX_MIME,
            key=str(f),
        )
