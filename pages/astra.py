"""Astra Input Builder page.

Prepares the Astra input workbooks. Reuses the GoC list and the
analysis year produced by the latest Sunrise run, so no duplicate
input form is needed here.
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
    "the output is not produced yet.\n\n"
    "`NEW_BUSINESS_PPOS.xlsx` is fully wired and uses the GoC list + "
    "analysis year from the last Sunrise run.",
    icon="🚧",
)

st.subheader("1. Upload input files")
uploaded = st.file_uploader(
    "Excel files",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
)

# Pull the GoC list and analysis year from session state (Sunrise sets them).
goc_names: list[str] | None = st.session_state.get("sunrise_goc_names")
year: int | None = st.session_state.get("sunrise_year")

st.subheader("2. Parameters from Sunrise")
sunrise_ready = bool(goc_names) and year is not None
if sunrise_ready:
    st.success(
        f"Using **{len(goc_names)} GoC(s)** and analysis year **{year}** "
        f"from the last Sunrise run."
    )
    with st.expander("Show GoC list"):
        st.write(goc_names)
else:
    st.error(
        "No Sunrise run found in this session. Open the **Sunrise** page, "
        "run the pipeline, then come back here."
    )

run_clicked = st.button(
    "▶ Run pipeline",
    type="primary",
    disabled=not sunrise_ready,
)
if run_clicked and sunrise_ready:
    run_id = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    run_dir = RUNS_DIR / run_id
    inputs_dir = run_dir / "inputs"
    inputs_dir.mkdir(parents=True)
    input_paths: list[Path] = []
    for f in uploaded or []:
        p = inputs_dir / f.name
        p.write_bytes(f.getvalue())
        input_paths.append(p)
    st.session_state.astra_run_id = run_id

    with st.status("Running Astra pipeline...", expanded=True) as status:
        try:
            outputs = run_astra_phase1(
                input_paths=input_paths,
                run_dir=run_dir,
                goc_names=goc_names,
                year=int(year),
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
