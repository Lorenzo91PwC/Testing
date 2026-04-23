"""Astra Input Builder page.

Prepares the Astra input workbooks. Currently placeholder outputs only
(empty files) — population rules are not yet defined.
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
    "⚠️ **Reminder** — the population rules for `OCI_OPTION_CF_CLOSING.xlsx` "
    "and `OCI_OPTION_CF_OPENING.xlsx` are not yet defined. Running the "
    "pipeline produces empty placeholder files. Integrate the real rules "
    "as soon as the specification is available.",
    icon="🚧",
)

st.subheader("1. Upload input files")
uploaded = st.file_uploader(
    "Excel files",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
)

run_clicked = st.button("▶ Run pipeline", type="primary")
if run_clicked:
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
            outputs = run_astra_phase1(input_paths=input_paths, run_dir=run_dir)
            for out in outputs:
                st.write(f"✅ → `{out.name}` (empty placeholder)")
            status.update(label="Pipeline complete", state="complete")
        except Exception as e:
            status.update(label=f"Failed: {e}", state="error")
            st.exception(e)

# Show files produced in the current run
if st.session_state.astra_run_id:
    st.subheader("2. Run outputs")
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
