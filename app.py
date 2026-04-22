"""Excel Pipeline — Streamlit entry point.

A local web UI for running the Excel transformation pipeline and making
ad-hoc edits afterwards. All files stay on this machine.
"""
from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

from excel_pipeline.pipeline import run_phase1
from excel_pipeline.skill import list_run_files

load_dotenv()

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
ROOT = Path(__file__).parent
RUNS_DIR = ROOT / "runs"
RUNS_DIR.mkdir(exist_ok=True)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

ENTITIES: list[tuple[int, str]] = [
    (14, "MPS"),
    (6, "AAI"),
    (11, "DIRECT ITALY"),
    (19, "NOBIS"),
]

HAS_API_KEY = bool(os.getenv("ANTHROPIC_API_KEY"))

st.set_page_config(
    page_title="Excel Pipeline",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
st.session_state.setdefault("run_id", None)
st.session_state.setdefault("chat_history", [])

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📊 Excel Pipeline")
st.caption("Local Excel pipeline — files never leave this machine.")

col_main, col_chat = st.columns([2, 1])

with col_main:
    st.subheader("1. Upload input files")
    uploaded = st.file_uploader(
        "Excel files",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
    )

    st.subheader("2. Analysis parameters")
    col_year, col_sem = st.columns(2)
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
    entity = st.selectbox(
        "Entity to analyze",
        options=ENTITIES,
        format_func=lambda e: f"{e[0]} — {e[1]}",
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
        st.session_state.run_id = run_id
        st.session_state.chat_history = []

        with st.status("Running pipeline...", expanded=True) as status:
            try:
                status.update(label="Phase 1 in progress...")
                phase1_out = run_phase1(
                    input_paths=input_paths,
                    run_dir=run_dir,
                    entity_id=entity_id,
                    entity_name=entity_name,
                    year=int(year),
                    semester=int(semester),
                )
                st.write(f"✅ Phase 1 → `{phase1_out.name}`")

                status.update(label="Pipeline complete", state="complete")
            except Exception as e:
                status.update(label=f"Failed: {e}", state="error")
                st.exception(e)

    # Show files produced in the current run
    if st.session_state.run_id:
        st.subheader("3. Run outputs")
        run_dir = RUNS_DIR / st.session_state.run_id
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

with col_chat:
    st.subheader("💬 Ad-hoc edits")
    if not HAS_API_KEY:
        st.info(
            "Ad-hoc edits require an Anthropic API key. Add "
            "`ANTHROPIC_API_KEY` to `.env` to enable this panel. The main "
            "pipeline works without it."
        )
    elif st.session_state.run_id is None:
        st.info("Run the pipeline first, then ask for changes here.")
    else:
        from excel_pipeline.orchestrator import run_ad_hoc

        st.caption(f"Editing run `{st.session_state.run_id}`")

        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])

        if prompt := st.chat_input("Describe a change..."):
            st.session_state.chat_history.append(
                {"role": "user", "content": prompt}
            )
            with st.chat_message("user"):
                st.write(prompt)

            with st.chat_message("assistant"):
                with st.spinner("Working..."):
                    try:
                        response = run_ad_hoc(
                            request=prompt,
                            run_dir=RUNS_DIR / st.session_state.run_id,
                        )
                    except Exception as e:
                        response = f"Error: {e}"
                    st.write(response)
                    st.session_state.chat_history.append(
                        {"role": "assistant", "content": response}
                    )
            st.rerun()
