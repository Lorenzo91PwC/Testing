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

from excel_pipeline.orchestrator import run_ad_hoc, run_phase
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
# Guard: API key
# ---------------------------------------------------------------------------
if not os.getenv("ANTHROPIC_API_KEY"):
    st.error(
        "ANTHROPIC_API_KEY not found. Copy `.env.example` to `.env` "
        "and add your key from https://console.anthropic.com"
    )
    st.stop()

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📊 Excel Pipeline")
st.caption("Local Excel pipeline — files never leave this machine.")

col_main, col_chat = st.columns([2, 1])

with col_main:
    st.subheader("1. Upload input file")
    uploaded = st.file_uploader("Excel file", type=["xlsx", "xlsm"])
    entity = st.selectbox(
        "Entity to analyze",
        options=ENTITIES,
        format_func=lambda e: f"{e[0]} — {e[1]}",
    )

    if uploaded and entity and st.button("▶ Run pipeline", type="primary"):
        entity_id, entity_name = entity
        run_id = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        run_dir = RUNS_DIR / run_id
        run_dir.mkdir(parents=True)
        input_path = run_dir / "input.xlsx"
        input_path.write_bytes(uploaded.getvalue())
        st.session_state.run_id = run_id
        st.session_state.chat_history = []

        with st.status("Running pipeline...", expanded=True) as status:
            try:
                status.update(label="Phase 1 in progress...")
                phase1_out = run_phase(
                    phase="phase1",
                    input_path=input_path,
                    run_dir=run_dir,
                    entity_id=entity_id,
                    entity_name=entity_name,
                )
                st.write(f"✅ Phase 1 → `{phase1_out.name}`")

                status.update(label="Phase 2 in progress...")
                phase2_out = run_phase(
                    phase="phase2",
                    input_path=phase1_out,
                    run_dir=run_dir,
                    entity_id=entity_id,
                    entity_name=entity_name,
                )
                st.write(f"✅ Phase 2 → `{phase2_out.name}`")

                status.update(label="Pipeline complete", state="complete")
            except Exception as e:
                status.update(label=f"Failed: {e}", state="error")
                st.exception(e)

    # Show files produced in the current run
    if st.session_state.run_id:
        st.subheader("2. Run outputs")
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
    if st.session_state.run_id is None:
        st.info("Run the pipeline first, then ask for changes here.")
    else:
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
