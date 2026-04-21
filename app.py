"""Excel Pipeline — Streamlit entry point.

A local web UI for running the Excel transformation pipeline and making
ad-hoc edits afterwards. All files stay on this machine.
"""
from __future__ import annotations

import json
import os
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from excel_pipeline.orchestrator import run_ad_hoc
from excel_pipeline.skill import (
    find_file_by_suffix,
    get_unique_column_values,
    list_run_files,
)

CEDED_SUFFIX = "AAI_P&C_Ceded"
CEDED_SHEET = "AAI_P&C_Ceded"
CEDED_COLUMN = "G"

load_dotenv()

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
ROOT = Path(__file__).parent
RUNS_DIR = ROOT / "runs"
RUNS_DIR.mkdir(exist_ok=True)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

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
    st.subheader("1. Analysis date")
    analysis_date = st.date_input(
        "📅 Date of analysis",
        value=st.session_state.get("analysis_date", date.today()),
        key="analysis_date",
        help="Reference date used across this pipeline run.",
    )
    st.caption(f"Selected: **{analysis_date.isoformat()}**")

    st.divider()

    st.subheader("2. Upload input file - Sunrise input")
    uploaded = st.file_uploader(
        "Excel files",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help="You can upload multiple files — drag several at once or use the browse button.",
    )
    if uploaded:
        st.caption(f"{len(uploaded)} file(s) ready: " + ", ".join(f.name for f in uploaded))
        if st.button("💾 Save uploaded files"):
            run_id = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            run_dir = RUNS_DIR / run_id
            run_dir.mkdir(parents=True)
            for f in uploaded:
                (run_dir / f.name).write_bytes(f.getvalue())
            st.session_state.run_id = run_id
            st.session_state.chat_history = []
            st.success(
                f"Saved {len(uploaded)} file(s) to `runs/{run_id}/`."
            )

    # Browse & preview previously uploaded input files
    st.subheader("3. Browse & preview input files")
    input_files = sorted(
        (
            f
            for f in [*RUNS_DIR.glob("*/*.xlsx"), *RUNS_DIR.glob("*/*.xlsm")]
            if "_output" not in f.stem
        ),
        reverse=True,
    )
    if not input_files:
        st.info("No uploaded files yet. Upload one above to see it here.")
    else:
        selected = st.selectbox(
            "Pick an uploaded file",
            input_files,
            format_func=lambda p: p.name,
            key="preview_file",
        )
        try:
            xl = pd.ExcelFile(selected)
            sheet = st.selectbox("Sheet", xl.sheet_names, key="preview_sheet")
            df = xl.parse(sheet)
            st.caption(f"{len(df)} rows × {len(df.columns)} columns")
            st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.warning(f"Could not read `{selected.name}`: {e}")

    st.divider()

    # Run deterministic elaborations on the uploaded files
    st.subheader("4. Run elaborations")
    if not uploaded and st.session_state.run_id is None:
        st.info("Upload files and save them above to enable elaborations.")
    elif st.button("▶ Run elaborations", type="primary"):
        if st.session_state.run_id is None:
            run_id = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            run_dir = RUNS_DIR / run_id
            run_dir.mkdir(parents=True)
            for f in uploaded:
                (run_dir / f.name).write_bytes(f.getvalue())
            st.session_state.run_id = run_id
            st.session_state.chat_history = []
        else:
            run_dir = RUNS_DIR / st.session_state.run_id

        state: dict = {"analysis_date": analysis_date.isoformat()}

        with st.status("Running elaborations...", expanded=True) as status:
            try:
                status.update(label=f"Locating {CEDED_SUFFIX} file...")
                ceded_path = find_file_by_suffix(str(run_dir), CEDED_SUFFIX)
                if ceded_path is None:
                    raise FileNotFoundError(
                        f"No file ending in '{CEDED_SUFFIX}' found in the upload."
                    )
                st.write(f"✅ Found `{Path(ceded_path).name}`")

                status.update(label="Extracting unique Ceded portfolio codes...")
                codes = get_unique_column_values(
                    path=ceded_path,
                    sheet=CEDED_SHEET,
                    column=CEDED_COLUMN,
                )
                state["ceded_portfolio_codes"] = codes
                st.write(f"✅ {len(codes)} unique values in {CEDED_SHEET}!{CEDED_COLUMN}")

                (run_dir / "state.json").write_text(
                    json.dumps(state, indent=2, default=str), encoding="utf-8"
                )
                status.update(label="Elaborations complete", state="complete")
            except Exception as e:
                status.update(label=f"Failed: {e}", state="error")
                st.exception(e)

    # Show results and files produced in the current run
    if st.session_state.run_id:
        st.subheader("5. Run outputs")
        run_dir = RUNS_DIR / st.session_state.run_id
        state_path = run_dir / "state.json"
        if state_path.exists():
            state_data = json.loads(state_path.read_text(encoding="utf-8"))
            codes = state_data.get("ceded_portfolio_codes", [])
            if codes:
                with st.expander(
                    f"Ceded portfolio codes — {len(codes)} unique values",
                    expanded=True,
                ):
                    st.dataframe(
                        pd.DataFrame({"code": codes}),
                        use_container_width=True,
                        height=300,
                    )

        files = list_run_files(run_dir)
        if not files:
            st.info("No output Excel files yet.")
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
