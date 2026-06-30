"""Sunrise Input Builder page.

Prepares the Sunrise input workbooks from one or more ``_Ceded``/
``_Assumed`` files (each carrying an ``Input_Sunrise`` sheet), a
``Transcodifica_aggregazione_GOC_H_NH`` master list, and the
``Payment_Patterns_&_Risk_Adjustments`` workbook. All processing runs
locally.
"""
from __future__ import annotations

import os
import re
import shutil
from datetime import datetime
from pathlib import Path

import streamlit as st

from excel_pipeline.input_validation import (
    ValidationReport,
    validate_sunrise_inputs,
)
from excel_pipeline.pipeline import run_phase1
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

HAS_API_KEY = bool(os.getenv("ANTHROPIC_API_KEY"))

st.set_page_config(
    page_title="Sunrise Input Builder",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
st.session_state.setdefault("sunrise_run_id", None)
st.session_state.setdefault("chat_history", [])

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📊 Sunrise Input Builder")
st.caption("Local pipeline that prepares Sunrise input files — data never leaves this machine.")

col_main, col_chat = st.columns([2, 1])

with col_main:
    st.subheader("1. Upload input files")
    uploaded = st.file_uploader(
        "Excel or CSV files",
        type=["xlsx", "xlsm", "csv"],
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
            format_func=lambda s: "HY" if s == 1 else "FY",
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

    st.caption(
        "GOC not to be considered — type an 11-char GOC name "
        "(e.g. `IT05PABPPLE`) and press Enter to add. Every cohort "
        "year for the listed GoCs is removed from the run."
    )
    gocs_to_exclude_raw = st.multiselect(
        "GOC not to be considered",
        options=[],
        default=[],
        accept_new_options=True,
        label_visibility="collapsed",
        key="sunrise_gocs_to_exclude",
    )
    gocs_to_exclude = [
        str(g).strip() for g in gocs_to_exclude_raw if str(g).strip()
    ]

    run_clicked = st.button(
        "▶ Run pipeline",
        type="primary",
        disabled=not (uploaded and parsed_entities),
    )
    if run_clicked and uploaded and parsed_entities:
        run_id = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        run_dir = RUNS_DIR / run_id
        inputs_dir = run_dir / "inputs"
        inputs_dir.mkdir(parents=True)
        input_paths: list[Path] = []
        for f in uploaded:
            p = inputs_dir / f.name
            p.write_bytes(f.getvalue())
            input_paths.append(p)

        report = validate_sunrise_inputs(
            input_paths, year=int(year), semester=int(semester)
        )
        _render_validation_report(report)
        # Always persist the report next to the run, so the user can audit
        # both successful runs and blocked runs after the fact.
        report.save(run_dir / "validation.json")

        if report.is_blocking:
            # Errors already rendered above; nothing else to do.
            pass
        else:
            st.session_state.sunrise_run_id = run_id
            st.session_state.chat_history = []

            with st.status("Running pipeline...", expanded=True) as status:
                try:
                    status.update(label="Phase 1 in progress...")
                    phase1_result = run_phase1(
                        input_paths=input_paths,
                        run_dir=run_dir,
                        entities=parsed_entities,
                        year=int(year),
                        semester=int(semester),
                        gocs_to_exclude=gocs_to_exclude,
                    )
                    for out in phase1_result["outputs"]:
                        st.write(f"✅ Phase 1 → `{out.name}`")

                    # Expose to the Astra page via session_state:
                    # - (GoC, accident_year) pairs, replacing the legacy
                    #   AAI_P&C_Ceded upload that used to derive them;
                    # - the Health-perimeter GoCs read from the
                    #   Transcodifica file (column D == 'H'), replacing
                    #   the manual multiselect on the Astra page.
                    st.session_state.sunrise_goc_cohort_pairs = (
                        phase1_result["goc_cohort_pairs"]
                    )
                    st.session_state.sunrise_health_perimeter_gocs = (
                        phase1_result["health_perimeter_gocs"]
                    )

                    status.update(label="Pipeline complete", state="complete")
                except Exception as e:
                    status.update(label=f"Failed: {e}", state="error")
                    st.exception(e)

    # Show files produced in the current run
    if st.session_state.sunrise_run_id:
        st.subheader("3. Run outputs")
        run_dir = RUNS_DIR / st.session_state.sunrise_run_id
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
        saved_folder = load_pref("sunrise_output_folder", "")
        output_folder = st.text_input(
            "Output folder path (remembered between sessions)",
            value=saved_folder,
            placeholder=r"e.g. C:\Users\loren\Sunrise_outputs",
            key="sunrise_output_folder_input",
        )
        if output_folder != saved_folder:
            save_pref("sunrise_output_folder", output_folder)

        save_all = st.button(
            "📥 Save all Sunrise outputs to that folder",
            disabled=not (output_folder and files),
            key="sunrise_save_all",
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

with col_chat:
    st.subheader("💬 Ad-hoc edits")
    if not HAS_API_KEY:
        st.info(
            "Ad-hoc edits require an Anthropic API key. Add "
            "`ANTHROPIC_API_KEY` to `.env` to enable this panel. The main "
            "pipeline works without it."
        )
    elif st.session_state.sunrise_run_id is None:
        st.info("Run the pipeline first, then ask for changes here.")
    else:
        from excel_pipeline.orchestrator import run_ad_hoc

        st.caption(f"Editing run `{st.session_state.sunrise_run_id}`")

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
                            run_dir=RUNS_DIR / st.session_state.sunrise_run_id,
                        )
                    except Exception as e:
                        response = f"Error: {e}"
                    st.write(response)
                    st.session_state.chat_history.append(
                        {"role": "assistant", "content": response}
                    )
            st.rerun()
