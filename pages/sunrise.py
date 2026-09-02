"""Sunrise Input Builder page.

Prepares the Sunrise input workbooks from one or more ``_Ceded``/
``_Assumed`` files (each carrying an ``Input_Sunrise`` sheet), a
``Transcodifica_aggregazione_GOC_H_NH`` master list, and the
``Payment_Patterns_&_Risk_Adjustments`` workbook. All processing runs
locally.
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

st.set_page_config(
    page_title="Sunrise Input Builder",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
st.session_state.setdefault("sunrise_run_id", None)

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📊 Sunrise Input Builder")
st.caption("Local pipeline that prepares Sunrise input files — data never leaves this machine.")

st.warning(
    "**Warning!** The following outputs must be provided or modified by "
    "the user:\n\n"
    "- `IFRs17Rates.csv`",
    icon="⚠️",
)

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
    "**Entity to analyze** — la prima colonna è l'etichetta "
    "dell'entity nel formato `ID - nome` (es. `6 - AAI`). La "
    "seconda colonna è il suffisso usato per identificare gli sheet "
    "Risk_Adjustment / Payment_pattern nel workbook "
    "`Payment_Patterns_&_Risk_Adjustments`: cerca "
    "`ra_<suffisso>_REINS` e `pp_<suffisso>_REINS`. Puoi aggiungere "
    "righe con il `+` in fondo alla tabella; il run usa la riga "
    "selezionata sotto."
)
entities_df = st.data_editor(
    pd.DataFrame(
        [
            {"Entity": "6 - AAI", "Sheet suffix": "AAI"},
            {"Entity": "14 - MPS", "Sheet suffix": "AMAD"},
        ]
    ),
    num_rows="dynamic",
    hide_index=True,
    column_config={
        "Entity": st.column_config.TextColumn(
            "Entity",
            required=True,
            help="Formato: `ID - nome`, es. `6 - AAI`.",
        ),
        "Sheet suffix": st.column_config.TextColumn(
            "Sheet suffix",
            required=True,
            help=(
                "Suffisso degli sheet dentro il file "
                "Payment_Patterns_&_Risk_Adjustments "
                "(`ra_<suffisso>_REINS`, `pp_<suffisso>_REINS`)."
            ),
        ),
    },
    key="sunrise_entities_editor",
)
_entity_rows: list[tuple[int, str, str]] = []
_invalid_entries: list[str] = []
for _, _row in entities_df.iterrows():
    _ent_raw = str(_row.get("Entity") or "").strip()
    _suf = str(_row.get("Sheet suffix") or "").strip()
    if not _ent_raw and not _suf:
        continue
    _match = _ENTITY_FREEFORM_RE.match(_ent_raw) if _ent_raw else None
    if _match is None or not _suf:
        _invalid_entries.append(_ent_raw or "(vuoto)")
        continue
    _entity_rows.append(
        (int(_match.group(1)), _match.group(2).strip(), _suf)
    )
if _invalid_entries:
    st.warning(
        "Righe della tabella entities ignorate — la prima colonna deve "
        "essere nel formato `ID - nome` e la seconda colonna non vuota. "
        f"Ignorate: {_invalid_entries}",
        icon="⚠️",
    )

parsed_entities: list[tuple[int, str]] = []
sheet_suffix: str = ""
if _entity_rows:
    _labels = [
        f"{eid} - {ename}  →  sheet suffix: {suf}"
        for eid, ename, suf in _entity_rows
    ]
    _picked_idx = st.radio(
        "Riga usata per il run",
        options=range(len(_entity_rows)),
        format_func=lambda i: _labels[i],
        key="sunrise_entity_pick",
    )
    _eid, _ename, sheet_suffix = _entity_rows[_picked_idx]
    parsed_entities = [(_eid, _ename)]

st.caption(
    "**GOC not to be considered** — type an 11-char GOC name "
    "(e.g. `IT05PABPPLE`) and press Enter to add. Every "
    "cohort year for the listed GoCs is removed from the run."
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

st.caption(
    "**GOC to rename** — indica direttamente il binomio `GOC+coorte` "
    "(es. `IT05PABPPLE2024`) sia in *Old* sia in *New*. Il rename opera "
    "solo sulla specifica coppia (GoC, coorte), non su tutte le coorti "
    "della stessa GoC. Il *New* NON può coincidere con un `GOC+coorte` "
    "già presente negli input: in tal caso il run si ferma con errore. "
    "Formato: nome GoC seguito dai 4 caratteri dell'anno di coorte, "
    "senza separatori."
)
rename_df = st.data_editor(
    pd.DataFrame([{"Old GOC+cohort": "", "New GOC+cohort": ""}]),
    num_rows="dynamic",
    hide_index=True,
    column_config={
        "Old GOC+cohort": st.column_config.TextColumn(
            "Old GOC+cohort",
            help="GOC+coorte come compare negli input (es. `IT05PABPPLE2024`).",
        ),
        "New GOC+cohort": st.column_config.TextColumn(
            "New GOC+cohort",
            help="Nuovo GOC+coorte. Non deve già esistere nei dati.",
        ),
    },
    key="sunrise_goc_renames_editor",
)
goc_renames: dict[str, str] = {}
for _, row in rename_df.iterrows():
    old = str(row.get("Old GOC+cohort") or "").strip()
    new = str(row.get("New GOC+cohort") or "").strip()
    if old and new:
        goc_renames[old] = new

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
        input_paths,
        year=int(year),
        semester=int(semester),
        sheet_suffix=sheet_suffix,
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
        try:
            with st.spinner("Running pipeline..."):
                phase1_result = run_phase1(
                    input_paths=input_paths,
                    run_dir=run_dir,
                    entities=parsed_entities,
                    year=int(year),
                    semester=int(semester),
                    sheet_suffix=sheet_suffix,
                    gocs_to_exclude=gocs_to_exclude,
                    goc_renames=goc_renames,
                )

            # Warnings surface as top-level yellow banners so they can't
            # get hidden inside a collapsed status box.
            for w in phase1_result.get("warnings", []):
                st.warning(w, icon="⚠️")

            st.success(
                f"Pipeline complete — "
                f"{len(phase1_result['outputs'])} file(s) prodotti.",
                icon="✅",
            )

            # Expose to the Astra page via session_state:
            # - (GoC, accident_year) pairs, replacing the legacy
            #   AAI_P&C_Ceded upload that used to derive them;
            # - the Health-perimeter GoCs read from the Transcodifica
            #   file (column D == 'H'), replacing the manual
            #   multiselect on the Astra page.
            st.session_state.sunrise_goc_cohort_pairs = (
                phase1_result["goc_cohort_pairs"]
            )
            st.session_state.sunrise_health_perimeter_gocs = (
                phase1_result["health_perimeter_gocs"]
            )
        except Exception as e:
            st.error(f"❌ Pipeline failed: {e}")
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
