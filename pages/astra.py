"""Astra Input Builder page.

Prepares the Astra input workbooks. Section 1 mirrors the Sunrise page;
the rest of the process will be added once specified.
"""
from __future__ import annotations

import streamlit as st

st.set_page_config(
    page_title="Astra Input Builder",
    page_icon="🚀",
    layout="wide",
)

st.title("🚀 Astra Input Builder")
st.caption("Local pipeline that prepares Astra input files — data never leaves this machine.")

st.subheader("1. Upload input files")
uploaded = st.file_uploader(
    "Excel files",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
)

st.info(
    "The remaining sections (parameters, run, outputs) will be filled in "
    "once the Astra process is defined."
)
