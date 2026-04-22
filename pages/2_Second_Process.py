"""Second Sunrise process — placeholder.

Add a page per distinct input-building process. Streamlit auto-discovers
every `.py` file in `pages/` and shows it in the sidebar, using the
filename (with leading `N_` stripped, underscores turned into spaces)
as the label. Rename this file when the process is defined.
"""
from __future__ import annotations

import streamlit as st

st.set_page_config(page_title="Second Process", page_icon="🛠️", layout="wide")

st.title("🛠️ Second Process")
st.caption("Placeholder — to be filled once the process is specified.")

st.info(
    "This page is a placeholder. Describe the new Sunrise-input-building "
    "process and its inputs/outputs, and this page will be replaced with "
    "a proper Section 1/2/3 layout like the main one."
)
