"""Multi-page entry point.

Wires the sidebar navigation between the two input-building processes.
Each page lives under `pages/`. Nothing happens here beyond routing.
"""
from __future__ import annotations

import streamlit as st
from dotenv import load_dotenv

load_dotenv()

pg = st.navigation(
    [
        st.Page("pages/sunrise.py", title="Sunrise", icon="📊", default=True),
        st.Page("pages/astra.py", title="Astra", icon="🚀"),
    ]
)
pg.run()
