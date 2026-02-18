import io
from datetime import date
import streamlit as st

from rota_engine import read_template, ensure_monday, build_workbook

st.set_page_config(page_title="Rota Generator", layout="wide")

st.title("Rota Generator")

with st.expander("How it works", expanded=False):
    st.markdown(
        """
Upload your rota template, pick the week commencing date and number of weeks, then generate an Excel rota.

Key behaviours:
- Fixed Front Desk and Triage bands
- Phones requirement from hourly matrix (applies to both half-hours) + uplift for call-handler leave
- Bookings driven to weekly target (and hourly matrix if provided)
- EMIS / Docman fill remaining capacity toward weekly targets
- Breaks only for shifts > 6h, spread 12:00–14:00 site-balanced
- No 30-minute floaters (single-slot fragments are smoothed)
- Site timelines are the single source of truth: editing them updates Totals
"""
    )

uploaded = st.file_uploader("Upload rota template (.xlsx)", type=["xlsx"])

c1, c2 = st.columns(2)
with c1:
    start_date = st.date_input("Week commencing (Monday)", value=date.today())
with c2:
    weeks = int(st.number_input("Number of weeks", min_value=1, max_value=12, value=1, step=1))

start_monday = ensure_monday(start_date)

if uploaded:
    try:
        tpl = read_template(uploaded.getvalue())
        if st.button("Generate Excel rota", type="primary"):
            wb = build_workbook(tpl, start_monday, weeks)
            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            out_name = f"rota_{start_monday.isoformat()}_{weeks}w.xlsx"
            st.download_button(
                "Download Excel rota",
                data=bio.getvalue(),
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error("Could not process that template.")
        st.exception(e)
else:
    st.info("Upload your template to begin.")
