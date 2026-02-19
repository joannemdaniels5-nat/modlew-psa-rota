import streamlit as st
from datetime import date
import io

from rota_engine_v31 import read_template, build_workbook, ensure_monday

st.set_page_config(page_title="ModLew PSA Rota Generator", page_icon="🗓️", layout="wide")

st.title("PSA Rota Generator")

with st.expander("What this does", expanded=False):
    st.markdown(
        """
- Upload the **rota template** (xlsx).
- Choose the week start date.
- Generate an **Excel rota workbook** with:
  - Site timelines (editable)
  - Coverage sheet (updates from timelines)
  - Totals sheet (updates from timelines)
"""
    )

uploaded = st.file_uploader("Upload template (.xlsx)", type=["xlsx"])

col1, col2 = st.columns([1,1])
with col1:
    week_start = st.date_input("Week starting (Monday)", value=ensure_monday(date.today()))
with col2:
    weeks = st.number_input("Number of weeks", min_value=1, max_value=4, value=1, step=1)

if uploaded is None:
    st.info("Upload the template to begin.")
    st.stop()

try:
    tpl = read_template(uploaded.getvalue())
except Exception as e:
    st.error(f"Template error: {e}")
    st.stop()

if st.button("Generate rota", type="primary"):
    with st.spinner("Building workbook…"):
        wb = build_workbook(tpl, ensure_monday(week_start), int(weeks))
        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)

    st.success("Done.")
    st.download_button(
        "Download rota workbook",
        data=bio.getvalue(),
        file_name=f"rota_{ensure_monday(week_start).isoformat()}_{weeks}w.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
