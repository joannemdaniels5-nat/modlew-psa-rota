
import streamlit as st
from datetime import date
import io

# IMPORT THE FIXED ENGINE
from rota_engine_v34_6_HARDFIX_PHONES_EMAIL import (
    read_template,
    build_workbook,
    ensure_monday,
    recalc_workbook_from_site_timelines
)

st.set_page_config(page_title="ModLew PSA Rota Generator", page_icon="📅")

st.title("ModLew PSA Rota Generator v37")

uploaded = st.file_uploader("Upload template (.xlsx)", type=["xlsx"])

if uploaded:
    tpl = read_template(uploaded.getvalue())

    start_date = st.date_input("Start Monday", value=date.today())
    weeks = st.number_input("Number of weeks", min_value=1, max_value=8, value=1)

    if st.button("Generate Rota"):
        monday = ensure_monday(start_date)
        wb = build_workbook(tpl, monday, weeks)

        bio = io.BytesIO()
        wb.save(bio)

        st.download_button(
            label="Download Generated Rota",
            data=bio.getvalue(),
            file_name=f"rota_{monday}_{weeks}w.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    edited = st.file_uploader("Upload edited rota to recalculate coverage", type=["xlsx"], key="recalc")

    if edited:
        out_bytes = recalc_workbook_from_site_timelines(edited.getvalue())
        st.download_button(
            label="Download Recalculated Workbook",
            data=out_bytes,
            file_name="recalculated_rota.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
