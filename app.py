import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import streamlit as st
from datetime import date
import io

from rota_engine_v37_fast import (
    read_template,
    build_workbook,
    ensure_monday,
    recalc_workbook_from_site_timelines,
)

st.set_page_config(page_title="ModLew PSA Rota Generator", page_icon="📅")
st.title("ModLew PSA Rota Generator v37")

tab1, tab2 = st.tabs(["Generate rota", "Recalculate from edited timelines"])

# --- Session state holders so download buttons persist after reruns
if "generated_bytes" not in st.session_state:
    st.session_state.generated_bytes = None
if "generated_name" not in st.session_state:
    st.session_state.generated_name = None
if "recalc_bytes" not in st.session_state:
    st.session_state.recalc_bytes = None
if "recalc_name" not in st.session_state:
    st.session_state.recalc_name = "recalculated_rota.xlsx"

with tab1:
    st.markdown("Upload the **input template** and generate a rota workbook.")
    uploaded = st.file_uploader("Upload input template (.xlsx)", type=["xlsx"], key="tpl")

    if uploaded:
        tpl = read_template(uploaded.getvalue())
        start_date = st.date_input("Start Monday", value=date.today(), key="start")
        weeks = st.number_input("Number of weeks", min_value=1, max_value=8, value=1, step=1, key="weeks")

        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("Generate Rota", type="primary"):
                monday = ensure_monday(start_date)
                with st.spinner("Building rota workbook..."):
                    wb = build_workbook(tpl, monday, int(weeks))
                    bio = io.BytesIO()
                    wb.save(bio)

                st.session_state.generated_bytes = bio.getvalue()
                st.session_state.generated_name = f"rota_{monday}_{int(weeks)}w.xlsx"

        # Download button persists once generated
        if st.session_state.generated_bytes:
            st.download_button(
                label="Download Generated Rota",
                data=st.session_state.generated_bytes,
                file_name=st.session_state.generated_name or "rota.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("Upload the template to enable generation.")

with tab2:
    st.markdown("Upload an **edited rota workbook** (site timelines) and recalculate Coverage + Totals (with colours).")
    edited = st.file_uploader("Upload edited rota (.xlsx)", type=["xlsx"], key="recalc")

    if edited:
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("Recalculate", type="primary"):
                with st.spinner("Recalculating coverage + totals..."):
                    out_bytes = recalc_workbook_from_site_timelines(edited.getvalue())
                st.session_state.recalc_bytes = out_bytes

        if st.session_state.recalc_bytes:
            st.download_button(
                label="Download Recalculated Workbook",
                data=st.session_state.recalc_bytes,
                file_name=st.session_state.recalc_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("Upload an edited rota workbook to enable recalculation.")
