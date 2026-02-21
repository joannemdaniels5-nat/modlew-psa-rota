import streamlit as st
from datetime import date
import io

from rota_engine_v34_6 import read_template, build_workbook, ensure_monday, recalc_workbook_from_site_timelines

st.set_page_config(page_title="ModLew PSA Rota Generator", page_icon="🗓️", layout="wide")

st.title("ModLew PSA Rota Generator")

tab1, tab2 = st.tabs(["Generate rota", "Recalculate from edited timelines"])

with tab1:
    st.markdown("Upload the **input template** and generate a rota workbook.")
    uploaded = st.file_uploader("Upload input template (.xlsx)", type=["xlsx"], key="tpl")

    c1, c2 = st.columns(2)
    with c1:
        start_date = st.date_input("Week commencing (Monday)", value=date.today(), key="wc")
    with c2:
        weeks = int(st.number_input("Number of weeks", min_value=1, max_value=12, value=1, step=1, key="weeks"))

    if uploaded:
        try:
            tpl = read_template(uploaded.getvalue())
            st.success("Template loaded.")
            if st.button("Generate rota workbook", type="primary"):
                wb = build_workbook(tpl, ensure_monday(start_date), weeks)
                bio = io.BytesIO()
                wb.save(bio)
                bio.seek(0)
                out_name = f"rota_{ensure_monday(start_date).isoformat()}_{weeks}w.xlsx"
                st.download_button("Download rota workbook", data=bio.getvalue(), file_name=out_name,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error("Could not process the template.")
            st.exception(e)
    else:
        st.info("Upload the input template to generate a rota.")

with tab2:
    st.markdown(
        "Upload a rota workbook **you have edited** (site timelines), and it will rebuild **Coverage** and **Totals** "
        "from those timelines. This is the 'safe edit' workflow."
    )
    edited = st.file_uploader("Upload edited rota workbook (.xlsx)", type=["xlsx"], key="edited")
    if edited:
        try:
            if st.button("Recalculate Coverage + Totals", type="primary"):
                out_bytes = recalc_workbook_from_site_timelines(edited.getvalue())
                st.download_button(
                    "Download recalculated workbook",
                    data=out_bytes,
                    file_name="rota_recalculated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as e:
            st.error("Could not recalculate the workbook.")
            st.exception(e)
