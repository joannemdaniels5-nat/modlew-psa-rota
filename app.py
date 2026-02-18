
import io
from datetime import date
import streamlit as st

from rota_engine_fixed import read_template, build_workbook, ensure_monday

st.set_page_config(page_title="Rota Generator", layout="wide")

st.title("Rota Generator")
st.caption("Upload your rota template → generate the Excel rota.")

uploaded = st.file_uploader("Upload rota template (.xlsx)", type=["xlsx"])

c1, c2 = st.columns(2)
with c1:
    start_date = st.date_input("Week commencing (Monday)", value=date.today())
with c2:
    weeks = int(st.number_input("Number of weeks", min_value=1, max_value=12, value=1, step=1))

if uploaded:
    try:
        tpl = read_template(uploaded.getvalue())
        st.success("Template loaded.")
        if st.button("Generate rota", type="primary"):
            start_monday = ensure_monday(start_date)
            wb = build_workbook(tpl, start_monday, weeks)
            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            st.download_button(
                "Download rota Excel",
                data=bio.getvalue(),
                file_name=f"rota_{start_monday.isoformat()}_{weeks}w.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error("Template could not be processed.")
        st.exception(e)
else:
    st.info("Upload your template to begin.")
