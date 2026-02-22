import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import streamlit as st
from datetime import date
import io

from rota_engine_v37_7 import read_template, build_workbook, ensure_monday, recalc_workbook_from_site_timelines

st.set_page_config(page_title="ModLew PSA Rota Generator", page_icon="🗓️", layout="wide")
st.title("ModLew PSA Rota Generator")

st.markdown("### 1) Generate rota")
uploaded = st.file_uploader("Upload input template (.xlsx)", type=["xlsx"], key="tpl")
weeks = st.number_input("Weeks", min_value=1, max_value=8, value=1, step=1)
wk_start = st.date_input("Week commencing (any date in that week)", value=date.today())

if st.button("Generate rota", type="primary", disabled=(uploaded is None)):
    try:
        tpl = read_template(uploaded.getvalue())
        start_monday = ensure_monday(wk_start)
        wb = build_workbook(tpl, start_monday, int(weeks))
        bio = io.BytesIO()
        wb.save(bio)
        st.success("Rota generated.")
        st.download_button(
            "Download rota (.xlsx)",
            data=bio.getvalue(),
            file_name=f"rota_{start_monday.isoformat()}_{int(weeks)}w.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.exception(e)

st.markdown("---")
st.markdown("### 2) Recalculate (after you edit the site timelines)")
st.caption("Upload the generated rota after edits. This rebuilds Coverage + Totals from the edited timelines and re-applies colours.")
uploaded2 = st.file_uploader("Upload edited rota (.xlsx)", type=["xlsx"], key="edited")

if st.button("Recalculate workbook", disabled=(uploaded2 is None)):
    try:
        out_bytes = recalc_workbook_from_site_timelines(uploaded2.getvalue())
        st.success("Recalculated.")
        st.download_button(
            "Download recalculated workbook (.xlsx)",
            data=out_bytes,
            file_name="rota_recalculated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.exception(e)
