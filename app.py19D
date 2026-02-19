import io
from datetime import date
import streamlit as st
from rota_engine import read_template, ensure_monday, build_workbook

st.set_page_config(page_title="Rota Generator v14+++ (Full Refactor)", layout="wide")

st.title("Rota Generator v14+++ — Full Refactor")

st.markdown(
    """
**What this version does (rules-first):**
- Fixed Front Desk bands per site (08–11 / 11–13:30 / 13:30–16 / 16–18:30)
- Fixed Triage bands (SLGP + JEN) (08–10:30 / 10:30–13:00 / 13:30–16:00)
- Email mandatory 10:30–16:00 (JEN/BGS by default)
- Awaiting/PSA Admin mandatory 10:00–16:00 (site-of-day rule)
- Phones requirement from hourly matrix (applies to both half-hours) + +1 per call handler on leave
- Bookings uses hourly matrix if populated (otherwise soft weekly target only)
- Soft weekly targets for EMIS (20h) and Docman (12.5h) as fillers
- Breaks 12–14 spread by site, avoids tiny <1h fragments where possible
- Master timeline + formula-linked site timelines + counts and totals
"""
)

uploaded = st.file_uploader("Upload your rota template (.xlsx)", type=["xlsx"])

c1, c2 = st.columns(2)
with c1:
    start_date = st.date_input("Week commencing (Monday)", value=date.today())
with c2:
    weeks = int(st.number_input("Number of weeks", min_value=1, max_value=12, value=1, step=1))

start_monday = ensure_monday(start_date)

if uploaded:
    try:
        tpl = read_template(uploaded.getvalue())
        st.success(f"Loaded template: {len(tpl.staff)} staff | Weekly targets: {tpl.weekly_targets}")
        if st.button("Generate rota and download Excel", type="primary"):
            wb = build_workbook(tpl, start_monday, weeks)
            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            out_name = f"rota_v14ppp_{start_monday.isoformat()}_{weeks}w.xlsx"
            st.download_button(
                "📊 Download Excel rota",
                data=bio.getvalue(),
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error("Could not process the template.")
        st.exception(e)
else:
    st.info("Upload your completed template to generate the rota.")
