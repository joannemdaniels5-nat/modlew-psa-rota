
import io
from datetime import date
import streamlit as st
from rota_engine import read_template, ensure_monday, build_workbook

st.set_page_config(page_title="Rota Generator v15", layout="wide")

st.title("Rota Generator v15 — Weighted + Block-Stable (No Floaters)")

st.markdown(
"""
**Key behaviour (v15):**
- **Weights drive everything**: highest weight first, then least-used (default weight = 3)
- **No 30-min floaters**: tasks are assigned in blocks (Phones min 1.5h; others min 2.5h; Docman min 3h)
- **Front Desk & Triage are band-locked** (fixed shift bands)
- **Bookings**: SLGP first; if still behind weekly target, can use other sites (but **never** breaks Phones minimum)
- **EMIS/Docman**: JEN/BGS first; SLGP only if needed
- **Breaks only for >6h shifts** and only for staff with BreakRequired=True
- Unassigned time is shown as **Misc_Tasks** (grouped via block logic)
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
        st.success(f"Loaded template: {len(tpl.staff)} staff | Weekly targets: {tpl.weekly_targets}")
        if st.button("Generate rota and download Excel", type="primary"):
            wb = build_workbook(tpl, start_monday, weeks)
            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            out_name = f"rota_v15_{start_monday.isoformat()}_{weeks}w.xlsx"
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
