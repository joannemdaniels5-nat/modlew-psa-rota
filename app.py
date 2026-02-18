import io
from datetime import date
import streamlit as st

from rota_engine import read_template, ensure_monday, build_workbook


# =========================
# Page setup + styling
# =========================
st.set_page_config(page_title="PSA Rota Planner", layout="wide")

st.markdown(
    """
<style>
/* Subtle dashboard feel */
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
h1, h2, h3 { letter-spacing: -0.02em; }
.small-muted { color: rgba(49,51,63,0.65); font-size: 0.95rem; margin-top: -0.6rem; }
.card {
  background: rgba(250, 250, 252, 1);
  border: 1px solid rgba(49,51,63,0.10);
  border-radius: 16px;
  padding: 16px 16px 8px 16px;
  box-shadow: 0 1px 2px rgba(0,0,0,0.03);
}
.card h3 { margin: 0 0 0.5rem 0; }
.hr { border-top: 1px solid rgba(49,51,63,0.10); margin: 0.75rem 0 0.75rem 0; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown("""# PSA Rota Planner
<div class="small-muted">Operational scheduling dashboard</div>
""", unsafe_allow_html=True)

# =========================
# State
# =========================
if "out_bytes" not in st.session_state:
    st.session_state.out_bytes = None
    st.session_state.out_name = None
    st.session_state.stats = None


# =========================
# Layout
# =========================
left, right = st.columns([1.05, 0.95], gap="large")

with left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Generate rota")

    uploaded = st.file_uploader(
        "Upload rota template (.xlsx)",
        type=["xlsx"],
        help="Use the standard rota template and upload the filled file here.",
        key="template_uploader",
    )

    c1, c2 = st.columns(2)
    with c1:
        start_date = st.date_input("Week commencing", value=date.today(), help="Select any date in the week; it will snap to Monday.")
    with c2:
        weeks = int(
            st.number_input(
                "Weeks",
                min_value=1,
                max_value=12,
                value=1,
                step=1,
                help="How many weeks to generate.",
            )
        )

    start_monday = ensure_monday(start_date)

    generate = st.button(
        "🚀 Generate rota",
        type="primary",
        use_container_width=True,
        disabled=(uploaded is None),
    )

    st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("ℹ️ How to use this planner", expanded=False):
        st.markdown(
            """
- Upload your completed template (Staff, WorkingHours, Leave, Targets, etc.)
- Choose a week commencing date and number of weeks
- Click **Generate rota**
- Download the Excel output and make any edits on the **MasterTimeline** sheet (other sheets update from it)
            """
        )

with right:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Status")

    if uploaded is None and st.session_state.out_bytes is None:
        st.info("Upload a template to get started.")
    elif uploaded is not None:
        st.success("Template ready.")
        st.caption(f"Week commencing: **{start_monday.isoformat()}** • Weeks: **{weeks}**")

    # Generate (runs once user clicks)
    if generate and uploaded is not None:
        try:
            with st.spinner("Building rota…"):
                tpl = read_template(uploaded.getvalue())
                wb = build_workbook(tpl, start_monday, weeks)

                bio = io.BytesIO()
                wb.save(bio)
                bio.seek(0)

                out_name = f"rota_{start_monday.isoformat()}_{weeks}w.xlsx"

                st.session_state.out_bytes = bio.getvalue()
                st.session_state.out_name = out_name
                # Minimal, non-noisy stats
                staff_count = len(getattr(tpl, "staff", []) or [])
                st.session_state.stats = {
                    "staff_count": staff_count,
                    "weeks": weeks,
                    "week_start": start_monday.isoformat(),
                }

            st.success("Rota generated.")
        except Exception as e:
            # Keep it user-friendly; no stack trace spam on the UI
            st.error("Couldn’t generate the rota from this template.")
            with st.expander("Show details", expanded=False):
                st.write(str(e))

    if st.session_state.out_bytes:
        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        stats = st.session_state.stats or {}
        if stats.get("staff_count"):
            st.caption(f"Staff detected: **{stats['staff_count']}**")

        st.download_button(
            "⬇️ Download Excel rota",
            data=st.session_state.out_bytes,
            file_name=st.session_state.out_name or "rota.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("Tip: edit **MasterTimeline** to adjust assignments; linked sheets will update automatically.")

    st.markdown('</div>', unsafe_allow_html=True)
