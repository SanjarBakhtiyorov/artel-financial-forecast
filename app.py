# app.py
import os
import tempfile
import datetime as dt
import streamlit as st

# Import the wrapper that runs your full pipeline and returns the output Excel path
from monthly_forecast_artel_service_full import run_analysis

st.set_page_config(page_title="Artel Monthly Forecast", page_icon="📊", layout="wide")

st.title("📊 Artel Financial Forecast & YoY Analysis Tool")
st.caption("Upload SAP Excel, enter inputs, and generate automated monthly reports with YoY comparison.")

# ---------------- Sidebar: Configuration ----------------
st.sidebar.header("Configuration")

default_month = dt.datetime.now().strftime("%Y-%m")
month = st.sidebar.text_input("📅 Report Month (YYYY-MM)", value=default_month)

vat_percent = st.sidebar.number_input("VAT Rate (%)", min_value=0.0, max_value=100.0, value=12.0, step=0.5)

# VAT accounting mode used by your backend (choose the one your code expects)
vat_mode_label = st.sidebar.selectbox("VAT Mode", ["extract", "exclusive"], index=0)
#   - "extract"   -> your code subtracts VAT from VAT-included net
#   - "exclusive" -> your code expects net and adds VAT later (if used)

call_center_revenue = st.sidebar.number_input("Call Center Revenue (USD, VAT-included)", value=3000.0, step=100.0)
admin_forecast      = st.sidebar.number_input("Admin Costs Forecast (USD, net of VAT)", value=250000.0, step=1000.0)

st.sidebar.markdown("---")
exclude_sundays = st.sidebar.checkbox("Exclude Sundays from Forecast", value=True)
nonempty_only   = st.sidebar.checkbox("Forecast Non-Empty Days Only", value=True)

st.sidebar.markdown("---")
st.sidebar.header("YoY Comparison")
baseline_month = st.sidebar.text_input(
    "Previous period (YYYY-MM) – optional",
    value=""
)
yoy_mode = st.sidebar.selectbox(
    "Baseline Mode",
    ["Skip YoY Comparison", "Same Dataset (auto-detect)", "Upload Previous-Year File"],
    index=0
)

# ---------------- Main: File Uploads ----------------
st.subheader("📁 Upload SAP Excel (Current Period)")
uploaded_files = st.file_uploader(
    "Upload one or more SAP Excel files (.xls / .xlsx)",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

prev_file_obj = None
if yoy_mode == "Upload Previous-Year File":
    prev_file_obj = st.file_uploader("Upload Previous Year SAP Excel (.xls / .xlsx)", type=["xls", "xlsx"])

run = st.button("🚀 Run Forecast")

# ---------------- Run Pipeline ----------------
if run:
    if not uploaded_files:
        st.error("Please upload at least one SAP Excel file for the current period.")
        st.stop()

    with st.spinner("Processing…"):
        # Save uploaded files to a temp folder
        tmp_dir = tempfile.mkdtemp()
        in_paths = []
        for f in uploaded_files:
            p = os.path.join(tmp_dir, f.name)
            with open(p, "wb") as out:
                out.write(f.getbuffer())
            in_paths.append(p)

        prev_path = None
        if prev_file_obj is not None:
            prev_path = os.path.join(tmp_dir, prev_file_obj.name)
            with open(prev_path, "wb") as out:
                out.write(prev_file_obj.getbuffer())

        try:
            output_path = run_analysis(
                in_files=in_paths,
                prev_file=prev_path if yoy_mode == "Upload Previous-Year File" else None,
                call_center=float(call_center_revenue),
                admin_forecast=float(admin_forecast),
                vat_rate=float(vat_percent) / 100.0,
                vat_mode=vat_mode_label,
                month=month,
                forecast_nonempty_only=bool(nonempty_only),
                no_exclude_sundays=not bool(exclude_sundays),
                out_name="artel_report",
                prev_month_override=baseline_month.strip() or None,   # <— NEW
            )

        except Exception as e:
            st.error(f"❌ Error while running analysis: {e}")
            st.stop()

        st.success("✅ Forecast completed!")
        st.caption("Click below to download the Excel report.")
        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Excel Report",
                data=f,
                file_name=os.path.basename(output_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("👆 Upload your SAP Excel file(s), adjust settings in the sidebar, then click **Run Forecast**.")


