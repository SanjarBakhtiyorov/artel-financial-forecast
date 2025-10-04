import streamlit as st
import pandas as pd
import datetime as dt
import os
import tempfile
from monthly_forecast_artel_service_full import main as run_forecast
# top of app.py
from monthly_forecast_artel_service_full import run_analysis

st.set_page_config(
    page_title="Artel Monthly Forecast Tool",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Artel Financial Forecast & YoY Analysis Tool")
st.markdown("Upload SAP Excel, enter inputs, and generate automated monthly reports with YoY comparison.")

# --- Sidebar Inputs ---
st.sidebar.header("Configuration")
month = st.sidebar.text_input("📅 Report Month (YYYY-MM)", dt.datetime.now().strftime("%Y-%m"))
vat_rate = st.sidebar.number_input("VAT Rate (%)", value=12.0)
call_center_revenue = st.sidebar.number_input("Call Center Revenue (USD, VAT-included)", value=3000.0)
admin_forecast = st.sidebar.number_input("Admin Costs Forecast (USD, net of VAT)", value=250000.0)
exclude_sundays = st.sidebar.checkbox("Exclude Sundays from Forecast", value=True)
nonempty_only = st.sidebar.checkbox("Forecast Non-Empty Days Only", value=True)

st.sidebar.markdown("---")
st.sidebar.header("YoY Comparison")
yoy_option = st.sidebar.selectbox(
    "Select YoY Baseline Mode",
    [
        "Skip YoY Comparison",
        "Same Dataset (auto detect)",
        "Upload Previous-Year File"
    ]
)

prev_file = None
if yoy_option == "Upload Previous-Year File":
    prev_file = st.sidebar.file_uploader("Upload Previous Year SAP Excel", type=["xls", "xlsx"])

# --- File Upload ---
st.subheader("📁 Upload SAP Excel (Current Period)")
uploaded_files = st.file_uploader(
    "Upload one or more SAP Excel files",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files and st.button("🚀 Run Forecast"):
    with st.spinner("Processing data... please wait ⏳"):
        # Save uploaded files to temp directory
        tmp_dir = tempfile.mkdtemp()
        paths = []
        for file in uploaded_files:
            path = os.path.join(tmp_dir, file.name)
            with open(path, "wb") as f:
                f.write(file.getbuffer())
            paths.append(path)

        # Save prev file if provided
        prev_path = None
        if prev_file:
            prev_path = os.path.join(tmp_dir, prev_file.name)
            with open(prev_path, "wb") as f:
                f.write(prev_file.getbuffer())

        # Run forecast main()
        try:
            # IMPORTANT: convert 12.0% -> 0.12
            output_path = run_analysis(
                in_files=paths,
                prev_file=prev_path if yoy_option == "Upload Previous-Year File" else None,
                call_center=float(call_center_revenue),
                admin_forecast=float(admin_forecast),
                vat_rate=float(vat_rate) / 100.0,
                vat_mode="exclusive",          # or your default
                month=month,
                forecast_nonempty_only=nonempty_only,
                no_exclude_sundays=not exclude_sundays,
                out_name="artel_report"
)
st.success("✅ Forecast completed!")
with open(output_path, "rb") as f:
    st.download_button("⬇️ Download Excel report", f, file_name=os.path.basename(output_path))

        except Exception as e:
            st.error(f"❌ Error: {e}")

else:
    st.info("👆 Upload your SAP Excel file(s), then click **Run Forecast**.")

