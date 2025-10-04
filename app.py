# app.py
import os
import tempfile
import datetime as dt
import streamlit as st

# Your backend wrapper must exist in the same repo
# def run_analysis(..., prev_month_override=None) -> str (path to output .xlsx)
from monthly_forecast_artel_service_full import run_analysis


# ------------------------------- UI SETUP -------------------------------
st.set_page_config(page_title="Artel Monthly Forecast", page_icon="📊", layout="wide")
st.title("📊 Artel Financial Forecast & YoY Analysis Tool")
st.caption("Upload SAP Excel, enter inputs, and generate monthly reports with YoY comparison.")


# ------------------------------- HELPERS --------------------------------
def _xlrd_available() -> bool:
    try:
        import xlrd  # noqa: F401
        return True
    except Exception:
        return False


def _has_xls(files) -> bool:
    if not files:
        return False
    return any(os.path.splitext(f.name)[1].lower() == ".xls" for f in files)


# ------------------------------- SIDEBAR --------------------------------
st.sidebar.header("Configuration")

default_month = dt.datetime.now().strftime("%Y-%m")
month = st.sidebar.text_input("📅 Report Month (YYYY-MM)", value=default_month)

vat_percent = st.sidebar.number_input("VAT Rate (%)", min_value=0.0, max_value=100.0, value=12.0, step=0.5)
vat_mode = st.sidebar.selectbox("VAT Mode", ["extract", "exclusive"], index=0)

call_center_revenue = st.sidebar.number_input("Call Center Revenue (USD, VAT-included)", value=3000.0, step=100.0)
admin_forecast = st.sidebar.number_input("Admin Costs Forecast (USD, net of VAT)", value=250000.0, step=1000.0)

st.sidebar.markdown("---")
exclude_sundays = st.sidebar.checkbox("Exclude Sundays from Forecast", value=True)
nonempty_only = st.sidebar.checkbox("Forecast Non-Empty Days Only", value=True)

st.sidebar.markdown("---")
st.sidebar.header("YoY Comparison")
yoy_mode = st.sidebar.selectbox(
    "Baseline Mode",
    ["Skip YoY Comparison", "Same Dataset (auto-detect)", "Upload Previous-Year File"],
    index=0
)

# Optional override if user wants to compare to a different month than "same month last year"
baseline_month = st.sidebar.text_input("Previous period (YYYY-MM) – optional", value="")


# ------------------------------- UPLOADS --------------------------------
st.subheader("📁 Upload SAP Excel (Current Period)")
cur_files = st.file_uploader(
    "Upload one or more SAP Excel files (.xls / .xlsx)",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

prev_file_obj = None
if yoy_mode == "Upload Previous-Year File":
    st.subheader("📁 Upload Previous Year SAP Excel (.xls / .xlsx)")
    prev_file_obj = st.file_uploader("Upload baseline file", type=["xls", "xlsx"])


# Pre-flight warnings
if _has_xls(cur_files) or (prev_file_obj and prev_file_obj.name.lower().endswith(".xls")):
    if not _xlrd_available():
        st.error(
            "You uploaded a .XLS file but the server has no `xlrd` installed. "
            "Add `xlrd>=2.0.1` to requirements.txt and redeploy, or upload .XLSX."
        )

run = st.button("🚀 Run Forecast")


# ------------------------------- RUN JOB --------------------------------
if run:
    if not cur_files:
        st.error("Please upload at least one SAP Excel file for the current period.")
        st.stop()

    if yoy_mode == "Upload Previous-Year File" and not prev_file_obj:
        st.error("Baseline Mode is 'Upload Previous-Year File' but no baseline file was provided.")
        st.stop()

    with st.spinner("Processing…"):
        # Save uploaded files to a temp dir
        tmp_dir = tempfile.mkdtemp()
        cur_paths = []
        for f in cur_files:
            p = os.path.join(tmp_dir, f.name)
            with open(p, "wb") as out:
                out.write(f.getbuffer())
            cur_paths.append(p)

        prev_path = None
        if prev_file_obj is not None:
            prev_path = os.path.join(tmp_dir, prev_file_obj.name)
            with open(prev_path, "wb") as out:
                out.write(prev_file_obj.getbuffer())

        # Decide prev-file usage
        prev_file_for_run = prev_path if yoy_mode == "Upload Previous-Year File" else None
        prev_month_override = baseline_month.strip() or None

        try:
            output_path = run_analysis(
                in_files=cur_paths,
                prev_file=prev_file_for_run,
                call_center=float(call_center_revenue),
                admin_forecast=float(admin_forecast),
                vat_rate=float(vat_percent) / 100.0,  # 12.0% -> 0.12
                vat_mode=vat_mode,
                month=month,
                forecast_nonempty_only=bool(nonempty_only),
                no_exclude_sundays=not bool(exclude_sundays),
                out_name="artel_report",
                prev_month_override=prev_month_override,  # NEW: supports manual baseline month
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
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

else:
    st.info("👆 Upload your SAP Excel file(s), adjust settings in the sidebar, then click **Run Forecast**.")
