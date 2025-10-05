# app.py
import os
import tempfile
import datetime as dt
import streamlit as st
import pandas as pd
from typing import Dict
import altair as alt

# Backend entrypoint (must be in the repo)
# def run_analysis(..., prev_month_override=None) -> str
from monthly_forecast_artel_service_full import run_analysis

# ---------------------------- UI SETUP ----------------------------
st.set_page_config(page_title="Artel Monthly Forecast", page_icon="📊", layout="wide")
st.title("📊 Artel Financial Forecast & YoY Analysis Tool")
st.caption("Upload SAP Excel, enter inputs, and generate monthly reports with optional YoY comparison.")

# ---------------------------- HELPERS -----------------------------
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

# ---------------------------- SIDEBAR -----------------------------
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
baseline_month = st.sidebar.text_input("Previous period (YYYY-MM) – optional", value="")

# ---------------------------- ANALYSIS HELPERS (for extra views) ----------------------------
def _read_report_tables(xlsx_path: str) -> Dict[str, pd.DataFrame]:
    """Read all sheets from the generated Excel into a dict of DataFrames."""
    try:
        xls = pd.ExcelFile(xlsx_path)
        return {name: xls.parse(name) for name in xls.sheet_names}
    except Exception as e:
        st.error(f"Could not load report for charts: {e}")
        return {}

def _render_revenue_charts(tables: Dict[str, pd.DataFrame], vat_rate: float):
    st.markdown("### 📈 Revenue: Actual vs Forecast")
    actual_excl_cc, forecast_excl_cc = None, None
    rec = tables.get("Reconciliation")
    if rec is not None and {"Metric", "Value"}.issubset(rec.columns):
        def _pick(metric_key: str):
            m = rec.loc[rec["Metric"].astype(str).str.strip().eq(metric_key), "Value"]
            return float(m.iloc[0]) if not m.empty and pd.notna(m.iloc[0]) else None
        actual_excl_cc   = _pick("Revenue After VAT (excl CC)")
        forecast_excl_cc = _pick("Forecast (After VAT, excl CC)")

    if actual_excl_cc is None or forecast_excl_cc is None:
        summary = tables.get("Summary")
        if summary is not None and {"Metric","Value"}.issubset(summary.columns):
            actual = summary.loc[summary["Metric"].eq("Revenue After VAT"), "Value"]
            actual_excl_cc = float(actual.iloc[0]) if not actual.empty else None

    data = []
    if actual_excl_cc is not None:
        data.append(("Actual (After VAT excl CC)", actual_excl_cc))
    if forecast_excl_cc is not None:
        data.append(("Forecast (Month-End)", forecast_excl_cc))
    if data:
        dfc = pd.DataFrame(data, columns=["Metric", "Value"]).set_index("Metric")
        st.bar_chart(dfc)
    else:
        st.warning("No revenue values available for chart.")

def _render_by_correspondent(tables: Dict[str, pd.DataFrame], vat_rate: float):
    st.markdown("### 🏭 Top Correspondents (After VAT)")
    df = tables.get("By_Correspondent")
    if df is None or df.empty:
        st.info("By_Correspondent sheet not found.")
        return

    # Prefer net before VAT → convert to After VAT
    val_col = "net_before_vat_usd" if "net_before_vat_usd" in df.columns else "gross_amount_usd"

    work = df.copy()
    # Drop the synthetic CC line
    work = work[work["Correspondent"].astype(str) != "CALL_CENTER"]

    # Choose display name: use correspondent_name if present, else the numeric ID
    if "correspondent_name" in work.columns:
        work["display"] = work["correspondent_name"].fillna(work["Correspondent"].astype(str))
    else:
        work["display"] = work["Correspondent"].astype(str)

    # After VAT
    work["after_vat"] = work[val_col].astype(float) / (1.0 + vat_rate)

    # Aggregate in case multiple IDs share the same display name
    top = (
        work.groupby("display", as_index=False)["after_vat"]
            .sum()
            .sort_values("after_vat", ascending=False)
            .head(10)
    )

    if top.empty:
        st.info("No correspondent data to display.")
        return

    # Horizontal bar chart for clearer labels
    chart = (
        alt.Chart(top)
        .mark_bar()
        .encode(
            x=alt.X("after_vat:Q", title="After VAT (USD)"),
            y=alt.Y("display:N", sort="-x", title="Correspondent"),
            tooltip=[alt.Tooltip("display:N", title="Correspondent"),
                     alt.Tooltip("after_vat:Q", title="After VAT (USD)", format=",.2f")]
        )
        .properties(height=30 * len(top), width="container")
    )
    st.altair_chart(chart, use_container_width=True)

    # Also show a small table under the chart
    st.dataframe(
        top.rename(columns={"display": "Correspondent", "after_vat": "After VAT (USD)"}),
        use_container_width=True,
    )

def _render_warranty_share(tables: Dict[str, pd.DataFrame], vat_rate: float):
    st.markdown("### 🧩 Warranty Structure (After VAT)")
    df = tables.get("By_Warranty")
    if df is None or df.empty:
        st.info("By_Warranty sheet not found.")
        return
    base = df.copy()
    amt = "amount_usd" if "amount_usd" in base.columns else None
    g1  = "g1_transport_usd" if "g1_transport_usd" in base.columns else None
    if amt is None:
        st.info("Expected columns not present in By_Warranty.")
        return
    base["after_vat"] = (base[amt] - (base[g1] if g1 in base.columns else 0.0)) / (1.0 + vat_rate)
    base = base[["Warranty", "after_vat"]].groupby("Warranty", as_index=False).sum()
    base = base[base["after_vat"] > 0]
    if base.empty:
        st.info("No positive warranty values.")
        return
    total = base["after_vat"].sum()
    base["Share %"] = (base["after_vat"] / total * 100.0).round(2)
    st.dataframe(base)
    st.bar_chart(base.set_index("Warranty")["after_vat"])

def _render_daily_trend(tables: Dict[str, pd.DataFrame]):
    st.markdown("### 📅 Daily Trend (After VAT excl CC)")
    df = tables.get("Daily_Revenue")
    if df is None or df.empty:
        st.info("Daily_Revenue sheet not found.")
        return
    if not {"Date","After VAT (excl CC)"}.issubset(df.columns):
        st.info("Daily_Revenue does not have expected columns.")
        return
    d = df.copy()
    d = d[pd.to_datetime(d["Date"], errors="coerce").notna()]
    d = d.sort_values("Date")
    d = d.set_index("Date")[["After VAT (excl CC)"]]
    st.line_chart(d)

def _render_yoy_views(tables: Dict[str, pd.DataFrame]):
    st.markdown("### 📊 YoY Comparison")
    ym = tables.get("YoY_Monthly")
    yd = tables.get("YoY_Daily")
    if ym is not None and not ym.empty:
        st.write("**Monthly Comparison**")
        st.dataframe(ym, use_container_width=True)
    else:
        st.info("YoY_Monthly sheet not found.")
    if yd is not None and not yd.empty:
        st.write("**Daily Comparison**")
        cols = [c for c in yd.columns if "After VAT" in c]
        if "Day" in yd.columns and cols:
            tmp = yd.copy()
            tmp = tmp[tmp["Day"].apply(lambda x: isinstance(x, (int, float)))]
            tmp = tmp.set_index("Day")[cols]
            st.line_chart(tmp)
        st.dataframe(yd, use_container_width=True)
    else:
        st.info("YoY_Daily sheet not found.")

def _render_pl_forecast(tables: Dict[str, pd.DataFrame]):
    st.markdown("### 💼 P&L Forecast (After VAT)")
    df = tables.get("P&L_Forecast")
    if df is None or df.empty:
        st.info("P&L_Forecast sheet not found.")
        return

    # Handle both exports: sometimes first column is 'Line', sometimes 'Metric'
    if "Line" in df.columns:
        label_col = "Line"
    elif "Metric" in df.columns:
        label_col = "Metric"
        # Keep a view with a synthetic 'Line' for downstream text matching
        df = df.rename(columns={"Metric": "Line"})
        label_col = "Line"
    else:
        st.warning("P&L_Forecast sheet has no 'Line' or 'Metric' column.")
        st.dataframe(df, use_container_width=True)
        return

    # Find a numeric value column (prefer 'Value', else take the first numeric-like second column)
    value_col = "Value" if "Value" in df.columns else None
    if value_col is None:
        # try the second column if present
        if len(df.columns) >= 2:
            value_col = df.columns[1]
        else:
            st.warning("P&L_Forecast sheet has no recognizable value column.")
            st.dataframe(df, use_container_width=True)
            return

    # Show table
    st.dataframe(df[[label_col, value_col]], use_container_width=True)

    # Robust label lookup (works with en-dash/hyphen etc.)
    def _pick(label: str):
        if label_col not in df.columns or value_col not in df.columns:
            return None
        col_txt = df[label_col].astype(str).str.strip()
        # normalize dashes
        target = label.replace("–", "-").strip()
        match = col_txt.str.replace("–", "-", regex=False).eq(target)
        m = df.loc[match, value_col]
        return float(m.iloc[0]) if not m.empty and pd.notna(m.iloc[0]) else None

    total_rev_f   = _pick("Total Revenue – Forecast for month (After VAT)")
    admin_cost_f  = _pick("Administration Costs – Forecast for month (After VAT / net)")
    op_profit_f   = _pick("Operating Profit – Forecast for month")
    op_margin_pct = _pick("Operating Margin % – Forecast for month")

    # KPI cards
    c1, c2, c3, c4 = st.columns(4)
    if total_rev_f is not None:   c1.metric("Total Revenue (Forecast)", f"${total_rev_f:,.2f}")
    if admin_cost_f is not None:  c2.metric("Admin Costs (Forecast)", f"${admin_cost_f:,.2f}")
    if op_profit_f is not None:   c3.metric("Operating Profit (Forecast)", f"${op_profit_f:,.2f}")
    if op_margin_pct is not None: c4.metric("Operating Margin % (Forecast)", f"{op_margin_pct:.2f}%")

    # Simple bar chart
    bars = []
    if total_rev_f is not None:  bars.append(("Total Revenue", total_rev_f))
    if admin_cost_f is not None: bars.append(("Admin Costs", admin_cost_f))
    if op_profit_f is not None:  bars.append(("Operating Profit", op_profit_f))
    if bars:
        chart_df = pd.DataFrame(bars, columns=["Item", "Value"]).set_index("Item")
        st.bar_chart(chart_df)



# ---------------------------- UPLOADS -----------------------------
st.subheader("📁 Upload SAP Excel (Current Period)")
cur_files = st.file_uploader(
    "Upload one or more SAP Excel files (.xls / .xlsx)",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
)

prev_file_obj = None
if yoy_mode == "Upload Previous-Year File":
    st.subheader("📁 Upload Previous Year SAP Excel (.xls / .xlsx)")
    prev_file_obj = st.file_uploader("Upload baseline file", type=["xls", "xlsx"])

# Pre-flight warning for legacy .XLS
if _has_xls(cur_files) or (prev_file_obj and prev_file_obj.name.lower().endswith(".xls")):
    if not _xlrd_available():
        st.warning(
            "You uploaded a legacy **.XLS** file but the server has no **xlrd** installed. "
            "Either upload **.XLSX** or add `xlrd>=2.0.1` to `requirements.txt` and redeploy.",
            icon="⚠️",
        )

run = st.button("🚀 Run Forecast", key="btn_run_forecast")

# ---------------------------- RUN JOB -----------------------------
if run:
    if not cur_files:
        st.error("Please upload at least one SAP Excel file for the current period.")
        st.stop()

    if yoy_mode == "Upload Previous-Year File" and not prev_file_obj:
        st.error("Baseline Mode is 'Upload Previous-Year File' but no baseline file was provided.")
        st.stop()

    with st.spinner("Processing…"):
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
                prev_month_override=prev_month_override,
            )
        except Exception as e:
            st.error(f"❌ Error while running analysis: {e}")
            st.stop()

        # Persist context for reruns (so buttons work after click)
        st.session_state["report_ready"] = True
        st.session_state["report_path"] = output_path
        st.session_state["vat_rate_eff"] = float(vat_percent) / 100.0

        st.success("✅ Forecast completed!")

# ---------------------------- PERSISTENT DOWNLOAD ----------------------------
if st.session_state.get("report_ready") and st.session_state.get("report_path"):
    rp = st.session_state["report_path"]
    if os.path.isfile(rp):
        st.caption("Click below to download the Excel report.")
        with open(rp, "rb") as f:
            st.download_button(
                label="⬇️ Download Excel Report",
                data=f,
                file_name=os.path.basename(rp),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{os.path.basename(rp)}",
            )

# ---------------------------- ADDITIONAL ANALYSIS VIEWS ----------------------------
st.markdown("---")
st.subheader("📊 Additional Analysis Views")

col_a, col_b, col_c, col_d, col_e, col_f = st.columns(6)
btn_rev   = col_a.button("📈 Revenue Charts",     key="btn_rev")
btn_corr  = col_b.button("🏭 By Correspondent",   key="btn_corr")
btn_warr  = col_c.button("🧩 Warranty Structure", key="btn_warr")
btn_daily = col_d.button("📅 Daily Trend",        key="btn_daily")
btn_yoy   = col_e.button("📊 YoY Compare",        key="btn_yoy")
btn_pl    = col_f.button("💼 P&L",                key="btn_pl")

if any([btn_rev, btn_corr, btn_warr, btn_daily, btn_yoy, btn_pl]):
    rp = st.session_state.get("report_path")
    if not rp or not os.path.isfile(rp):
        st.warning("No generated report found. Please run the forecast first.")
    else:
        tables = _read_report_tables(rp)
        vat_rate_eff = st.session_state.get("vat_rate_eff", float(vat_percent) / 100.0)

        if btn_rev:
            _render_revenue_charts(tables, vat_rate_eff)
        if btn_corr:
            _render_by_correspondent(tables, vat_rate_eff)
        if btn_warr:
            _render_warranty_share(tables, vat_rate_eff)
        if btn_daily:
            _render_daily_trend(tables)
        if btn_yoy:
            _render_yoy_views(tables)
        if btn_pl:
            _render_pl_forecast(tables)

# ---------------------------- FOOTER ----------------------------
if not st.session_state.get("report_ready"):
    st.info("👆 Upload your SAP Excel file(s), adjust settings in the sidebar, then click **Run Forecast**.")




