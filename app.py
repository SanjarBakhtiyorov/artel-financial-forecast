# -*- coding: utf-8 -*-
"""
Created on Sat Oct 11 17:44:54 2025

@author: 6185
"""

# app.py
import os
import re
import tempfile
import datetime as dt
from typing import Dict, Optional, List

import streamlit as st
import pandas as pd
import numpy as np

# Try Altair; fall back to Streamlit charts if unavailable
try:
    import altair as alt  # noqa: F401
except Exception:  # pragma: no cover
    alt = None

# Backend entrypoint (must exist in the repo)
# def run_analysis(..., prev_month_override=None) -> str
from monthly_forecast_artel_service_full import run_analysis

# ============================ UI SETUP ============================
st.set_page_config(page_title="Artel Monthly Forecast", page_icon="📊", layout="wide")
st.title("📊 Artel Financial Forecast & YoY Analysis Tool")
st.caption("Upload SAP Excel, enter inputs, and generate monthly reports with optional YoY comparison.")

# ============================ HELPERS =============================
def _xlrd_available() -> bool:
    try:
        import xlrd  # noqa: F401
        return True
    except Exception:
        return False


def _has_xls(files: Optional[List[st.runtime.uploaded_file_manager.UploadedFile]]) -> bool:
    if not files:
        return False
    return any(os.path.splitext(f.name)[1].lower() == ".xls" for f in files)


def fmt_money(x) -> str:
    """12345.6 -> '12 345.60' (space thousands)."""
    try:
        return f"{float(x):,.2f}".replace(",", " ")
    except Exception:
        return str(x)


def fmt_pct_ratio(x) -> str:
    """Treat x as ratio (e.g., 0.23 -> 23.00%)."""
    try:
        return f"{float(x):.2f}%"
    except Exception:
        return str(x)


def fmt_pct_metric(x) -> str:
    """
    For metrics coming from sheets (may already be like 12 or 0.12).
    Convert fractions in [-1,1] to percent; keep others as-is.
    """
    try:
        v = float(x)
        if -1.0 <= v <= 1.0:
            v *= 100.0
        return f"{v:.2f}%"
    except Exception:
        return str(x)


def _save_uploaded_files(files: List[st.runtime.uploaded_file_manager.UploadedFile], tmp_dir: str) -> List[str]:
    paths = []
    for f in files:
        p = os.path.join(tmp_dir, f.name)
        with open(p, "wb") as out:
            out.write(f.getbuffer())
        paths.append(p)
    return paths


@st.cache_data(show_spinner=False)
def _read_report_tables_cached(xlsx_path: str) -> Dict[str, pd.DataFrame]:
    """Read all sheets from the generated Excel into a dict of DataFrames."""
    try:
        xls = pd.ExcelFile(xlsx_path)
        return {name: xls.parse(name) for name in xls.sheet_names}
    except Exception as e:
        st.error(f"Could not load report for charts: {e}")
        return {}


def _read_report_tables(xlsx_path: str) -> Dict[str, pd.DataFrame]:
    if not os.path.isfile(xlsx_path):
        return {}
    return _read_report_tables_cached(xlsx_path)


def _pick_value(df: pd.DataFrame, label_col: str, value_col: str, label: str) -> Optional[float]:
    s = df.loc[df[label_col].astype(str).str.strip().eq(label), value_col]
    if s.empty or pd.isna(s.iloc[0]):
        return None
    try:
        return float(s.iloc[0])
    except Exception:
        return None


# ============================ SIDEBAR =============================
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

# ============================ UPLOADS =============================
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

# ============================ RENDERERS ===========================
def render_revenue_charts(tables: Dict[str, pd.DataFrame], vat_rate: float):
    st.markdown("### 📈 Revenue: Actual vs Forecast")
    actual_excl_cc, forecast_excl_cc = None, None

    rec = tables.get("Reconciliation")
    if rec is not None and {"Metric", "Value"}.issubset(rec.columns):
        def _pick(metric_key: str):
            m = rec.loc[rec["Metric"].astype(str).str.strip().eq(metric_key), "Value"]
            return float(m.iloc[0]) if not m.empty and pd.notna(m.iloc[0]) else None

        actual_excl_cc   = _pick("Revenue After VAT (excl CC)")
        forecast_excl_cc = _pick("Forecast (After VAT, excl CC)")

    if (actual_excl_cc is None) or (forecast_excl_cc is None):
        summary = tables.get("Summary")
        if summary is not None and {"Metric", "Value"}.issubset(summary.columns):
            m = summary.loc[summary["Metric"].eq("Revenue After VAT"), "Value"]
            if not m.empty:
                actual_excl_cc = float(m.iloc[0])

    data = []
    if actual_excl_cc is not None:
        data.append(("Actual (After VAT excl CC)", actual_excl_cc))
    if forecast_excl_cc is not None:
        data.append(("Forecast (Month-End)", forecast_excl_cc))

    if not data:
        st.warning("No revenue values available for chart.")
        return

    dfc = pd.DataFrame(data, columns=["Metric", "Value"])
    if alt:
        chart = (
            alt.Chart(dfc)
            .mark_bar()
            .encode(x="Metric:N", y=alt.Y("Value:Q", title="USD"), tooltip=["Metric", alt.Tooltip("Value:Q", format=",.2f")])
            .properties(width="container", height=300)
        )
        st.altair_chart(chart, use_container_width=True)
    else:
        st.bar_chart(dfc.set_index("Metric"))


def render_by_correspondent(tables: Dict[str, pd.DataFrame], vat_rate: float):
    st.markdown("### 🏭 Top Correspondents (After VAT)")
    df = tables.get("By_Correspondent")
    if df is None or df.empty:
        st.info("By_Correspondent sheet not found.")
        return

    val_col = "net_before_vat_usd" if "net_before_vat_usd" in df.columns else "gross_amount_usd"
    work = df.copy()

    # Exclude call center pseudo-row if present
    if "Correspondent" in work.columns:
        work = work[work["Correspondent"].astype(str) != "CALL_CENTER"]

    # Display name
    if "correspondent_name" in work.columns:
        work["display"] = work["correspondent_name"].fillna(work.get("Correspondent", "").astype(str))
    else:
        work["display"] = work.get("Correspondent", "").astype(str)

    # After VAT
    work["after_vat"] = pd.to_numeric(work[val_col], errors="coerce") / (1.0 + float(vat_rate))

    # Aggregate & Top 10
    top = (
        work.groupby("display", as_index=False)["after_vat"]
        .sum()
        .sort_values("after_vat", ascending=False)
        .head(10)
    )
    if top.empty:
        st.info("No correspondent data to display.")
        return

    top["after_vat_fmt"] = top["after_vat"].apply(fmt_money)

    if alt:
        chart = (
            alt.Chart(top)
            .mark_bar()
            .encode(
                x=alt.X("after_vat:Q", title="After VAT (USD)"),
                y=alt.Y("display:N", sort="-x", title="Correspondent"),
                tooltip=[
                    alt.Tooltip("display:N", title="Correspondent"),
                    alt.Tooltip("after_vat:Q", title="After VAT (USD)", format=",.2f"),
                ],
            )
            .properties(height=30 * len(top), width="container")
        )
        st.altair_chart(chart, use_container_width=True)
    else:
        st.bar_chart(top.set_index("display")["after_vat"])

    table = top[["display", "after_vat_fmt"]].rename(
        columns={"display": "Correspondent", "after_vat_fmt": "After VAT (USD)"}
    )
    st.dataframe(table, use_container_width=True)


def render_warranty_share(tables: Dict[str, pd.DataFrame], vat_rate: float):
    st.markdown("### 🧩 Warranty Structure (After VAT)")
    df = tables.get("By_Warranty")
    if df is None or df.empty:
        st.info("By_Warranty sheet not found.")
        return

    base = df.copy()
    amt = "amount_usd" if "amount_usd" in base.columns else None
    g1 = "g1_transport_usd" if "g1_transport_usd" in base.columns else None
    if amt is None:
        st.info("Expected columns not present in By_Warranty.")
        return

    base["after_vat"] = (pd.to_numeric(base[amt], errors="coerce") - (pd.to_numeric(base[g1], errors="coerce") if g1 in base.columns else 0.0)) / (1.0 + float(vat_rate))
    base = base[["Warranty", "after_vat"]].groupby("Warranty", as_index=False).sum()
    base = base[base["after_vat"] > 0]

    if base.empty:
        st.info("No positive warranty values.")
        return

    total = base["after_vat"].sum()
    base["Share %"] = (base["after_vat"] / total * 100.0).round(2)

    st.dataframe(base.rename(columns={"after_vat": "After VAT (USD)"}), use_container_width=True)

    if alt:
        chart = (
            alt.Chart(base)
            .mark_bar()
            .encode(
                x=alt.X("Warranty:N", title="Warranty"),
                y=alt.Y("after_vat:Q", title="After VAT (USD)"),
                tooltip=[alt.Tooltip("Warranty:N"), alt.Tooltip("after_vat:Q", title="After VAT (USD)", format=",.2f")]
            )
            .properties(width="container", height=300)
        )
        st.altair_chart(chart, use_container_width=True)
    else:
        st.bar_chart(base.set_index("Warranty")["after_vat"])


def render_daily_trend(tables: Dict[str, pd.DataFrame]):
    st.markdown("### 📅 Daily Trend (After VAT excl CC)")
    df = tables.get("Daily_Revenue")
    if df is None or df.empty:
        st.info("Daily_Revenue sheet not found.")
        return
    if not {"Date", "After VAT (excl CC)"}.issubset(df.columns):
        st.info("Daily_Revenue does not have expected columns.")
        return

    d = df.copy()
    d["Date"] = pd.to_datetime(d["Date"], errors="coerce")
    d = d.dropna(subset=["Date"]).sort_values("Date").set_index("Date")[["After VAT (excl CC)"]]
    st.line_chart(d)


def render_yoy_views(tables: Dict[str, pd.DataFrame]):
    st.markdown("### 📊 YoY Comparison")

    # ---------- YoY_Monthly ----------
    ym = tables.get("YoY_Monthly")
    if ym is not None and not ym.empty:
        dfm = ym.copy()
        dfm.columns = [str(c).strip() for c in dfm.columns]

        metric_col = "Metric" if "Metric" in dfm.columns else dfm.columns[0]
        if "Value" in dfm.columns:
            value_col = "Value"
        else:
            num_cols = [c for c in dfm.columns if pd.api.types.is_numeric_dtype(dfm[c])]
            value_col = num_cols[0] if num_cols else dfm.columns[1]

        def _to_float(x):
            if x is None or (isinstance(x, float) and pd.isna(x)):
                return None
            s = str(x).strip().replace(" ", "").replace(",", "")
            if s.endswith("%"):
                s = s[:-1]
            try:
                return float(s)
            except Exception:
                return None

        def _pct_display(v: Optional[float]) -> str:
            if v is None:
                return ""
            v = float(v)
            # Interpret absolute values >1 as whole-percent entries from sheet
            if abs(v) > 1:
                return f"{v:.2f}%"
            return f"{v*100:.2f}%"

        def _pick_raw(label):
            s = dfm.loc[dfm[metric_col].astype(str).str.strip().eq(label), value_col]
            return None if s.empty else s.iloc[0]

        cur_proj     = _to_float(_pick_raw("Projected Revenue (After VAT, excl CC) – Current"))
        prev_actual  = _to_float(_pick_raw("Actual Revenue (After VAT, excl CC) – Previous Period"))
        delta_money  = _to_float(_pick_raw("Δ vs Previous Period"))
        pct_vs_prev  = _to_float(_pick_raw("% vs Previous Period"))

        c1, c2, c3, c4 = st.columns(4)
        if cur_proj    is not None: c1.metric("Current Projected (After VAT, excl CC)", fmt_money(cur_proj))
        if prev_actual is not None: c2.metric("Prev Period Actual (After VAT, excl CC)", fmt_money(prev_actual))
        if delta_money is not None: c3.metric("Δ vs Prev", fmt_money(delta_money))
        if pct_vs_prev is not None: c4.metric("% vs Prev", _pct_display(pct_vs_prev))

        # Show tidy table
        is_pct_row = dfm[metric_col].astype(str).str.contains(r"margin|% vs|%$", case=False, regex=True)
        out = dfm[[metric_col, value_col]].copy()
        out["Value (display)"] = [
            fmt_pct_metric(v) if is_pct_row.iloc[i] else fmt_money(v)
            for i, v in enumerate(out[value_col].values)
        ]
        st.dataframe(out.rename(columns={metric_col: "Metric"})[[ "Metric", "Value (display)"]], use_container_width=True)
    else:
        st.info("YoY_Monthly sheet not found.")

    # ---------- YoY_Daily ----------
    yd = tables.get("YoY_Daily")
    if yd is None or yd.empty:
        st.info("YoY_Daily sheet not found.")
        return

    st.write("**Daily Comparison**")
    d = yd.copy()
    d.columns = [str(c).strip() for c in d.columns]

    # Identify columns
    day_col = "Day" if "Day" in d.columns else d.columns[0]
    cur_col = next((c for c in d.columns if c.lower().startswith("current") and "after vat" in c.lower()), None)
    prev_col = next((c for c in d.columns if (c.lower().startswith("prev") or "prevyear" in c.lower()) and "after vat" in c.lower()), None)

    # If structure unknown, just show table
    if cur_col is None or prev_col is None:
        st.dataframe(d, use_container_width=True)
        return

    # Ensure numeric & sort by day
    d[day_col] = pd.to_numeric(d[day_col], errors="coerce")
    d[cur_col] = pd.to_numeric(d[cur_col], errors="coerce")
    d[prev_col] = pd.to_numeric(d[prev_col], errors="coerce")
    d = d.dropna(subset=[day_col]).sort_values(day_col)

    # Safe deltas and % vs prev
    eps = 1e-6
    d["Delta"] = d[cur_col] - d[prev_col]
    d["% vs Prev"] = (d[cur_col] - d[prev_col]) / d[prev_col].where(d[prev_col].abs() > eps, np.nan)

    # Chart
    line_cols = [c for c in [cur_col, prev_col] if c in d.columns]
    if line_cols:
        chart_df = d[[day_col] + line_cols].melt(id_vars=[day_col], var_name="Series", value_name="Value")
        if alt:
            chart = (
                alt.Chart(chart_df)
                .mark_line(point=True)
                .encode(
                    x=alt.X(f"{day_col}:Q", title="Day of Month"),
                    y=alt.Y("Value:Q", title="After VAT (USD)"),
                    color="Series:N",
                    tooltip=[alt.Tooltip(f"{day_col}:Q", title="Day"),
                             alt.Tooltip("Series:N"),
                             alt.Tooltip("Value:Q", title="After VAT (USD)", format=",.2f")],
                )
                .properties(width="container", height=280)
            )
            st.altair_chart(chart, use_container_width=True)
        else:
            wide = d.set_index(day_col)[line_cols]
            st.line_chart(wide)

    # Table
    show = d[[day_col, cur_col, prev_col, "Delta", "% vs Prev"]].copy()
    show[cur_col] = show[cur_col].apply(fmt_money)
    show[prev_col] = show[prev_col].apply(fmt_money)
    show["Delta"] = show["Delta"].apply(lambda v: "" if pd.isna(v) else fmt_money(v))
    show["% vs Prev"] = show["% vs Prev"].apply(lambda v: "" if pd.isna(v) else f"{v * 100:.2f}%")
    show = show.rename(
        columns={
            day_col: "Day",
            cur_col: "Current After VAT (excl CC)",
            prev_col: "PrevYear After VAT (excl CC)",
        }
    )
    st.dataframe(show, use_container_width=True)


def render_pl_forecast(tables: Dict[str, pd.DataFrame]):
    st.markdown("### 💼 P&L Forecast (After VAT)")

    # Candidate sheet names
    sheet_candidates = ["P&L_Forecast", "P&L Forecast", "PL_Forecast", "P&L", "PL"]
    df = None
    for k in sheet_candidates:
        if k in tables and tables[k] is not None and not tables[k].empty:
            df = tables[k]
            break
    if df is None or df.empty:
        st.info("P&L_Forecast sheet not found.")
        return

    # Normalize headers
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Detect label & value columns robustly
    label_candidates = [c for c in df.columns if c.lower() in ("line", "metric", "label", "item", "name")]
    if label_candidates:
        label_col = label_candidates[0]
    else:
        nonnum = [c for c in df.columns if df[c].dtype == "object"]
        label_col = nonnum[0] if nonnum else df.columns[0]

    value_col = next((c for c in df.columns if c.lower() == "value"), None)
    if value_col is None:
        numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        if numeric_cols:
            value_col = numeric_cols[0]
        else:
            # fallback: choose the column with most numeric-like entries
            best_c, best_cnt = None, -1
            for c in df.columns:
                coerced = pd.to_numeric(df[c], errors="coerce")
                cnt = coerced.notna().sum()
                if cnt > best_cnt:
                    best_c, best_cnt = c, cnt
            if best_c is None or best_cnt <= 0:
                st.warning("P&L_Forecast: couldn’t identify a numeric value column.")
                st.dataframe(df, use_container_width=True)
                return
            df[best_c] = pd.to_numeric(df[best_c], errors="coerce")
            value_col = best_c

    labels_norm = df[label_col].astype(str).str.replace("–", "-", regex=False).str.strip()

    def _pick(regex_pat: str) -> Optional[float]:
        pat = re.compile(regex_pat, re.IGNORECASE)
        mask = labels_norm.apply(lambda x: bool(pat.search(x)))
        m = df.loc[mask, value_col]
        return float(m.iloc[0]) if not m.empty and pd.notna(m.iloc[0]) else None

    # Pull the numbers for KPI tiles & ratios
    total_rev_f   = (_pick(r"total\s+revenue.*forecast") or
                     _pick(r"revenue.*forecast.*\(after\s*vat\)") or
                     _pick(r"total.*forecast.*revenue"))
    admin_cost_f  = (_pick(r"admin(istration)?\s+costs?.*forecast") or _pick(r"overhead.*forecast"))
    op_profit_f   = (_pick(r"operating\s+profit.*forecast") or _pick(r"op\s*profit.*forecast"))
    op_margin_pct = (_pick(r"operating\s+margin.*forecast") or _pick(r"op.*margin.*forecast"))

    service_f = _pick(r"revenue.*service.*forecast.*after\s*vat.*excl.*cc")
    cc_f      = _pick(r"revenue.*call\s*center.*forecast")
    total_rev_actual = _pick(r"total\s+revenue.*actual\s+to\s+date")
    active_days = _pick(r"forecast\s+active\s+days")
    month_days  = _pick(r"forecast\s+month\s+days")

    # Tidy preview table (formatted)
    is_pct = labels_norm.str.contains(r"margin|%", case=False, regex=True)
    table = df[[label_col, value_col]].copy()
    table["Value (display)"] = [
        (fmt_pct_metric(v) if is_pct.iloc[i] else fmt_money(v))
        for i, v in enumerate(table[value_col].values)
    ]
    st.dataframe(table[[label_col, "Value (display)"]].rename(columns={label_col: "Metric"}), use_container_width=True)

    # KPI tiles
    c1, c2, c3, c4 = st.columns(4)
    if total_rev_f is not None:   c1.metric("Total Revenue (Forecast)", fmt_money(total_rev_f))
    if admin_cost_f is not None:  c2.metric("Admin Costs (Forecast)", fmt_money(admin_cost_f))
    if op_profit_f is not None:   c3.metric("Operating Profit (Forecast)", fmt_money(op_profit_f))
    if op_margin_pct is not None: c4.metric("Operating Margin % (Forecast)", fmt_pct_metric(op_margin_pct))

    # Ratios
    ratios = []
    if (total_rev_f not in (None, 0)) and (admin_cost_f is not None):
        ratios.append(("Admin Costs / Total Revenue", fmt_pct_ratio(admin_cost_f / total_rev_f)))
    if (total_rev_f not in (None, 0)) and (service_f is not None):
        ratios.append(("Service Share of Revenue (Forecast)", fmt_pct_ratio(service_f / total_rev_f)))
    if (total_rev_f not in (None, 0)) and (cc_f is not None):
        ratios.append(("Call Center Share of Revenue (Forecast)", fmt_pct_ratio(cc_f / total_rev_f)))
    if (total_rev_f not in (None, 0)) and (total_rev_actual is not None):
        ratios.append(("Progress vs Forecast (Actual / Forecast)", fmt_pct_ratio(total_rev_actual / total_rev_f)))
    if (active_days not in (None, 0)) and (month_days is not None):
        ratios.append(("Forecast Factor (Month / Active Days)", f"{float(month_days)/float(active_days):.2f}×"))
    if (op_profit_f is not None) and (total_rev_f not in (None, 0)):
        if op_profit_f < 0:
            gap = abs(op_profit_f)
            ratios.append(("Breakeven Gap (to 0 profit)", fmt_money(gap)))
            ratios.append(("Breakeven Gap as % of Revenue", fmt_pct_ratio(gap / total_rev_f)))
        else:
            ratios.append(("Breakeven Status", "At/Above Breakeven"))

    if ratios:
        st.markdown("#### 📐 Key Ratios")
        st.dataframe(pd.DataFrame(ratios, columns=["Ratio", "Value"]), use_container_width=True)
    else:
        st.info("No ratios could be computed from this P&L sheet.")


def render_yoy_warranty(tables: Dict[str, pd.DataFrame]):
    st.markdown("### 🧩 YoY Warranty Mix (Pie)")
    df = tables.get("YoY_Warranty")
    if df is None or df.empty:
        st.info("YoY_Warranty sheet not found.")
        return

    d = df.copy()
    d.columns = [str(c).strip() for c in d.columns]
    if "Warranty" not in d.columns:
        st.dataframe(d, use_container_width=True)
        return

    # Keep only G1/G2/G3
    d = d[d["Warranty"].isin(["G1", "G2", "G3"])]

    # Expected columns (tolerant)
    amt_cur   = "Current After VAT (excl CC)"
    amt_proj  = "Current Projected After VAT (excl CC)"
    amt_prev  = "PrevYear After VAT (excl CC)"
    sh_cur    = "Current Share % (Actual)"
    sh_proj   = "Current Share % (Projected)"
    sh_prev   = "PrevYear Share %"

    def ensure_shares(df_in, amt_col, share_col):
        out = df_in.copy()
        if amt_col not in out.columns:
            return out
        if share_col not in out.columns:
            total = pd.to_numeric(out[amt_col], errors="coerce").sum(skipna=True)
            out[share_col] = np.where(total != 0, (pd.to_numeric(out[amt_col], errors="coerce") / total) * 100.0, np.nan)
        return out

    for (a, s) in [(amt_cur, sh_cur), (amt_proj, sh_proj), (amt_prev, sh_prev)]:
        d = ensure_shares(d, a, s)

    def make_pie(df_in, value_col, share_col, title):
        if (value_col not in df_in.columns) or (share_col not in df_in.columns):
            return None
        tmp = df_in[["Warranty", value_col, share_col]].copy()
        tmp = tmp.rename(columns={value_col: "Amount", share_col: "Share"})
        tmp["Amount_f"] = tmp["Amount"].apply(fmt_money)
        tmp["Share_f"] = tmp["Share"].apply(fmt_pct_metric)
        tmp["Title"] = title
        return tmp

    datasets = [
        make_pie(d, amt_cur,  sh_cur,  "Current – Actual"),
        make_pie(d, amt_proj, sh_proj, "Current – Projected"),
        make_pie(d, amt_prev, sh_prev, "Previous Year – Actual"),
    ]
    pies = [p for p in datasets if p is not None]
    if not pies:
        st.warning("YoY_Warranty: columns not found to build pies.")
        st.dataframe(d, use_container_width=True)
        return

    cols = st.columns(len(pies))
    for col, data in zip(cols, pies):
        with col:
            st.caption(f"**{data['Title'].iat[0]}**")
            if alt:
                chart = (
                    alt.Chart(data)
                    .mark_arc()
                    .encode(
                        theta=alt.Theta("Share:Q", title="Share", stack=True),
                        color=alt.Color("Warranty:N", title="Warranty"),
                        tooltip=[alt.Tooltip("Warranty:N"),
                                 alt.Tooltip("Share:Q", title="Share", format=".2f"),
                                 alt.Tooltip("Amount:Q", title="Amount (USD)", format=",.2f")],
                    )
                    .properties(width="container", height=260)
                )
                st.altair_chart(chart, use_container_width=True)
            st.dataframe(
                data[["Warranty", "Share_f", "Amount_f"]].rename(columns={"Share_f": "Share", "Amount_f": "Amount (USD)"}),
                use_container_width=True
            )

# ============================ RUN JOB =============================
if run:
    if not cur_files:
        st.error("Please upload at least one SAP Excel file for the current period.")
        st.stop()
    if yoy_mode == "Upload Previous-Year File" and not prev_file_obj:
        st.error("Baseline Mode is 'Upload Previous-Year File' but no baseline file was provided.")
        st.stop()

    with st.spinner("Processing…"):
        tmp_dir = tempfile.mkdtemp()

        cur_paths = _save_uploaded_files(cur_files, tmp_dir)

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

# ============================ DOWNLOAD ============================
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

# ====================== ADDITIONAL ANALYSIS VIEWS =================
st.markdown("---")
st.subheader("📊 Additional Analysis Views")

col_a, col_b, col_c, col_d, col_e, col_f, col_g = st.columns(7)
btn_rev   = col_a.button("📈 Revenue Charts",       key="btn_rev")
btn_corr  = col_b.button("🏭 By Correspondent",     key="btn_corr")
btn_warr  = col_c.button("🧩 Warranty Structure",   key="btn_warr")
btn_daily = col_d.button("📅 Daily Trend",          key="btn_daily")
btn_yoy   = col_e.button("📊 YoY Compare",          key="btn_yoy")
btn_pl    = col_f.button("💼 P&L",                  key="btn_pl")
btn_yoyw  = col_g.button("🧩 YoY Warranty (Pie)",   key="btn_yoyw")

if any([btn_rev, btn_corr, btn_warr, btn_daily, btn_yoy, btn_pl, btn_yoyw]):
    rp = st.session_state.get("report_path")
    if not rp or not os.path.isfile(rp):
        st.warning("No generated report found. Please run the forecast first.")
    else:
        tables = _read_report_tables(rp)
        vat_rate_eff = st.session_state.get("vat_rate_eff", float(vat_percent) / 100.0)

        if btn_rev:
            render_revenue_charts(tables, vat_rate_eff)
        if btn_corr:
            render_by_correspondent(tables, vat_rate_eff)
        if btn_warr:
            render_warranty_share(tables, vat_rate_eff)
        if btn_daily:
            render_daily_trend(tables)
        if btn_yoy:
            render_yoy_views(tables)
        if btn_pl:
            render_pl_forecast(tables)
        if btn_yoyw:
            render_yoy_warranty(tables)

# ============================ FOOTER =============================
if not st.session_state.get("report_ready"):
    st.info("👆 Upload your SAP Excel file(s), adjust settings in the sidebar, then click **Run Forecast**.")



