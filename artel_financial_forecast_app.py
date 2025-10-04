import io
import math
import calendar
from datetime import date, datetime
from typing import Tuple, Dict

import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------
# Global defaults (single source of truth)
# -----------------------------
VAT_RATE_DEFAULT = 0.12  # 12%
VAT_MODE_DEFAULT = "extract"  # "extract" (net VAT out of gross) or "add" (subtract VAT from amount)
DATE_SOURCE_DEFAULT = "Data of Document"  # or "Data of transaction"
ONLY_USD_DEFAULT = True
EXCLUDE_SUNDAYS_DEFAULT = True
FORECAST_NONEMPTY_ONLY_DEFAULT = True

# Business rules
G1_WARRANTY_LABEL = "G1"
CALL_LABELS = {"ВЫЗОВ", "CALL", "CALL CENTER", "TRANSPORT", "TRANSPORTATION", "ТРАНСПОРТ", "ВЫЗОВ МАСТЕРА"}
EXCEPTION_CORRESPONDENTS_REVENUE = {
    "2100000175",
    "2100000170",
    "2100000226",
    "2100000229",
}

# Canonical column names we expect after normalization
REQUIRED_COLS = {
    "Correspondent",
    "Number of Documents",
    "Data of Document",
    "Data of transaction",
    "Number of request",
    "Material/SAP Code",
    "Name",
    "Qty",
    "Measurement",
    "amount",  # normalized lower
    "Currency",
    "Amount in USD",
    "UZS",
    "Warranty",
}


# -----------------------------
# Utilities
# -----------------------------

def _normalize_vat_mode(v) -> str:
    v = str(v).strip().lower()
    if v in {"extract", "exclusive", "net", "net_of_vat", "netto"}:
        return "extract"
    if v in {"add", "inclusive", "gross", "brutto"}:
        return "add"
    return VAT_MODE_DEFAULT


def _calc_after_vat(amount_vat_incl: float, vat_rate: float, vat_mode: str) -> float:
    """Return amount after VAT according to mode.
    - extract: treat input as VAT-inclusive and extract net = amount/(1+rate)
    - add: treat input as VAT-inclusive gross and subtract VAT portion: amount*(1-rate)
      (kept for backwards compatibility with earlier spreadsheets)
    """
    mode = _normalize_vat_mode(vat_mode)
    if pd.isna(amount_vat_incl):
        return 0.0
    if mode == "extract":
        return float(amount_vat_incl) / (1.0 + float(vat_rate))
    else:
        return float(amount_vat_incl) * (1.0 - float(vat_rate))


def _coerce_datetime(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, (pd.Timestamp, datetime, date)):
        return pd.to_datetime(x)
    # try parse str/number (excel serial etc.)
    try:
        return pd.to_datetime(x, errors="coerce")
    except Exception:
        return pd.NaT


def _standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # Trim and lower headers, then map some known variants
    rename_map = {c: c.strip() for c in df.columns}
    df = df.rename(columns=rename_map)

    # Ensure amount synonyms map consistently
    syns = {
        "Amount": "amount",
        "AMOUNT": "amount",
        "Сумма": "amount",
        "mount": "amount",
        "amount": "amount",
    }
    for s, t in syns.items():
        if s in df.columns and t not in df.columns:
            df = df.rename(columns={s: t})

    # Normalize object/string columns
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].astype(str).str.strip()

    # Coerce numeric-like
    for c in ["Qty", "amount", "Amount in USD", "UZS"]:
        if c in df.columns:
            df[c] = (pd.to_numeric(df[c].astype(str).str.replace(" ", "").str.replace(",", "."), errors="coerce")
                      .fillna(0.0))

    # Coerce dates
    for c in ["Data of Document", "Data of transaction"]:
        if c in df.columns:
            df[c] = df[c].apply(_coerce_datetime)

    return df


def _duplicate_requests_block(df: pd.DataFrame) -> pd.DataFrame:
    """Return a small block listing duplicate 'Number of request' rows (first 100)."""
    dups = pd.DataFrame()
    if "Number of request" in df.columns:
        nr = df["Number of request"].astype(str).str.strip()
        valid = nr.ne("") & nr.notna()
        dup_mask = nr[valid].duplicated(keep=False)
        dups = df.loc[valid].loc[dup_mask].sort_values("Number of request").head(100).copy()
        if not dups.empty:
            dups["Issue"] = "Duplicate Number of request"
    return dups


def _is_call_row(name_value: str) -> bool:
    if not isinstance(name_value, str):
        return False
    nv = name_value.strip().upper()
    return nv in {s.upper() for s in CALL_LABELS}


# -----------------------------
# Core business aggregation
# -----------------------------

def compute_kpis(
    df_raw: pd.DataFrame,
    target_month: date,
    *,
    date_source: str = DATE_SOURCE_DEFAULT,
    vat_rate: float = VAT_RATE_DEFAULT,
    vat_mode: str = VAT_MODE_DEFAULT,
    only_usd: bool = ONLY_USD_DEFAULT,
    exclude_sundays: bool = EXCLUDE_SUNDAYS_DEFAULT,
    forecast_nonempty_only: bool = FORECAST_NONEMPTY_ONLY_DEFAULT,
    call_center_revenue: float = 0.0,
    admin_forecast: float = 0.0,
) -> Tuple[pd.DataFrame, Dict[str, float], pd.DataFrame]:
    """Return (detail_df, metrics_dict, dq_block) for UI rendering."""
    vat_mode = _normalize_vat_mode(vat_mode)

    df = _standardize_columns(df_raw)

    # Optionally filter to USD only
    if only_usd and "Currency" in df.columns:
        df = df[df["Currency"].str.upper().eq("USD") | df["Currency"].eq(1)]  # some exports have 1 for USD

    # Select date column
    date_col = date_source if date_source in df.columns else (
        "Data of Document" if "Data of Document" in df.columns else "Data of transaction"
    )

    # Keep only target month
    df["_date"] = df[date_col]
    df = df[df["_date"].dt.to_period("M") == pd.Period(target_month, freq="M")]

    # Tag CALL rows and G1 warranty
    df["_is_call"] = df["Name"].apply(_is_call_row) if "Name" in df.columns else False
    df["_is_g1"] = df["Warranty"].astype(str).str.upper().eq(G1_WARRANTY_LABEL)

    # Identify G1 transport to subtract
    g1_transport = df.loc[df["_is_call"] & df["_is_g1"], "amount"].sum()

    # Exception correspondents: always revenue even if call
    if "Correspondent" in df.columns:
        mask_exc = df["Correspondent"].astype(str).isin(EXCEPTION_CORRESPONDENTS_REVENUE)
    else:
        mask_exc = pd.Series(False, index=df.index)

    # Total before adjustments
    total_before_adj = float(df["amount"].sum())

    # Net revenue (subtract G1 Transport except exception correspondents)
    mask_g1_call = df["_is_call"] & df["_is_g1"] & ~mask_exc
    less_g1_transport = float(df.loc[mask_g1_call, "amount"].sum())
    net_revenue = total_before_adj - less_g1_transport

    # After VAT (core), then add call center
    after_vat_excl_cc = _calc_after_vat(net_revenue, vat_rate, vat_mode)
    revenue_after_vat = after_vat_excl_cc + float(call_center_revenue)

    # Simple daily forecast: average per active day * count of target days
    if forecast_nonempty_only:
        # active days are days present in data (non-zero amount)
        daily = df.groupby(df["_date"].dt.date)["amount"].sum()
        base_days = (daily > 0).sum()
        base_amount = daily.sum()
    else:
        # use all calendar days up to 'today in month' or full month if month is in the past
        daily = df.groupby(df["_date"].dt.date)["amount"].sum()
        base_days = calendar.monthrange(target_month.year, target_month.month)[1]
        base_amount = daily.sum()

    # target month days respecting Sunday exclusion
    _, mdays = calendar.monthrange(target_month.year, target_month.month)
    all_dates = pd.date_range(date(target_month.year, target_month.month, 1), periods=mdays, freq="D")
    if exclude_sundays:
        month_days = int((all_dates.weekday != 6).sum())
        # also filter base days by Sunday exclusion
        base_active_dates = pd.to_datetime(list(daily.index))
        base_days = int(((pd.Series(base_active_dates).dt.weekday != 6)).sum()) if forecast_nonempty_only else month_days
    else:
        month_days = mdays

    avg_per_day = (base_amount / base_days) if base_days else 0.0
    forecast_amount = avg_per_day * month_days

    # Admin manual overlay
    forecast_total = forecast_amount + float(admin_forecast)

    # Build metrics
    metrics = {
        "Total Revenue (before adj)": total_before_adj,
        "Less G1 Transport": less_g1_transport,
        "Net Revenue": net_revenue,
        "Less VAT": net_revenue - after_vat_excl_cc,
        "Revenue After VAT (excl CC)": after_vat_excl_cc,
        "Call Center": float(call_center_revenue),
        "Revenue After VAT (incl CC)": revenue_after_vat,
        "Forecast base active days": int(base_days),
        "Forecast month days": int(month_days),
        "Forecast (model)": float(forecast_amount),
        "Admin forecast adj": float(admin_forecast),
        "Forecast (total)": float(forecast_total),
    }

    dq = _duplicate_requests_block(df)

    return df, metrics, dq


# -----------------------------
# Streamlit UI
# -----------------------------

def _format_usd(x: float) -> str:
    return f"${x:,.2f}"


def main():
    st.set_page_config(page_title="Artel Financial Forecast", layout="wide")
    st.title("Artel Financial Forecast")

    # Sidebar controls
    with st.sidebar:
        st.header("Settings")
        vat_rate = st.number_input("VAT rate (e.g., 0.12 for 12%)", min_value=0.0, max_value=1.0, value=VAT_RATE_DEFAULT, step=0.01)
        vat_mode = _normalize_vat_mode(st.selectbox("VAT mode", options=["extract", "add", "exclusive", "inclusive", "gross", "net"], index=0))
        date_source = st.selectbox("Date source", options=["Data of Document", "Data of transaction"], index=0)
        only_usd = st.checkbox("Only USD rows", value=ONLY_USD_DEFAULT)
        exclude_sundays = st.checkbox("Exclude Sundays (forecast)", value=EXCLUDE_SUNDAYS_DEFAULT)
        nonempty_only = st.checkbox("Forecast base = non-empty days only", value=FORECAST_NONEMPTY_ONLY_DEFAULT)

        call_center = st.number_input("Call center revenue (USD)", value=0.0, step=100.0)
        admin_forecast = st.number_input("Admin forecast adjustment (USD)", value=0.0, step=100.0)

        month = st.date_input(
            "Target month",
            value=date.today().replace(day=1),
            min_value=date(2000, 1, 1),
            max_value=date(2100, 12, 31),
        )

    st.markdown(
        """
        **Instructions**
        1. Upload the raw SAP Excel export (sheet with the columns used below).
        2. Verify the sidebar settings match your local script.
        3. Compare the metrics. Use the *Debug* panel to confirm effective configuration.
        """
    )

    uploaded = st.file_uploader("Upload SAP Excel (.xlsx, .xls)", type=["xlsx", "xls"])

    if uploaded is None:
        st.info("Upload an Excel export to begin.")
        st.stop()

    # Read Excel (first sheet by default)
    try:
        df_raw = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")
        st.stop()

    detail_df, metrics, dq = compute_kpis(
        df_raw,
        target_month=month,
        date_source=date_source,
        vat_rate=vat_rate,
        vat_mode=vat_mode,
        only_usd=only_usd,
        exclude_sundays=exclude_sundays,
        forecast_nonempty_only=nonempty_only,
        call_center_revenue=call_center,
        admin_forecast=admin_forecast,
    )

    # KPIs
    st.subheader("KPIs")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Revenue (before adj)", _format_usd(metrics["Total Revenue (before adj)"]))
        st.metric("Net Revenue", _format_usd(metrics["Net Revenue"]))
        st.metric("Revenue After VAT (excl CC)", _format_usd(metrics["Revenue After VAT (excl CC)"]))
    with col2:
        st.metric("Less G1 Transport", _format_usd(metrics["Less G1 Transport"]))
        st.metric("Less VAT", _format_usd(metrics["Less VAT"]))
        st.metric("Call Center", _format_usd(metrics["Call Center"]))
    with col3:
        st.metric("Revenue After VAT (incl CC)", _format_usd(metrics["Revenue After VAT (incl CC)"]))
        st.metric("Forecast (model)", _format_usd(metrics["Forecast (model)"]))
        st.metric("Forecast (total)", _format_usd(metrics["Forecast (total)"]))

    # Tables
    with st.expander("Reconciliation table", expanded=True):
        rec = pd.DataFrame(
            {
                "Metric": list(metrics.keys()),
                "Value (USD)": [
                    metrics["Total Revenue (before adj)"],
                    metrics["Less G1 Transport"],
                    metrics["Net Revenue"],
                    metrics["Less VAT"],
                    metrics["Revenue After VAT (excl CC)"],
                    metrics["Call Center"],
                    metrics["Revenue After VAT (incl CC)"],
                    metrics["Forecast base active days"],
                    metrics["Forecast month days"],
                    metrics["Forecast (model)"],
                    metrics["Admin forecast adj"],
                    metrics["Forecast (total)"],
                ],
            }
        )
        st.dataframe(rec, use_container_width=True)

    with st.expander("Detail (filtered to target month)"):
        show_cols = [c for c in [
            "Correspondent",
            "Number of Documents",
            "Data of Document",
            "Data of transaction",
            "Number of request",
            "Material/SAP Code",
            "Name",
            "Qty",
            "Measurement",
            "amount",
            "Currency",
            "Amount in USD",
            "UZS",
            "Warranty",
        ] if c in detail_df.columns]
        st.dataframe(detail_df[show_cols], use_container_width=True)

    with st.expander("Data Quality checks"):
        if dq.empty:
            st.success("No duplicate 'Number of request' values detected (first 100).")
        else:
            st.warning("Duplicates detected for 'Number of request'. Review below (first 100 shown):")
            st.dataframe(dq, use_container_width=True)

    with st.expander("Debug (effective config)"):
        st.json(
            {
                "vat_rate": vat_rate,
                "vat_mode_normalized": vat_mode,
                "date_source": date_source,
                "only_usd": only_usd,
                "exclude_sundays": exclude_sundays,
                "forecast_nonempty_only": nonempty_only,
                "month": month.strftime("%Y-%m"),
            }
        )


if __name__ == "__main__":
    main()
