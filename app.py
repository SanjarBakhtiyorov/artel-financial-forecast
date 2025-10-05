# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
import calendar
import io

# ---------------- CONFIG ----------------
VAT_RATE = 0.12
SPECIAL_CORR = [2100000175, 2100000170, 2100000226, 2100000229]

RUS_TO_ENG_COLS = {
    "Кредитор": "Correspondent",
    "Номер документа": "Number of Documents",
    "Дата документа": "Data of Document",
    "Дата проводки": "Data of transaction",
    "Номер заявки": "Number of request",
    "Материал/Услуга": "Material/SAP Code",
    "Название": "Name",
    "Количество": "Qty",
    "Единица измерения": "Measurement",
    "Сумма": "Amount",
    "Валюта": "Currency",
    "Гарантия": "Warranty",
}

# ---------------- HELPERS ----------------
def translate_columns(df):
    out = df.copy()
    for c in df.columns:
        if c.strip() in RUS_TO_ENG_COLS:
            out = out.rename(columns={c: RUS_TO_ENG_COLS[c.strip()]})
    return out

def compute_g1_transport(df):
    name_is_call = df["Name"].str.contains("ВЫЗОВ", case=False, na=False)
    is_g1 = df["Warranty"].str.upper().eq("G1")
    is_special = df["Correspondent"].isin(SPECIAL_CORR)
    df["g1_transport"] = np.where(name_is_call & is_g1 & (~is_special), df["Amount"], 0.0)
    return df

def calc_forecast(df, after_vat_excl_cc, date_source="Data of Document"):
    ds = df[date_source].dropna()
    if ds.empty:
        return 0.0
    unique_days = ds.dt.date.nunique()
    mode_period = ds.dt.to_period("M").mode()
    if len(mode_period) == 0:
        return 0.0
    y, m = mode_period.iloc[0].year, mode_period.iloc[0].month
    days_in_month = calendar.monthrange(y, m)[1]
    return round((after_vat_excl_cc / unique_days) * days_in_month, 2)

def build_summary(df, call_center_revenue):
    total_before_adj = df["Amount"].sum()
    less_g1_transport = df["g1_transport"].sum()
    net_vat_incl = total_before_adj + call_center_revenue - less_g1_transport
    net_vat_incl_excl_cc = total_before_adj - less_g1_transport

    less_vat = net_vat_incl - (net_vat_incl / (1 + VAT_RATE))
    revenue_after_vat = net_vat_incl / (1 + VAT_RATE)
    after_vat_excl_cc = net_vat_incl_excl_cc / (1 + VAT_RATE)

    forecast = calc_forecast(df, after_vat_excl_cc)
    g3_amount = df.loc[df["Warranty"].eq("G3"), "Amount"].sum()
    g3_after_vat = g3_amount / (1 + VAT_RATE)
    g3_share = (g3_after_vat / after_vat_excl_cc * 100) if after_vat_excl_cc != 0 else 0.0

    summary = pd.DataFrame({
        "Metric": [
            "Total Revenue (before adj)",
            "Call Center",
            "Less G1 Transport",
            "Net Revenue (VAT incl.)",
            "Less VAT (12%)",
            "Revenue After VAT",
            "Forecast (After VAT, excl CC)",
            "G3 Share (%)"
        ],
        "Value": [
            round(total_before_adj, 2),
            round(call_center_revenue, 2),
            round(less_g1_transport, 2),
            round(net_vat_incl, 2),
            round(less_vat, 2),
            round(revenue_after_vat, 2),
            round(forecast, 2),
            round(g3_share, 2)
        ]
    })
    return summary, revenue_after_vat, forecast, g3_share, after_vat_excl_cc

def build_by_correspondent(df):
    g = (
        df.groupby("Correspondent", dropna=False)
          .agg(total_amount=("Amount", "sum"))
          .reset_index()
    )
    g["total_amount"] = g["total_amount"].round(2)
    return g.sort_values("total_amount", ascending=False)

# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="Monthly Revenue Report", layout="wide")

st.title("📊 Monthly Revenue Report (VAT Extracted)")

uploaded = st.file_uploader("Upload SAP Excel (.xls/.xlsx)", type=["xls", "xlsx"])

if uploaded:
    df = pd.read_excel(uploaded)
    df = translate_columns(df)

    # Data prep
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df["Correspondent"] = pd.to_numeric(df["Correspondent"], errors="coerce")
    df["Warranty"] = df["Warranty"].astype(str).str.upper().str.strip()
    df["Name"] = df["Name"].astype(str).str.strip()
    for c in ["Data of Document", "Data of transaction"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    df = compute_g1_transport(df)

    # Month Selector
    if "Data of Document" in df.columns:
        df["month_period"] = df["Data of Document"].dt.to_period("M")
        months = sorted(df["month_period"].dropna().unique().astype(str))
        if len(months) > 0:
            selected_month = st.selectbox("Select month to analyze:", months, index=len(months)-1)
            df = df[df["month_period"] == selected_month]
            st.success(f"✅ Filtered data for {selected_month}")
        else:
            st.warning("⚠️ No valid dates found in Data of Document column.")
    else:
        st.warning("⚠️ Column 'Data of Document' not found — cannot filter by month.")

    call_center_revenue = st.number_input("Enter Call Center revenue (USD, VAT-incl.)", min_value=0.0, step=100.0)

    if st.button("Run Report"):
        summary, after_vat, forecast, g3_share, after_vat_excl_cc = build_summary(df, call_center_revenue)
        st.success("✅ Calculation complete!")

        st.subheader("Summary")
        st.dataframe(summary, use_container_width=True)

        col1, col2, col3 = st.columns(3)
        col1.metric("Revenue After VAT", f"${after_vat:,.2f}")
        col2.metric("Forecast (After VAT, excl CC)", f"${forecast:,.2f}")
        col3.metric("G3 Share (%)", f"{g3_share:.2f}%")

        # --- Chart 1: Forecast vs Actual ---
        st.subheader("📈 Forecast vs Actual (After VAT, excl Call Center)")
        chart_df = pd.DataFrame({
            "Metric": ["Actual (After VAT excl CC)", "Forecast (Month-End)"],
            "Value": [after_vat_excl_cc, forecast]
        })
        st.bar_chart(chart_df.set_index("Metric"))

        # --- Chart 2: Revenue by Correspondent ---
        st.subheader("🏭 Revenue by Correspondent (After VAT)")
        by_corr = build_by_correspondent(df)
        by_corr["after_vat"] = by_corr["total_amount"] / (1 + VAT_RATE)
        top_corr = by_corr.head(10).set_index("Correspondent")
        st.bar_chart(top_corr["after_vat"])

        # --- Chart 3: Warranty Structure Pie ---
        st.subheader("🧩 Warranty Structure")
        warr_df = df.groupby("Warranty")["Amount"].sum().reset_index()
        warr_df["Amount"] = warr_df["Amount"] / (1 + VAT_RATE)
        warr_df = warr_df[warr_df["Amount"] > 0]
        if not warr_df.empty:
            st.write("G1 vs G3 Share (After VAT)")
            st.dataframe(warr_df)
            st.plotly_chart(
                {
                    "data": [
                        {
                            "labels": warr_df["Warranty"],
                            "values": warr_df["Amount"],
                            "type": "pie"
                        }
                    ],
                    "layout": {"title": "Warranty Share"}
                }
            )

        # --- Download Excel ---
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            summary.to_excel(writer, index=False, sheet_name="Summary")
            df.to_excel(writer, index=False, sheet_name="Detailed")
            by_corr.to_excel(writer, index=False, sheet_name="By_Correspondent")
            warr_df.to_excel(writer, index=False, sheet_name="By_Warranty")
        st.download_button("📥 Download Excel Report", data=out.getvalue(), file_name="monthly_revenue_report.xlsx")

else:
    st.info("Please upload a SAP Excel file to begin.")
