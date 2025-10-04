# monthly_forecast_artel_service_full.py
# -*- coding: utf-8 -*-
import os
import glob
import datetime as dt
from typing import Dict, Optional, Tuple

import numpy as np
import pandas as pd

# ---------------------------- CONSTANTS ----------------------------

VAT_RATE = 0.12                   # default 12%
VAT_MODE = "extract"              # "extract" (net includes VAT) or "exclusive"
DATE_SOURCE = "Data of Document"  # which date column to slice/aggregate on
TOP_NAMES = 10

# G1 "ВЫЗОВ" billed -> NOT deducted
SPECIAL_CORR_DEFAULT = [2100000175, 2100000170, 2100000226, 2100000229]

CORRESPONDENT_MAP_DEFAULT = {
    2100000035: "ЗАВОД КОНДИЦИОНЕРОВ",
    2100000051: "ЗАВОД ХОЛОДИЛЬНИКОВ",
    2100000017: "ЗАВОД СТИРАЛЬНЫХ МАШИН (П)",
    2100000175: "ЗАВОД СТИРАЛЬНЫХ МАШИН (А)",
    2100000003: "ЗАВОД ТЕЛЕВИЗОРОВ",
    2100000000: "ЗАВОД ГАЗОВЫХ ПЛИТ",
    2100000038: "ЗАВОД ЭЛЕКТРИЧЕСКИХ ВОДОН",
    2100000004: "DOMESTIC TRADE (ИМПОРТ)",
    2100000170: "ЗАВОД ПРОМЫШЛЕННЫХ КОНДИЦ",
    2100000226: "ЗАВОД ХОЛОДИЛЬНИКОВ ТЕХНО",
    2100000020: "ЗАВОД ПЫЛЕСОСОВ",
    2100000320: "ЗАВОД МИНИ ПЕЧИ",
    2100000085: "ЗАВОД КУХОННЫХ ВЫТЯЖЕК",
    2100000183: "ЗАВОД КОТЛОВ",
    2100000075: "МБТ МЕЛКАЯ БЫТОВАЯ ТЕХНИК",
    2100000229: "ПРОМЫШЛЕННЫЕ КОНДИЦИОНЕРЫ",
    2100000252: "Витринный холодильник",
    2000002496: "LAMO ELECTROTECH ООО",
    2100000021: "ЗАВОД МИКРОВОЛНОВЫХ ПЕЧЕЙ",
    2100000087: "LIGHTING GOODS ЧП",
    2000002277: "Холодильники Климасан",
    2100000086: "ЗАВОД МОБИЛЬНЫХ ТЕЛЕФОНОВ",
}

RU2EN_COLS = {
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

REQUIRED_COLS = [
    "Correspondent", "Number of Documents", "Data of Document", "Data of transaction",
    "Number of request", "Material/SAP Code", "Name", "Qty", "Measurement",
    "Amount", "Currency", "Warranty"
]

# ---------------------------- UTILITIES ----------------------------

def parse_money(x) -> float:
    if x is None:
        return 0.0
    s = str(x).strip().replace(" ", "")
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return 0.0

def read_excel_any(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        try:
            import xlrd  # noqa
        except Exception as e:
            raise RuntimeError(
                "This .XLS file requires xlrd>=2.0.1. "
                "Add it to requirements.txt or upload .XLSX."
            ) from e
        return pd.read_excel(path, engine="xlrd")
    return pd.read_excel(path, engine="openpyxl")

def translate_columns(df: pd.DataFrame) -> pd.DataFrame:
    # rename RU->EN if present
    ren = {c: RU2EN_COLS[c] for c in df.columns if c in RU2EN_COLS}
    if ren:
        df = df.rename(columns=ren)
    return df

def normalize(df: pd.DataFrame) -> pd.DataFrame:
    df = translate_columns(df)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    out = df.copy()
    out["Correspondent"] = pd.to_numeric(out["Correspondent"], errors="coerce").astype("Int64")
    out["Name"] = out["Name"].astype(str).str.strip()
    out["Warranty"] = out["Warranty"].astype(str).str.upper().str.strip()
    out["Amount"] = pd.to_numeric(out["Amount"], errors="coerce").fillna(0.0)

    for col in ["Data of Document", "Data of transaction"]:
        out[col] = pd.to_datetime(out[col], errors="coerce")

    out["correspondent_name"] = out["Correspondent"].map(CORRESPONDENT_MAP_DEFAULT)

    # Currency sanity
    mask_non_usd = out["Currency"].astype(str).str.upper().ne("USD")
    if mask_non_usd.any():
        # Keep minimal print for Streamlit logs
        print("⚠ Non-USD rows detected (showing first 5):")
        print(out.loc[mask_non_usd, ["Currency", "Amount"]].head(5).to_string(index=False))
    return out

def compute_g1_transport(df: pd.DataFrame, special_corr) -> pd.DataFrame:
    is_call = df["Name"].str.contains("ВЫЗОВ", case=False, na=False)
    is_g1 = df["Warranty"].eq("G1")
    is_special = df["Correspondent"].isin(special_corr)
    df = df.copy()
    df["g1_transport"] = np.where(is_call & is_g1 & (~is_special), df["Amount"], 0.0)
    return df

# replace the whole function
def _calc_after_vat(value, vat_rate: float, vat_mode: str):
    """
    Return 'after-VAT' value. Works with scalar, numpy arrays, or pandas Series.
    For our reporting we always want net = value / (1 + VAT).
    """
    base = 1.0 + float(vat_rate or 0.0)
    # No float() casting of 'value' — keep it vectorized.
    return value / base

def _month_day_counts(year: int, month: int) -> int:
    import calendar
    return calendar.monthrange(year, month)[1]

def _forecast_day_counts(
    df: pd.DataFrame,
    date_source: str,
    month: Optional[str],
    vat_rate: float, vat_mode: str,
    nonempty_only: bool,
    exclude_sundays: bool
) -> Tuple[int, int]:
    """Return (active_days, month_days) for forecast denominator."""
    if month:
        y, m = map(int, month.split("-"))
        month_days = _month_day_counts(y, m)
        mask = df[date_source].dt.to_period("M").astype(str).eq(f"{y:04d}-{m:02d}")
        s = df.loc[mask, date_source].dropna()
    else:
        s = df[date_source].dropna()
        if s.empty:
            return 0, 0
        per = s.dt.to_period("M").mode()
        if per.empty:
            return 0, 0
        yrmo = str(per.iloc[0])
        y, m = map(int, yrmo.split("-"))
        month_days = _month_day_counts(y, m)
        s = s[s.dt.to_period("M").astype(str).eq(yrmo)]

    if s.empty:
        return 0, month_days

    days = s.dt.normalize().unique()
    active_dates = pd.to_datetime(pd.Series(days))
    if nonempty_only:
        # use revenue>0 days (excl CC; treated upstream)
        agg = df.groupby(df[date_source].dt.normalize())["Amount"].sum()
        active_dates = agg[agg > 0].index

    if exclude_sundays:
        active_dates = [d for d in active_dates if pd.Timestamp(d).weekday() != 6]

    return len(active_dates), month_days

def infer_prev_month(month: Optional[str], df_all: pd.DataFrame, date_source: str) -> Optional[str]:
    if month:
        y, m = map(int, month.split("-"))
        return f"{y-1:04d}-{m:02d}"
    # infer from dominant month in df_all, then minus a year
    s = pd.to_datetime(df_all[date_source], errors="coerce")
    if s.notna().sum() == 0:
        return None
    p = s.dt.to_period("M")
    if p.empty:
        return None
    cur = str(p.mode().iloc[0])
    y, m = map(int, cur.split("-"))
    return f"{y-1:04d}-{m:02d}"

def infer_month_from_df(df: Optional[pd.DataFrame], date_source: str) -> Optional[str]:
    if df is None or date_source not in df.columns:
        return None
    s = pd.to_datetime(df[date_source], errors="coerce")
    s = s.dropna()
    if s.empty:
        return None
    return str(s.dt.to_period("M").mode().iloc[0])

# ---------------------------- CORE BUILDERS ----------------------------

def prepare_period_df(
    df_all: pd.DataFrame,
    corr_map: Dict[int, str],
    special_corr: list,
    month: Optional[str],
    date_source: str
) -> pd.DataFrame:
    df = normalize(df_all)
    df = compute_g1_transport(df, special_corr)

    if month:
        mask = df[date_source].dt.to_period("M").astype(str).eq(month)
        df = df.loc[mask].copy()
    return df

def _summary_table(df: pd.DataFrame, call_center_revenue: float, vat_rate: float, vat_mode: str) -> pd.DataFrame:
    total_before_adj = df["Amount"].sum()
    less_g1_transport = df["g1_transport"].sum()
    net_vat_incl = total_before_adj + call_center_revenue - less_g1_transport
    less_vat = net_vat_incl - (net_vat_incl / (1 + vat_rate))
    after_vat = net_vat_incl / (1 + vat_rate)
    # G3 share (% of gross before VAT extraction; excl CC and G1 deduction applied)
    g3_amt = df.loc[df["Warranty"].eq("G3"), "Amount"].sum()
    g3_share = (g3_amt / (total_before_adj if total_before_adj else np.nan)) * 100.0

    out = pd.DataFrame({
        "Metric": [
            "Total Revenue (before adj)",
            "Call center",
            "Less G1 Transport",
            "Net Revenue",
            "Less VAT (12%)",
            "Revenue After VAT",
            "Forecast (After VAT, excl Call Center)",
            "G3 Share (%)",
        ],
        "Value": [
            round(total_before_adj, 2),
            round(call_center_revenue, 2),
            round(less_g1_transport, 2),
            round(net_vat_incl, 2),
            round(less_vat, 2),
            round(after_vat, 2),
            None,  # placeholder filled in process_dataframe
            round(g3_share, 2) if pd.notna(g3_share) else ""
        ]
    })
    return out, after_vat, total_before_adj, less_g1_transport

def _daily_revenue(df: pd.DataFrame, date_source: str, vat_rate: float, vat_mode: str) -> pd.DataFrame:
    g = (
        df.assign(Date=df[date_source].dt.normalize())
          .groupby("Date", dropna=True)
          .agg(Amount=("Amount", "sum"), g1=("g1_transport", "sum"))
          .reset_index()
    )
    g["After VAT (excl CC)"] = _calc_after_vat(g["Amount"] - g["g1"], vat_rate, vat_mode).round(2)
    return g[["Date", "After VAT (excl CC)"]].sort_values("Date")

def build_yoy_monthly(
    df_cur: pd.DataFrame,
    df_prev: pd.DataFrame,
    vat_rate: float, vat_mode: str,
    date_source: str,
    nonempty_only: bool, exclude_sundays: bool,
    month: Optional[str]
) -> pd.DataFrame:
    tot_cur = df_cur["Amount"].sum()
    g1_cur = df_cur["g1_transport"].sum()
    aft_ex_cur_actual = _calc_after_vat(tot_cur - g1_cur, vat_rate, vat_mode)

    act_days, mon_days = _forecast_day_counts(
        df_cur, date_source, month, vat_rate, vat_mode, nonempty_only, exclude_sundays
    )
    factor = (mon_days / act_days) if (act_days and mon_days) else 0.0
    projected_cur = aft_ex_cur_actual * factor

    tot_prev = df_prev["Amount"].sum()
    g1_prev = df_prev["g1_transport"].sum()
    aft_ex_prev_actual = _calc_after_vat(tot_prev - g1_prev, vat_rate, vat_mode)

    delta = projected_cur - aft_ex_prev_actual
    pct = (delta / aft_ex_prev_actual * 100.0) if aft_ex_prev_actual else np.nan

    return pd.DataFrame({
        "Metric": [
            "Projected Revenue (After VAT, excl CC) – Current",
            "Actual Revenue (After VAT, excl CC) – Previous year",
            "Δ vs Previous year",
            "% vs Previous year"
        ],
        "Value": [
            round(projected_cur, 2),
            round(aft_ex_prev_actual, 2),
            round(delta, 2),
            round(pct, 2) if pd.notna(pct) else ""
        ]
    })

def build_yoy_daily(
    df_cur: pd.DataFrame,
    df_prev: pd.DataFrame,
    vat_rate: float, vat_mode: str,
    date_source: str
) -> pd.DataFrame:
    cur = _daily_revenue(df_cur, date_source, vat_rate, vat_mode).rename(
        columns={"After VAT (excl CC)": "Current After VAT (excl CC)"}
    )
    prev = _daily_revenue(df_prev, date_source, vat_rate, vat_mode).rename(
        columns={"After VAT (excl CC)": "PrevYear After VAT (excl CC)"}
    )
    out = pd.merge(cur, prev, on="Date", how="outer").sort_values("Date")
    out["Delta"] = (out["Current After VAT (excl CC)"].fillna(0) -
                    out["PrevYear After VAT (excl CC)"].fillna(0)).round(2)
    return out

def build_yoy_warranty(
    df_cur: pd.DataFrame,
    df_prev: pd.DataFrame,
    vat_rate: float, vat_mode: str,
    date_source: str,
    nonempty_only: bool, exclude_sundays: bool,
    month: Optional[str]
) -> pd.DataFrame:
    def w_after_vat(df):
        g = (df.groupby("Warranty")
               .agg(amt=("Amount", "sum"), g1=("g1_transport", "sum"))
               ).reset_index()
        g["AfterVAT"] = _calc_after_vat(g["amt"] - g["g1"], vat_rate, vat_mode)
        return g[["Warranty", "AfterVAT"]]

    cur = w_after_vat(df_cur).rename(columns={"AfterVAT": "Current After VAT (excl CC)"})
    prev = w_after_vat(df_prev).rename(columns={"AfterVAT": "PrevYear After VAT (excl CC)"})
    m = pd.DataFrame({"Warranty": ["G1", "G2", "G3"]}).merge(cur, on="Warranty", how="left").merge(prev, on="Warranty", how="left")

    act_days, mon_days = _forecast_day_counts(df_cur, date_source, month, vat_rate, vat_mode, nonempty_only, exclude_sundays)
    factor = (mon_days / act_days) if (act_days and mon_days) else 0.0
    m["Current Projected After VAT (excl CC)"] = (m["Current After VAT (excl CC)"] * factor).round(2)

    cur_total_a = float(m["Current After VAT (excl CC)"].sum(skipna=True))
    cur_total_p = float(m["Current Projected After VAT (excl CC)"].sum(skipna=True))
    prev_total  = float(m["PrevYear After VAT (excl CC)"].sum(skipna=True))

    m["Current Share % (Actual)"]    = np.where(cur_total_a != 0, m["Current After VAT (excl CC)"] / cur_total_a * 100.0, np.nan)
    m["Current Share % (Projected)"] = np.where(cur_total_p != 0, m["Current Projected After VAT (excl CC)"] / cur_total_p * 100.0, np.nan)
    m["PrevYear Share %"]            = np.where(prev_total  != 0, m["PrevYear After VAT (excl CC)"] / prev_total  * 100.0, np.nan)

    m["Delta (Proj - Prev)"] = (m["Current Projected After VAT (excl CC)"] - m["PrevYear After VAT (excl CC)"]).round(2)
    m["% vs Prev (Projected)"] = np.where(
        m["PrevYear After VAT (excl CC)"].fillna(0) != 0,
        (m["Delta (Proj - Prev)"] / m["PrevYear After VAT (excl CC)"] * 100.0).round(2),
        np.nan
    )

    total_row = {
        "Warranty": "TOTAL",
        "Current After VAT (excl CC)": round(cur_total_a, 2),
        "PrevYear After VAT (excl CC)": round(prev_total, 2),
        "Current Projected After VAT (excl CC)": round(cur_total_p, 2),
        "Current Share % (Actual)": 100.0 if cur_total_a else np.nan,
        "Current Share % (Projected)": 100.0 if cur_total_p else np.nan,
        "PrevYear Share %": 100.0 if prev_total else np.nan,
        "Delta (Proj - Prev)": round(cur_total_p - prev_total, 2) if (prev_total or cur_total_p) else np.nan,
        "% vs Prev (Projected)": np.nan
    }
    m = pd.concat([m, pd.DataFrame([total_row])], ignore_index=True)
    return m

# ---------------------------- PROCESS + EXPORT ----------------------------

def process_dataframe(
    df_all: pd.DataFrame,
    call_center_revenue: float,
    admin_cost_forecast: float,
    vat_rate: float, vat_mode: str,
    special_corr: list,
    corr_map: Dict[int, str],
    month: Optional[str], date_source: str,
    top_n_names: int,
    nonempty_only: bool, exclude_sundays: bool
) -> Dict[str, pd.DataFrame]:

    df = normalize(df_all)
    df = compute_g1_transport(df, special_corr)

    # duplicate check: by Number of request (not Number of Documents)
    dq_dups = df[df["Number of request"].duplicated(keep=False)].sort_values("Number of request")

    # core summary
    summary, after_vat_actual, total_before_adj, less_g1 = _summary_table(df, call_center_revenue, vat_rate, vat_mode)

    # Forecast (After VAT, excl CC): based on current actual excl CC & G1
    # compute active days & month days using df (current period slice if month provided)
    if month:
        df_for_days = df[df[date_source].dt.to_period("M").astype(str) == month]
    else:
        df_for_days = df.copy()

    act_days, mon_days = _forecast_day_counts(df_for_days, date_source, month, vat_rate, vat_mode, nonempty_only, exclude_sundays)
    factor = (mon_days / act_days) if (act_days and mon_days) else 0.0
    forecast_after_vat_excl_cc = round((_calc_after_vat(total_before_adj - less_g1, vat_rate, vat_mode) * factor), 2)

    # fill forecast into summary
    summary.loc[summary["Metric"].eq("Forecast (After VAT, excl Call Center)"), "Value"] = forecast_after_vat_excl_cc

    daily_rev = _daily_revenue(df, date_source, vat_rate, vat_mode)

    # minimal P&L forecast table (Revenue After VAT minus Admin)
    pnl = pd.DataFrame({
        "Metric": ["Revenue After VAT (actual)", "Admin costs (forecast)", "Operating Profit (forecast)"],
        "Value": [round(after_vat_actual, 2), round(admin_cost_forecast, 2), round(after_vat_actual - admin_cost_forecast, 2)]
    })

    tables = {
        "Summary": summary,
        "Daily_Revenue": daily_rev,
        "P&L_Forecast": pnl,
        "DQ_Checks": dq_dups,
        # Add other pivots (By_Correspondent, By_Warranty, Detailed) as needed
    }
    return tables

def export_excel(tables: Dict[str, pd.DataFrame], out_path: str) -> None:
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        # Summary vertical + light formatting
        if "Summary" in tables:
            df = tables["Summary"].copy()
            df.to_excel(writer, sheet_name="Summary", index=False, startrow=0, startcol=0)
            ws = writer.sheets["Summary"]
            header_fmt = writer.book.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
            for col, name in enumerate(df.columns):
                ws.write(0, col, name, header_fmt)
            ws.set_column(0, 0, 42)
            ws.set_column(1, 1, 20)

        for name, df in tables.items():
            if name == "Summary":
                continue
            # specific sheet names the user requested
            sheet = name
            if name == "P&L_Forecast":
                sheet = "P&L_Forecast"
            if name == "DQ_Checks":
                sheet = "DQ_Checks"
            df.to_excel(writer, sheet_name=sheet, index=False)

    print(f"✅ Saved: {out_path}")

# ---------------------------- STREAMLIT WRAPPER ----------------------------

def run_analysis(
    in_files,
    prev_file: Optional[str] = None,
    call_center: float = 0.0,
    admin_forecast: float = 0.0,
    vat_rate: float = VAT_RATE,
    vat_mode: str = VAT_MODE,
    month: Optional[str] = None,
    forecast_nonempty_only: bool = True,
    no_exclude_sundays: bool = False,
    out_name: Optional[str] = None,
    prev_month_override: Optional[str] = None,
) -> str:
    """Headless runner for Streamlit UI. Returns path to output Excel."""

    # Read current files
    if not in_files:
        raise ValueError("No input files provided.")
    rows = []
    for f in in_files:
        t = read_excel_any(f)
        t["__source_file"] = os.path.basename(f)
        rows.append(t)
    df_all = pd.concat(rows, ignore_index=True)

    special_corr = SPECIAL_CORR_DEFAULT
    corr_map = CORRESPONDENT_MAP_DEFAULT
    date_source = DATE_SOURCE

    tables = process_dataframe(
        df_all,
        parse_money(call_center),
        parse_money(admin_forecast),
        vat_rate, vat_mode,
        special_corr, corr_map,
        month, date_source,
        TOP_NAMES,
        forecast_nonempty_only, not no_exclude_sundays
    )

    # YoY preparation (only if a baseline file provided or same-dataset mode later)
    df_prev_all = None
    if prev_file:
        t = read_excel_any(prev_file)
        t["__source_file"] = os.path.basename(prev_file)
        df_prev_all = t

    # Decide previous month
    prev_month = prev_month_override or infer_month_from_df(df_prev_all, date_source) or infer_prev_month(month, df_all, date_source)

    # Current / Prev period slices
    df_cur_period = prepare_period_df(df_all, corr_map, special_corr, month, date_source)
    df_prev_source = df_prev_all if df_prev_all is not None else df_all
    df_prev_period = prepare_period_df(df_prev_source, corr_map, special_corr, prev_month, date_source)

    if df_cur_period is not None and not df_cur_period.empty and df_prev_period is not None and not df_prev_period.empty:
        yoy_monthly = build_yoy_monthly(
            df_cur_period, df_prev_period, vat_rate, vat_mode, date_source,
            forecast_nonempty_only, not no_exclude_sundays, month
        )
        yoy_daily = build_yoy_daily(
            df_cur_period, df_prev_period, vat_rate, vat_mode, date_source
        )
        yoy_warr = build_yoy_warranty(
            df_cur_period, df_prev_period, vat_rate, vat_mode,
            date_source, forecast_nonempty_only, not no_exclude_sundays, month
        )
        tables["YoY_Monthly"] = yoy_monthly
        tables["YoY_Daily"] = yoy_daily
        tables["YoY_Warranty"] = yoy_warr
    else:
        print("ℹ️ YoY baseline not available or empty slice; skipping YoY sheets.")

    base = out_name or f"monthly_revenue_VAT_{vat_mode}_{month or 'ALL'}"
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M")
    out_path = os.path.join(os.getcwd(), f"{base}__{stamp}.xlsx")
    export_excel(tables, out_path)
    return out_path

