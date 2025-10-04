# -*- coding: utf-8 -*-
"""
Created on Fri Oct  3 11:03:28 2025

@author: 6185
"""

# -*- coding: utf-8 -*-
"""
Monthly Revenue Report (All-in-One, Oct 2025)

Features:
- Auto-translate SAP headers (Russian -> English) with detection
- Forecast (After VAT, excl Call Center) with options:
    * --forecast-nonempty-only → base only on days with >0 revenue
    * Sundays excluded by default from both base and target month days
    * --no-exclude-sundays to include Sundays
- G3 Share (%) of service revenue (After VAT, excl CC)
- Admin costs prompt/arg + P&L Forecast (After VAT basis)
- VAT modes: extract (default) or add
- Month filter, date-source switch, single file or folder merge
- External CSVs for SPECIAL_CORR and correspondent name mapping
- Data-quality sheet, extra pivots, reconciliation, formatted Excel

Install once (if needed):
    pip install pandas openpyxl xlsxwriter xlrd pyyaml
"""

import os, sys, glob, calendar, datetime as dt, argparse, json
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd
try:
    import yaml  # optional
except Exception:
    yaml = None

# ------------------------- DEFAULT CONFIG -------------------------

VAT_RATE = 0.12
VAT_MODE = "extract"  # or "add"
DATE_SOURCE = "Data of Document"
TOP_N_NAMES = 50

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

REQUIRED_COLS = [
    "Correspondent", "Number of Documents", "Data of Document", "Data of transaction",
    "Number of request", "Material/SAP Code", "Name", "Qty", "Measurement",
    "Amount", "Currency", "Warranty"
]

# Column header translation (Russian → English)
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

# ------------------------- CLI / CONFIG -------------------------

def load_config(path: Optional[str]) -> dict:
    if not path:
        return {}
    if not os.path.isfile(path):
        return {}
    try:
        if path.lower().endswith((".yml", ".yaml")) and yaml:
            with open(path, "r", encoding="utf-8") as f:
                return yaml.safe_load(f) or {}
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def get_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Monthly Revenue Report (All-in-One)")
    p.add_argument("--in-file", help="Path to a single SAP Excel (.xls/.xlsx)")
    p.add_argument("--in-folder", help="Folder with multiple SAP Excels to consolidate")
    p.add_argument("--config", help="YAML/JSON config path (optional)")

    p.add_argument("--special-corr-csv", help="CSV with SPECIAL_CORR (one id per line)")
    p.add_argument("--corr-map-csv", help="CSV with columns [correspondent_id,name]")

    p.add_argument("--vat-rate", type=float, help="VAT rate, default 0.12")
    p.add_argument("--vat-mode", choices=["extract", "add"], help="VAT mode: extract|add")
    p.add_argument("--date-source", choices=["Data of Document", "Data of transaction"], help="Which date column to use")

    p.add_argument("--month", help="Filter month YYYY-MM (e.g., 2025-09)")
    p.add_argument("--top-n-names", type=int, help="Top N Names for By_Name pivot (default 50)")

    p.add_argument("--call-center", help="Call center revenue (USD, VAT-included). If omitted, prompts.")
    p.add_argument("--admin-forecast", help="Administration costs forecast for the month (USD, After VAT / net). If omitted, prompts.")
    p.add_argument("--out-name", help="Custom output file name (without extension)")

    p.add_argument("--forecast-nonempty-only", action="store_true",
                   help="Forecast uses only active days (>0 revenue) as base.")
    p.add_argument("--no-exclude-sundays", action="store_true",
                   help="Do NOT exclude Sundays (by default Sundays are excluded).")
    p.add_argument("--prev-in-file", help="Baseline (previous period) single Excel to compare against")
    p.add_argument("--prev-in-folder", help="Baseline (previous period) folder of Excels to merge")
    p.add_argument("--prev-month", help="Baseline month YYYY-MM (defaults to same month last year)")

    return p.parse_args()

# ------------------------- IO HELPERS -------------------------
def infer_prev_month(month: Optional[str], df: pd.DataFrame, date_source: str) -> Optional[str]:
    """
    If --prev-month not supplied, return same month last year
    based on --month if present, otherwise infer from modal month in df.
    """
    if month:
        y, m = map(int, month.split("-"))
        return f"{y-1:04d}-{m:02d}"
    # fallback: infer current modal month then minus 1 year
    if date_source in df.columns:
        ds = df[date_source].dropna()
        if not ds.empty:
            per = ds.dt.to_period("M").mode()
            if len(per) > 0:
                y, m = per.iloc[0].year, per.iloc[0].month
                return f"{y-1:04d}-{m:02d}"
    return None

def prepare_period_df(raw: pd.DataFrame,
                      corr_map: Dict[int, str],
                      special_corr: List[int],
                      month: Optional[str],
                      date_source: str) -> pd.DataFrame:
    """Translate → normalize → month filter → G1 logic."""
    df = translate_columns(raw, verbose=False)
    df, _dq = normalize(df, corr_map)
    df = apply_month_filter(df, month, date_source)
    df = compute_g1_transport(df, special_corr)
    return df
def load_list_from_csv(path: Optional[str]) -> Optional[List[int]]:
    if not path:
        return None
    try:
        s = pd.read_csv(path, header=None)[0].tolist()
        return [int(x) for x in s if pd.notna(x)]
    except Exception:
        return None

def load_map_from_csv(path: Optional[str]) -> Optional[Dict[int, str]]:
    if not path:
        return None
    try:
        df = pd.read_csv(path)
        key = "correspondent_id" if "correspondent_id" in df.columns else df.columns[0]
        val = "name" if "name" in df.columns else df.columns[1]
        m = {}
        for _, r in df.iterrows():
            try:
                k = int(r[key])
                v = str(r[val]).strip()
                m[k] = v
            except Exception:
                continue
        return m
    except Exception:
        return None
def prompt_prev_inputs(default_prev_month: Optional[str]) -> tuple[str, Optional[str], Optional[str]]:
    """
    Interactively ask how to supply the baseline (previous period) data.

    Returns (mode, path, prev_month):
      - mode: 'same' | 'file' | 'folder' | 'skip'
      - path: file or folder path (or None)
      - prev_month: 'YYYY-MM' or None (user may leave blank to auto-infer later)
    """
    print("\n=== YoY Baseline (Previous Period) Setup ===")
    print("Choose where the previous-Period data comes from:")
    print("  1) Same dataset as current (I will compare within the same file/folder)")
    print("  2) A separate file (choose a single Excel for previous period)")
    print("  3) A folder of Excels (merge all as previous period)")
    print("  4) Skip YoY comparison for now")
    choice = input("Enter 1/2/3/4 [default: 1]: ").strip() or "1"

    if choice == "4":
        return "skip", None, None

    prev_month_in = input(f"Enter previous period month YYYY-MM "
                          f"(e.g., 2024-10) [default: {default_prev_month or 'auto'}]: ").strip()
    prev_month = prev_month_in if prev_month_in else default_prev_month

    if choice == "1":
        return "same", None, prev_month

    if choice == "2":
        p = input("Path to previous-period file (.xls/.xlsx): ").strip().strip('"').strip("'")
        return "file", p if p else None, prev_month

    if choice == "3":
        p = input("Path to previous-period folder: ").strip().strip('"').strip("'")
        return "folder", p if p else None, prev_month

    # Fallback to same
    return "same", None, prev_month

# ------------------------- CORE HELPERS -------------------------

def parse_money(x: str) -> float:
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
    try:
        return pd.read_excel(path)
    except Exception:
        return pd.read_excel(path, engine="xlrd")

def _norm(s: str) -> str:
    return str(s).strip().lower().replace(" ", "").replace("\u00a0", "").replace("/", "").replace("\\", "")

_RUS_TO_ENG_NORM = { _norm(k): v for k, v in RUS_TO_ENG_COLS.items() }
_ENG_REQUIRED_SET = set(REQUIRED_COLS)

def translate_columns(df: pd.DataFrame, verbose: bool = True) -> pd.DataFrame:
    cols = list(df.columns)
    eng_hits = sum(1 for c in cols if str(c).strip() in _ENG_REQUIRED_SET)
    rus_hits = sum(1 for c in cols if _norm(c) in _RUS_TO_ENG_NORM)

    if eng_hits >= rus_hits and eng_hits > 0:
        df.columns = [str(c).strip() for c in cols]
        if verbose:
            print("ℹ️ Column language: English detected. Translation skipped.")
        return df

    new_cols, translated = [], 0
    for c in cols:
        key = _norm(c)
        if key in _RUS_TO_ENG_NORM:
            new_cols.append(_RUS_TO_ENG_NORM[key]); translated += 1
        else:
            new_cols.append(str(c).strip())
    df.columns = new_cols
    if verbose:
        print(f"✅ Translated {translated} column(s) from Russian to English.") if translated else \
              print("ℹ️ No recognizable Russian headers found. Nothing translated.")
    return df

# ------------------------- TRANSFORMATIONS -------------------------

def normalize(df: pd.DataFrame, corr_map: Dict[int, str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
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

    out["correspondent_name"] = out["Correspondent"].map(corr_map)

    # DQ
    dq_rows = []

    if "Currency" in out.columns:
        mask_non_usd = ~out["Currency"].astype(str).str.upper().eq("USD")
        if mask_non_usd.any():
            ex = out.loc[mask_non_usd, ["Currency", "Amount"]].head(50).copy()
            ex["Issue"] = "Currency != USD"
            dq_rows.append(ex)

    neg = out[out["Amount"] < 0].head(50).copy()
    if not neg.empty:
        neg["Issue"] = "Negative Amount"
        dq_rows.append(neg)

    for c in ["Data of Document", "Data of transaction"]:
        md = out[out[c].isna()].head(50).copy()
        if not md.empty:
            md["Issue"] = f"Missing {c}"
            dq_rows.append(md)

            
    # Duplicates by Number of request (ignore blanks)
    if "Number of request" in out.columns:
        nr = out["Number of request"].astype(str).str.strip()
        valid = nr.ne("") & nr.notna()
        dup_mask = nr[valid].duplicated(keep=False)
        dups = out.loc[valid].loc[dup_mask].sort_values("Number of request").head(100).copy()
    if not dups.empty:
        dups["Issue"] = "Duplicate Number of request"
        dq_rows.append(dups)


    dq = pd.concat(dq_rows, ignore_index=True) if dq_rows else pd.DataFrame([{"Info": "No data quality issues detected"}])

    # Print quick currency warning
    if "Currency" in out.columns:
        mask_non_usd = ~out["Currency"].astype(str).str.upper().eq("USD")
        if mask_non_usd.any():
            rows = out.loc[mask_non_usd, ["Currency", "Amount"]].head(10)
            print("⚠️ Warning: Found rows with Currency != USD. Showing a few examples:")
            print(rows.to_string(index=False))

    return out, dq

def compute_g1_transport(df: pd.DataFrame, special_corr: List[int]) -> pd.DataFrame:
    name_is_call = df["Name"].str.contains("ВЫЗОВ", case=False, na=False)
    is_g1 = df["Warranty"].eq("G1")
    is_special = df["Correspondent"].isin(special_corr)
    df["g1_transport"] = np.where(name_is_call & is_g1 & (~is_special), df["Amount"], 0.0)
    return df

def apply_month_filter(df: pd.DataFrame, month: Optional[str], date_source: str) -> pd.DataFrame:
    if not month:
        return df
    try:
        y, m = map(int, month.split("-"))
        first = dt.datetime(y, m, 1)
        last = dt.datetime(y + (m // 12), (m % 12) + 1, 1)
        mask = (df[date_source] >= first) & (df[date_source] < last)
        return df.loc[mask].copy()
    except Exception:
        print(f"⚠️ Invalid --month '{month}'. Expected YYYY-MM. Filter skipped.")
        return df

# ------------------------- FORECAST DAY LOGIC -------------------------

def _calc_after_vat(amount_vat_incl: float, vat_rate: float, vat_mode: str) -> float:
    return amount_vat_incl / (1 + vat_rate) if vat_mode == "extract" else amount_vat_incl * (1 - vat_rate)

def _forecast_day_counts(df: pd.DataFrame,
                         date_source: str,
                         month: Optional[str],
                         vat_rate: float,
                         vat_mode: str,
                         nonempty_only: bool,
                         exclude_sundays: bool) -> Tuple[int, int]:
    """
    Returns (active_days, month_days) for the forecast scaler.
    active_days:
        - unique dates within target month
        - optionally only days with >0 service revenue (after VAT, excl CC)
        - exclude Sundays if exclude_sundays=True
    month_days:
        - number of calendar days in month
        - exclude Sundays if exclude_sundays=True
    """
    if date_source not in df.columns:
        return 0, 0

    ds = df[date_source].dropna()
    if ds.empty:
        return 0, 0

    if month:
        try:
            y, m = map(int, month.split("-"))
        except Exception:
            per = ds.dt.to_period("M").mode()
            if per.empty:
                return 0, 0
            y, m = per.iloc[0].year, per.iloc[0].month
    else:
        per = ds.dt.to_period("M").mode()
        if per.empty:
            return 0, 0
        y, m = per.iloc[0].year, per.iloc[0].month

    start = dt.datetime(y, m, 1)
    end = dt.datetime(y + (m // 12), (m % 12) + 1, 1)

    dfm = df[(df[date_source] >= start) & (df[date_source] < end)].copy()
    # total month days (exclude Sundays optionally)
    month_days = sum(
        1 for d in pd.date_range(start, end - dt.timedelta(days=1))
        if (not exclude_sundays or d.weekday() != 6)
    )
    if dfm.empty:
        return 0, month_days

    dfm["_date_only"] = dfm[date_source].dt.date
    grp = dfm.groupby("_date_only").agg(amt=("Amount", "sum"), g1=("g1_transport", "sum"))
    # weekday lookup
    weekdays = {d.date(): d.weekday() for d in dfm[date_source]}
    grp["weekday"] = grp.index.map(lambda d: weekdays.get(d, None))

    # per-day service after VAT (excl CC)
    per_day_after_vat = _calc_after_vat(grp["amt"] - grp["g1"], vat_rate, vat_mode)

    mask_active = (per_day_after_vat > 0) if nonempty_only else pd.Series(True, index=per_day_after_vat.index)
    if exclude_sundays:
        mask_active &= (grp["weekday"] != 6)

    active_days = int(mask_active.sum())

    return active_days, month_days

# ------------------------- REPORT BUILDERS -------------------------
def build_yoy_monthly(
    df_cur: pd.DataFrame,
    df_prev: pd.DataFrame,
    vat_rate: float,
    vat_mode: str,
    date_source: str,
    nonempty_only: bool,
    exclude_sundays: bool,
    month: Optional[str]
) -> pd.DataFrame:
    """
    Compare Projected Revenue (After VAT, excl CC) for current month vs previous Period same month.
    """

    # Current-period totals
    tot_cur = df_cur["Amount"].sum()
    g1_cur = df_cur["g1_transport"].sum()
    aft_ex_cur_actual = _calc_after_vat(tot_cur - g1_cur, vat_rate, vat_mode)

    # Factor for projection (month days / active days)
    act_days, mon_days = _forecast_day_counts(
        df_cur, date_source, month, vat_rate, vat_mode,
        nonempty_only, exclude_sundays
    )
    factor = (mon_days / act_days) if (act_days and mon_days) else 0.0
    projected_cur = aft_ex_cur_actual * factor

    # Previous-period totals
    tot_prev = df_prev["Amount"].sum()
    g1_prev = df_prev["g1_transport"].sum()
    aft_ex_prev_actual = _calc_after_vat(tot_prev - g1_prev, vat_rate, vat_mode)

    # Deltas
    delta = projected_cur - aft_ex_prev_actual
    pct = (delta / aft_ex_prev_actual * 100.0) if aft_ex_prev_actual else np.nan

    return pd.DataFrame({
        "Metric": [
            "Projected Revenue (After VAT, excl CC) – Current",
            "Actual Revenue (After VAT, excl CC) – Previous Period",
            "Δ vs Previous Period",
            "% vs Previous Period"
        ],
        "Value": [
            round(projected_cur, 2),
            round(aft_ex_prev_actual, 2),
            round(delta, 2),
            round(pct, 2) if pd.notna(pct) else ""
        ]
    })

def build_yoy_daily(df_cur: pd.DataFrame, df_prev: pd.DataFrame,
                    vat_rate: float, vat_mode: str,
                    date_source: str) -> pd.DataFrame:
    """
    Daily compare: After VAT (excl CC), align by calendar day index (1..N) of month.
    """
    def daily_after_vat(df):
        dcol = date_source
        df = df.copy()
        df["_d"] = df[dcol].dt.date
        g = df.groupby("_d").agg(gross=("Amount","sum"), g1=("g1_transport","sum"))
        out = pd.DataFrame({
            "Date": g.index,
            "After VAT (excl CC)": _calc_after_vat(g["gross"] - g["g1"], vat_rate, vat_mode).values
        }).sort_values("Date")
        out["Day"] = out["Date"].apply(lambda d: d.day)
        return out

    dc = daily_after_vat(df_cur)
    dp = daily_after_vat(df_prev)

    # Build complete day grid using current month length (fallback to prev if empty)
    if not dc.empty:
        y, m = dc["Date"].iloc[0].year, dc["Date"].iloc[0].month
    elif not dp.empty:
        y, m = dp["Date"].iloc[0].year+1, dp["Date"].iloc[0].month  # rough inference
    else:
        return pd.DataFrame([{"Info": "No daily data in either period"}])

    days_in_month = calendar.monthrange(y, m)[1]
    idx = pd.DataFrame({"Day": range(1, days_in_month+1)})

    dc2 = idx.merge(dc[["Day","After VAT (excl CC)"]], on="Day", how="left").rename(
        columns={"After VAT (excl CC)": "Current After VAT (excl CC)"})
    dp2 = idx.merge(dp[["Day","After VAT (excl CC)"]], on="Day", how="left").rename(
        columns={"After VAT (excl CC)": "PrevYear After VAT (excl CC)"})

    out = idx.merge(dc2, on="Day").merge(dp2, on="Day")
    out["Delta"] = (out["Current After VAT (excl CC)"] - out["PrevYear After VAT (excl CC)"]).round(2)
    out["% vs Prev"] = np.where(out["PrevYear After VAT (excl CC)"].fillna(0)!=0,
                                (out["Delta"] / out["PrevYear After VAT (excl CC)"] * 100.0).round(2),
                                np.nan)
    # Optional totals
    out.loc[out.index.max()+1, ["Day","Current After VAT (excl CC)","PrevYear After VAT (excl CC)","Delta","% vs Prev"]] = [
        "Total",
        round(out["Current After VAT (excl CC)"].sum(skipna=True),2),
        round(out["PrevYear After VAT (excl CC)"].sum(skipna=True),2),
        round(out["Delta"].sum(skipna=True),2),
        ""
    ]
    return out

def build_yoy_warranty(
    df_cur: pd.DataFrame,
    df_prev: pd.DataFrame,
    vat_rate: float,
    vat_mode: str,
    date_source: str,
    nonempty_only: bool,
    exclude_sundays: bool,
    month: Optional[str],
) -> pd.DataFrame:
    """
    G1/G2/G3 dynamics with projection:
      - Current Actual After VAT (excl CC)
      - Current Projected After VAT (excl CC) = Actual * (month_days / active_days)
      - PrevYear Actual After VAT (excl CC)
      - Delta & % vs PrevYear (Projected - PrevYear)
      - Shares (%) for Actual, Projected, and PrevYear
    """

    def by_w_after_vat(df: pd.DataFrame) -> pd.DataFrame:
        g = (
            df.groupby("Warranty", dropna=False)
              .agg(amt=("Amount", "sum"), g1=("g1_transport", "sum"))
              .reset_index()
        )
        g["after_vat_excl_cc"] = _calc_after_vat(g["amt"] - g["g1"], vat_rate, vat_mode)
        return g[["Warranty", "after_vat_excl_cc"]]

    # Current actual per warranty
    cur = by_w_after_vat(df_cur).rename(columns={"after_vat_excl_cc": "Current After VAT (excl CC)"})
    # Prev-year actual per warranty
    prev = by_w_after_vat(df_prev).rename(columns={"after_vat_excl_cc": "PrevYear After VAT (excl CC)"})

    # Same warranty axis (G1/G2/G3)
    base_w = pd.DataFrame({"Warranty": ["G1", "G2", "G3"]})
    m = base_w.merge(cur, on="Warranty", how="left").merge(prev, on="Warranty", how="left")

    # --- Projection factor for current month (month_days / active_days)
    act_days, mon_days = _forecast_day_counts(
        df_cur, date_source, month, vat_rate, vat_mode, nonempty_only, exclude_sundays
    )
    factor = (mon_days / act_days) if (act_days and mon_days) else 0.0

    # Projected current per warranty
    m["Current Projected After VAT (excl CC)"] = (m["Current After VAT (excl CC)"] * factor).round(2)

    # Totals for shares
    cur_total_actual   = float(m["Current After VAT (excl CC)"].sum(skipna=True))
    cur_total_project  = float(m["Current Projected After VAT (excl CC)"].sum(skipna=True))
    prev_total_actual  = float(m["PrevYear After VAT (excl CC)"].sum(skipna=True))

    # Shares (%)
    m["Current Share % (Actual)"]   = np.where(cur_total_actual  != 0, m["Current After VAT (excl CC)"]            / cur_total_actual  * 100.0, np.nan)
    m["Current Share % (Projected)"] = np.where(cur_total_project != 0, m["Current Projected After VAT (excl CC)"] / cur_total_project * 100.0, np.nan)
    m["PrevYear Share %"]           = np.where(prev_total_actual != 0, m["PrevYear After VAT (excl CC)"]           / prev_total_actual * 100.0, np.nan)

    # Delta vs PrevYear (Projected - PrevYear)
    m["Delta (Proj - Prev)"] = (m["Current Projected After VAT (excl CC)"] - m["PrevYear After VAT (excl CC)"]).round(2)
    m["% vs Prev (Projected)"] = np.where(
        m["PrevYear After VAT (excl CC)"].fillna(0) != 0,
        (m["Delta (Proj - Prev)"] / m["PrevYear After VAT (excl CC)"] * 100.0).round(2),
        np.nan
    )

    # Round numeric columns
    for c in [
        "Current After VAT (excl CC)",
        "PrevYear After VAT (excl CC)",
        "Current Projected After VAT (excl CC)",
        "Current Share % (Actual)",
        "Current Share % (Projected)",
        "PrevYear Share %"
    ]:
        if c in m.columns:
            m[c] = m[c].round(2)

    # Append TOTAL row (keep numeric types clean; use np.nan for blanks)
    total_row = {
        "Warranty": "TOTAL",
        "Current After VAT (excl CC)": round(cur_total_actual, 2),
        "PrevYear After VAT (excl CC)": round(prev_total_actual, 2),
        "Current Projected After VAT (excl CC)": round(cur_total_project, 2),
        "Current Share % (Actual)": 100.0 if cur_total_actual else np.nan,
        "Current Share % (Projected)": 100.0 if cur_total_project else np.nan,
        "PrevYear Share %": 100.0 if prev_total_actual else np.nan,
        "Delta (Proj - Prev)": round(cur_total_project - prev_total_actual, 2) if prev_total_actual or cur_total_project else np.nan,
        "% vs Prev (Projected)": np.nan,
    }
    m = pd.concat([m, pd.DataFrame([total_row])], ignore_index=True)

    return m


def build_summary(df: pd.DataFrame,
                  call_center_revenue: float,
                  vat_rate: float,
                  vat_mode: str,
                  date_source: str,
                  month: Optional[str],
                  nonempty_only: bool,
                  exclude_sundays: bool) -> pd.DataFrame:
    total_before_adj = df["Amount"].sum()
    less_g1_transport = df["g1_transport"].sum()

    net_vat_incl = total_before_adj + call_center_revenue - less_g1_transport
    net_vat_incl_excl_cc = total_before_adj - less_g1_transport

    revenue_after_vat = _calc_after_vat(net_vat_incl, vat_rate, vat_mode)
    after_vat_excl_cc = _calc_after_vat(net_vat_incl_excl_cc, vat_rate, vat_mode)

    active_days, month_days = _forecast_day_counts(
        df, date_source, month, vat_rate, vat_mode, nonempty_only, exclude_sundays
    )
    forecast_after_vat_excl_cc = after_vat_excl_cc * (month_days / active_days) if (active_days and month_days) else 0.0

    g3_amount = df.loc[df["Warranty"].eq("G3"), "Amount"].sum()
    g3_after_vat = _calc_after_vat(g3_amount, vat_rate, vat_mode)
    g3_share_pct = (g3_after_vat / after_vat_excl_cc * 100.0) if after_vat_excl_cc else 0.0

    less_vat = net_vat_incl - revenue_after_vat if vat_mode == "extract" else net_vat_incl * vat_rate

    data = [
        ("Total Revenue (before adj)", total_before_adj),
        ("Call center", call_center_revenue),
        ("Less G1 Transport", less_g1_transport),
        ("Net Revenue", net_vat_incl),
        (f"Less VAT ({int(vat_rate*100)}%)", less_vat),
        ("Revenue After VAT", revenue_after_vat),
        ("Forecast (After VAT, excl Call Center)", forecast_after_vat_excl_cc),
        ("G3 Share (%)", g3_share_pct),
    ]
    df_out = pd.DataFrame(data, columns=["Metric", "Value"])
    df_out["Value"] = df_out["Value"].round(2)
    return df_out

def build_reconciliation(df: pd.DataFrame,
                         call_center_revenue: float,
                         vat_rate: float,
                         vat_mode: str,
                         date_source: str,
                         month: Optional[str],
                         nonempty_only: bool,
                         exclude_sundays: bool) -> pd.DataFrame:
    total_before_adj = df["Amount"].sum()
    less_g1_transport = df["g1_transport"].sum()

    net_vat_incl = total_before_adj + call_center_revenue - less_g1_transport
    net_vat_incl_excl_cc = total_before_adj - less_g1_transport

    after_vat = _calc_after_vat(net_vat_incl, vat_rate, vat_mode)
    after_vat_excl_cc = _calc_after_vat(net_vat_incl_excl_cc, vat_rate, vat_mode)

    active_days, month_days = _forecast_day_counts(
        df, date_source, month, vat_rate, vat_mode, nonempty_only, exclude_sundays
    )
    forecast_after_vat_excl_cc = after_vat_excl_cc * (month_days / active_days) if (active_days and month_days) else 0.0

    g3_amount = df.loc[df["Warranty"].eq("G3"), "Amount"].sum()
    g3_after_vat = _calc_after_vat(g3_amount, vat_rate, vat_mode)
    g3_share_pct = (g3_after_vat / after_vat_excl_cc * 100.0) if after_vat_excl_cc else 0.0

    less_vat = net_vat_incl - after_vat if vat_mode == "extract" else net_vat_incl * vat_rate

    df_out = pd.DataFrame({
        "Metric": [
            "Sum(Amount)", "Call Center", "Less G1 Transport",
            "Net (VAT-incl)", f"Less VAT @{int(vat_rate*100)}%", "Revenue After VAT",
            "Revenue After VAT (excl CC)", "Forecast (After VAT, excl CC)",
            "G3 After VAT", "G3 Share (%)",
            "Forecast active days", "Forecast month days"
        ],
        "Value": [
            round(total_before_adj, 2), round(call_center_revenue, 2), round(less_g1_transport, 2),
            round(net_vat_incl, 2), round(less_vat, 2), round(after_vat, 2),
            round(after_vat_excl_cc, 2), round(forecast_after_vat_excl_cc, 2),
            round(g3_after_vat, 2), round(g3_share_pct, 2),
            active_days, month_days
        ]
    })
    return df_out

def build_by_correspondent(df: pd.DataFrame, call_center_revenue: float) -> pd.DataFrame:
    g = (
        df.groupby(["Correspondent", "correspondent_name"], dropna=False)
          .agg(gross_amount_usd=("Amount", "sum"),
               g1_transport_usd=("g1_transport", "sum"),
               rows=("Amount", "size"))
          .reset_index()
    )
    g["net_before_vat_usd"] = g["gross_amount_usd"] - g["g1_transport_usd"]
    cc_row = pd.DataFrame([{
        "Correspondent": "CALL_CENTER",
        "correspondent_name": "Call Center Revenue",
        "gross_amount_usd": call_center_revenue,
        "g1_transport_usd": 0.0,
        "rows": 0,
        "net_before_vat_usd": call_center_revenue
    }])
    g = pd.concat([g, cc_row], ignore_index=True)
    for c in ["gross_amount_usd", "g1_transport_usd", "net_before_vat_usd"]:
        g[c] = g[c].round(2)
    return g

def build_by_warranty(df: pd.DataFrame) -> pd.DataFrame:
    g = (
        df.groupby(["Warranty"], dropna=False)
          .agg(amount_usd=("Amount", "sum"),
               g1_transport_usd=("g1_transport", "sum"),
               rows=("Amount", "size"))
          .reset_index()
    )
    g["net_before_vat_usd"] = (g["amount_usd"] - g["g1_transport_usd"]).round(2)
    return g

def build_by_name(df: pd.DataFrame, top_n: int) -> pd.DataFrame:
    g = (
        df.groupby(["Name"], dropna=False)
          .agg(amount_usd=("Amount", "sum"), rows=("Amount", "size"))
          .sort_values("amount_usd", ascending=False)
          .head(top_n)
          .reset_index()
    )
    g["amount_usd"] = g["amount_usd"].round(2)
    return g

def build_pl_forecast(df: pd.DataFrame,
                      call_center_revenue: float,
                      admin_cost_forecast: float,
                      vat_rate: float,
                      vat_mode: str,
                      date_source: str,
                      month: Optional[str],
                      nonempty_only: bool,
                      exclude_sundays: bool) -> pd.DataFrame:
    total_before_adj = df["Amount"].sum()
    less_g1_transport = df["g1_transport"].sum()

    net_vat_incl_excl_cc = total_before_adj - less_g1_transport
    service_after_vat_actual = _calc_after_vat(net_vat_incl_excl_cc, vat_rate, vat_mode)
    cc_after_vat_actual = _calc_after_vat(call_center_revenue, vat_rate, vat_mode)

    active_days, month_days = _forecast_day_counts(
        df, date_source, month, vat_rate, vat_mode, nonempty_only, exclude_sundays
    )
    factor = (month_days / active_days) if (active_days and month_days) else 0.0

    service_after_vat_forecast = service_after_vat_actual * factor
    cc_after_vat_forecast = cc_after_vat_actual * factor

    total_revenue_actual = service_after_vat_actual + cc_after_vat_actual
    total_revenue_forecast = service_after_vat_forecast + cc_after_vat_forecast

    operating_profit_forecast = total_revenue_forecast - admin_cost_forecast
    operating_margin_pct = (operating_profit_forecast / total_revenue_forecast * 100.0) if total_revenue_forecast else 0.0

    return pd.DataFrame({
        "Line": [
            "Revenue (Service) – Actual to date (After VAT, excl CC)",
            "Revenue (Call Center) – Actual to date (After VAT)",
            "Total Revenue – Actual to date (After VAT)",
            "Revenue (Service) – Forecast for month (After VAT, excl CC)",
            "Revenue (Call Center) – Forecast for month (After VAT)",
            "Total Revenue – Forecast for month (After VAT)",
            "Administration Costs – Forecast for month (After VAT / net)",
            "Operating Profit – Forecast for month",
            "Operating Margin % – Forecast for month",
            "Forecast active days (base)",
            "Forecast month days (target)"
        ],
        "Value": [
            round(service_after_vat_actual, 2),
            round(cc_after_vat_actual, 2),
            round(total_revenue_actual, 2),
            round(service_after_vat_forecast, 2),
            round(cc_after_vat_forecast, 2),
            round(total_revenue_forecast, 2),
            round(admin_cost_forecast, 2),
            round(operating_profit_forecast, 2),
            round(operating_margin_pct, 2),
            active_days,
            month_days
        ]
    })

def build_daily_revenue(df: pd.DataFrame,
                        vat_rate: float,
                        vat_mode: str,
                        date_source: str,
                        month: Optional[str]) -> pd.DataFrame:
    """
    Returns a calendar of the month with daily revenue totals.
    Includes days with no rows (0s).
    All figures exclude Call Center (we can’t allocate CC to dates).
    """
    if date_source not in df.columns:
        return pd.DataFrame([{"Info": f"Date column '{date_source}' not found"}])

    ds = df[date_source].dropna()
    if ds.empty:
        return pd.DataFrame([{"Info": "No dates present in dataset"}])

    # Decide target month
    if month:
        try:
            y, m = map(int, month.split("-"))
        except Exception:
            per = ds.dt.to_period("M").mode()
            if per.empty:
                return pd.DataFrame([{"Info": "Cannot infer month"}])
            y, m = per.iloc[0].year, per.iloc[0].month
    else:
        per = ds.dt.to_period("M").mode()
        if per.empty:
            return pd.DataFrame([{"Info": "Cannot infer month"}])
        y, m = per.iloc[0].year, per.iloc[0].month

    start = dt.datetime(y, m, 1)
    end = dt.datetime(y + (m // 12), (m % 12) + 1, 1)

    # Group actuals by date
    dfm = df[(df[date_source] >= start) & (df[date_source] < end)].copy()
    dfm["_date_only"] = dfm[date_source].dt.date
    grp = dfm.groupby("_date_only", dropna=False).agg(
        gross=("Amount", "sum"),
        g1=("g1_transport", "sum"),
    )

    # Build a complete calendar index
    days = [d.date() for d in pd.date_range(start, end - dt.timedelta(days=1))]
    out = []
    for d0 in days:
        gross = float(grp.loc[d0, "gross"]) if d0 in grp.index else 0.0
        g1 = float(grp.loc[d0, "g1"]) if d0 in grp.index else 0.0
        net_vat_incl_excl_cc = gross - g1
        after_vat_excl_cc = _calc_after_vat(net_vat_incl_excl_cc, vat_rate, vat_mode)
        weekday = dt.datetime(d0.year, d0.month, d0.day).strftime("%a")
        is_sunday = dt.datetime(d0.year, d0.month, d0.day).weekday() == 6
        out.append({
            "Date": d0,
            "Weekday": weekday,
            "Is Sunday": bool(is_sunday),
            "Gross Amount (USD)": round(gross, 2),
            "Less G1 Transport (USD)": round(g1, 2),
            "Net (VAT-incl, excl CC)": round(net_vat_incl_excl_cc, 2),
            "After VAT (excl CC)": round(after_vat_excl_cc, 2),
            "Active (>0)": bool(after_vat_excl_cc > 0),
        })

    df_out = pd.DataFrame(out)
    # Optional: cumulative totals
    df_out["Cumulative After VAT (excl CC)"] = df_out["After VAT (excl CC)"].cumsum().round(2)
    return df_out


# ------------------------- EXPORT -------------------------

def export_excel(tables: dict, out_path: str) -> None:
    """
    Export all sheets to Excel.

    Compact vertical tables (Summary, Reconciliation, P&L_Forecast) are written
    cell-by-cell as a 2-column table with styled header and cells ONLY for A1:B{n}.
    Other sheets are written via to_excel, autosized, and lightly number-formatted.
    """
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        wb = writer.book

        # ===== Common formats =====
        header_fmt   = wb.add_format({"bold": True, "bg_color": "#B8CCE4", "border": 1, "align": "left"})
        metric_fmt   = wb.add_format({"bold": True, "bg_color": "#DCE6F1", "border": 1, "align": "left"})
        value_fmt    = wb.add_format({"num_format": "#,##0.00", "border": 1, "align": "right"})
        money_fmt    = wb.add_format({"num_format": "#,##0.00"})
        generic_hdr  = wb.add_format({"bold": True, "bg_color": "#EFEFEF", "border": 1})

        def autosize(ws, df: pd.DataFrame):
            """Autosize columns (no column-wide formats)."""
            for i, col in enumerate(df.columns):
                # look at header + first 200 values for width
                sample = [len(str(col))] + [len(str(x)) for x in df[col].head(200)]
                width = max(10, min(60, max(sample) + 2))
                ws.set_column(i, i, width)

        compact_sheets = {"Summary", "Reconciliation", "P&L_Forecast"}

        for sheet, df in tables.items():
            df_out = df.copy()

            # ---------- Compact vertical tables ----------
            if sheet in compact_sheets and (
                {"Metric", "Value"}.issubset(df_out.columns) or {"Line", "Value"}.issubset(df_out.columns)
            ):
                # Normalize first column name to 'Metric' (P&L uses 'Line')
                if "Line" in df_out.columns and "Metric" not in df_out.columns:
                    df_out = df_out.rename(columns={"Line": "Metric"})

                # Create worksheet manually (do not use to_excel → full control on styling area)
                ws = writer.book.add_worksheet(sheet)
                writer.sheets[sheet] = ws

                # Column widths (no formats attached to columns)
                ws.set_column(0, 0, 35)  # Metric
                ws.set_column(1, 1, 20)  # Value

                # Header
                ws.write(0, 0, "Metric", header_fmt)
                ws.write(0, 1, "Value",  header_fmt)

                # Data rows (A2:B{n+1})
                n = len(df_out)
                for i in range(n):
                    # Metric text
                    ws.write(i + 1, 0, str(df_out.iloc[i, 0]), metric_fmt)
                    # Value with numeric formatting where possible
                    val = df_out.iloc[i, 1]
                    if pd.isna(val):
                        ws.write_blank(i + 1, 1, None, value_fmt)
                    elif isinstance(val, (int, float, np.integer, np.floating)):
                        ws.write_number(i + 1, 1, float(val), value_fmt)
                    else:
                        ws.write(i + 1, 1, str(val), value_fmt)

                # Done with this compact sheet
                continue

            # ---------- Regular sheets ----------
            df_out.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]

            # Style header row (only cells in row 1, not full row)
            for j, col_name in enumerate(df_out.columns):
                ws.write(0, j, col_name, generic_hdr)

            # Autosize
            autosize(ws, df_out)

            # Light numeric formatting for likely money/values
            for j, col_name in enumerate(df_out.columns):
                if df_out[col_name].dtype.kind in "fc" or any(
                    k in col_name.lower() for k in ["amount", "value", "revenue", "vat", "usd"]
                ):
                    ws.set_column(j, j, None, money_fmt)

        print(f"✅ Saved: {out_path}")



# ------------------------- PIPELINE -------------------------

def process_dataframe(df: pd.DataFrame,
                      call_center_revenue: float,
                      admin_cost_forecast: float,
                      vat_rate: float,
                      vat_mode: str,
                      special_corr: List[int],
                      corr_map: Dict[int, str],
                      month: Optional[str],
                      date_source: str,
                      top_n_names: int,
                      nonempty_only: bool,
                      exclude_sundays: bool) -> Dict[str, pd.DataFrame]:

    df = translate_columns(df)
    df, dq = normalize(df, corr_map)
    df = apply_month_filter(df, month, date_source)
    df = compute_g1_transport(df, special_corr)

    summary = build_summary(df, call_center_revenue, vat_rate, vat_mode, date_source, month, nonempty_only, exclude_sundays)
    by_corr = build_by_correspondent(df, call_center_revenue)
    daily = build_daily_revenue(df, vat_rate, vat_mode, date_source, month)
    by_warr = build_by_warranty(df)
    by_name = build_by_name(df, top_n_names)
    reconciliation = build_reconciliation(df, call_center_revenue, vat_rate, vat_mode, date_source, month, nonempty_only, exclude_sundays)
    pl_forecast = build_pl_forecast(df, call_center_revenue, admin_cost_forecast, vat_rate, vat_mode, date_source, month, nonempty_only, exclude_sundays)

    det = df.copy()
    if "g1_transport" in det.columns:
        det["g1_transport"] = det["g1_transport"].round(2)

    return {
        "Summary": summary,
        "Reconciliation": reconciliation,
        "By_Correspondent": by_corr,
        "By_Warranty": by_warr,
        "By_Name_Top": by_name,
        "P&L_Forecast": pl_forecast,
        "Daily_Revenue": daily,
        "Detailed": det,
        "DQ_Checks": dq,
    }

# ------------------------- MAIN -------------------------
# --- Streamlit adapter: run_analysis ---------------------------------
from types import SimpleNamespace
import os, glob, datetime as dt
import pandas as pd
VAT_RATE = 0.12
VAT_MODE = "EXTRACT"
DATE_SOURCE = "Data of Document"
TOP_NAMES = 10  # 👈 add this
def run_analysis(
    in_files,                          # list[str] - current period file paths
    prev_file=None,                    # str|None  - previous year file path
    call_center=0.0,                   # float     - USD (VAT-included)
    admin_forecast=0.0,                # float     - USD (after VAT)
    vat_rate=0.12,                     # 0.12 for 12%
    vat_mode="exclusive",              # or "extract" depending on your code
    month=None,                        # "YYYY-MM"
    forecast_nonempty_only=True,
    no_exclude_sundays=False,
    out_name=None
):
    """Headless runner for Streamlit. Returns output Excel path."""
    # --- mimic args/config the same way your main() expects ---
    args = SimpleNamespace(
        config=None,
        vat_rate=vat_rate,
        vat_mode=vat_mode,
        date_source=None,
        top_n_names=None,
        month=month,
        forecast_nonempty_only=forecast_nonempty_only,
        no_exclude_sundays=no_exclude_sundays,
        special_corr_csv=None,
        corr_map_csv=None,
        in_file=None,
        in_folder=None,
        call_center=call_center,
        admin_forecast=admin_forecast,
        prev_in_file=prev_file,
        prev_in_folder=None,
        prev_month=None,
        out_name=out_name,
    )

    # ---- read CURRENT files (no prompts) ----
    if not in_files:
        raise ValueError("No input files provided.")
    rows = []
    for f in in_files:
        t = read_excel_any(f)
        t["__source_file"] = os.path.basename(f)
        rows.append(t)
    df_all = pd.concat(rows, ignore_index=True)

    # ---- reuse your existing pipeline pieces ----
    cfg = {}
    vat_rate_eff = args.vat_rate if args.vat_rate is not None else VAT_RATE
    vat_mode_eff = args.vat_mode or VAT_MODE
    date_source = args.date_source or DATE_SOURCE
    top_n_names = args.top_n_names or TOP_NAMES if 'TOP_NAMES' in globals() else TOP_N_NAMES
    month_eff = args.month

    nonempty_only = bool(getattr(args, "forecast_nonempty_only", False))
    exclude_sundays = not bool(getattr(args, "no_exclude_sundays", False))

    special_corr = load_list_from_csv(args.special_corr_csv or None) or SPECIAL_CORR_DEFAULT
    corr_map     = load_map_from_csv(args.corr_map_csv or None)     or CORRESPONDENT_MAP_DEFAULT

    tables = process_dataframe(
        df_all,
        float(args.call_center),
        float(args.admin_forecast),
        vat_rate_eff, vat_mode_eff,
        special_corr, corr_map,
        month_eff, date_source,
        top_n_names,
        nonempty_only, exclude_sundays
    )

    # ---- YoY (no interactive prompt) ----
    prev_mode = "skip"
    df_prev_all = None
    if prev_file:
        prev_mode = "file"
        t = read_excel_any(prev_file)
        t["__source_file"] = os.path.basename(prev_file)
        df_prev_all = t

    # infer previous month if not given
    prev_month = infer_prev_month(month_eff, df_all, date_source)

    df_cur_period  = prepare_period_df(df_all,      corr_map, special_corr, month_eff, date_source)
    df_prev_source = df_prev_all if df_prev_all is not None else df_all
    df_prev_period = prepare_period_df(df_prev_source, corr_map, special_corr, prev_month, date_source)

    if (
        prev_mode != "skip"
        and df_prev_period is not None and not df_prev_period.empty
        and df_cur_period  is not None and not df_cur_period.empty
    ):
        yoy_monthly = build_yoy_monthly(
            df_cur_period, df_prev_period,
            vat_rate_eff, vat_mode_eff,
            date_source,
            nonempty_only, exclude_sundays,
            month_eff
        )
        yoy_daily = build_yoy_daily(
            df_cur_period, df_prev_period, vat_rate_eff, vat_mode_eff, date_source
        )
        yoy_warr = build_yoy_warranty(
            df_cur_period, df_prev_period,
            vat_rate_eff, vat_mode_eff,
            date_source, nonempty_only, exclude_sundays, month_eff
        )
        tables["YoY_Monthly"]  = yoy_monthly
        tables["YoY_Daily"]    = yoy_daily
        tables["YoY_Warranty"] = yoy_warr

    # ---- export and return path ----
    base  = out_name or (f"monthly_revenue_VAT_{vat_mode_eff}_{month_eff or 'ALL'}")
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M")
    out_path = os.path.join(os.getcwd(), f"{base}__{stamp}.xlsx")
    export_excel(tables, out_path)
    return out_path

def main():
    # ---------- 0) Args & config ----------
    args = get_args()
    cfg = load_config(args.config)

    vat_rate     = args.vat_rate if args.vat_rate is not None else cfg.get("vat_rate", VAT_RATE)
    vat_mode     = args.vat_mode or cfg.get("vat_mode", VAT_MODE)
    date_source  = args.date_source or cfg.get("date_source", DATE_SOURCE)
    top_n_names  = args.top_n_names or cfg.get("top_n_names", TOP_N_NAMES)
    month        = args.month or cfg.get("month")

    # Forecast switches
    nonempty_only   = bool(getattr(args, "forecast_nonempty_only", False))
    exclude_sundays = not bool(getattr(args, "no_exclude_sundays", False))

    # Maps / lists
    special_corr = load_list_from_csv(args.special_corr_csv or cfg.get("special_corr_csv")) or SPECIAL_CORR_DEFAULT
    corr_map     = load_map_from_csv(args.corr_map_csv or cfg.get("corr_map_csv")) or CORRESPONDENT_MAP_DEFAULT

    # Input location
    in_file   = args.in_file   or cfg.get("in_file")
    in_folder = args.in_folder or cfg.get("in_folder")

    # ---------- 1) Monetary inputs ----------
    # Call Center (VAT-included)
    if args.call_center is not None:
        call_center_revenue = parse_money(args.call_center)
    elif cfg.get("call_center") is not None:
        call_center_revenue = parse_money(str(cfg.get("call_center")))
    else:
        call_center_revenue = parse_money(input("Enter Call Center revenue for the period (USD, VAT-included): "))
    print(f"Call Center revenue set to: {call_center_revenue:,.2f} USD")

    # Admin costs (After VAT / net)
    if getattr(args, "admin_forecast", None) is not None:
        admin_cost_forecast = parse_money(args.admin_forecast)
    elif cfg.get("admin_forecast") is not None:
        admin_cost_forecast = parse_money(str(cfg.get("admin_forecast")))
    else:
        admin_cost_forecast = parse_money(input("Enter Administration costs FORECAST for the month (USD, After VAT / net): "))
    print(f"Administration costs (forecast) set to: {admin_cost_forecast:,.2f} USD\n")

    # ---------- 2) Read current-period files ----------
    files = []
    if in_file:
        files = [in_file]
    elif in_folder:
        files = sorted(glob.glob(os.path.join(in_folder, "*.xls*")))
    else:
        print("Enter path to SAP Excel (.xls/.xlsx) OR folder containing them.")
        p = input("File or Folder: ").strip().strip('"').strip("'")
        if os.path.isdir(p):
            files = sorted(glob.glob(os.path.join(p, "*.xls*")))
        elif os.path.isfile(p):
            files = [p]
        else:
            print(f"❌ Not found: {p}")
            sys.exit(1)

    if not files:
        print("❌ No input files found."); sys.exit(1)

    rows = []
    for f in files:
        try:
            t = read_excel_any(f)
            t["__source_file"] = os.path.basename(f)
            rows.append(t)
        except Exception as e:
            print(f"⚠️ Skipped {f}: {e}")
    if not rows:
        print("❌ No readable files."); sys.exit(1)

    df_all = pd.concat(rows, ignore_index=True)

    # ---------- 3) Build current-period tables ----------
    tables = process_dataframe(
        df_all,
        call_center_revenue,
        admin_cost_forecast,
        vat_rate, vat_mode,
        special_corr, corr_map,
        month, date_source,
        top_n_names,
        nonempty_only, exclude_sundays
    )

    # ---------- 4) YoY baseline prompt / inputs ----------
    default_prev_month = None
    if month:
        try:
            y, m = map(int, month.split("-"))
            default_prev_month = f"{y-1:04d}-{m:02d}"
        except Exception:
            pass

    if args.prev_in_file or args.prev_in_folder or args.prev_month:
        # CLI-driven
        if args.prev_in_file:
            prev_mode, prev_path = "file", args.prev_in_file
        elif args.prev_in_folder:
            prev_mode, prev_path = "folder", args.prev_in_folder
        else:
            prev_mode, prev_path = "same", None
        prev_month = args.prev_month  # may still be None
    else:
        # Interactive
        prev_mode, prev_path, prev_month = prompt_prev_inputs(default_prev_month)

    # ---------- 5) Read previous-period data (if needed) ----------
    df_prev_all = None
    if prev_mode == "file" and prev_path:
        if os.path.isfile(prev_path):
            try:
                t = read_excel_any(prev_path)
                t["__source_file"] = os.path.basename(prev_path)
                df_prev_all = t
            except Exception as e:
                print(f"⚠️ Skipped baseline file {prev_path}: {e}")
        else:
            print(f"⚠️ Baseline file not found: {prev_path}")

    elif prev_mode == "folder" and prev_path:
        if os.path.isdir(prev_path):
            prev_files = sorted(glob.glob(os.path.join(prev_path, "*.xls*")))
            if prev_files:
                prev_rows = []
                for f in prev_files:
                    try:
                        t = read_excel_any(f)
                        t["__source_file"] = os.path.basename(f)
                        prev_rows.append(t)
                    except Exception as e:
                        print(f"⚠️ Skipped baseline file {f}: {e}")
                if prev_rows:
                    df_prev_all = pd.concat(prev_rows, ignore_index=True)
            else:
                print(f"⚠️ No .xls/.xlsx files found in baseline folder: {prev_path}")
        else:
            print(f"⚠️ Baseline folder not found: {prev_path}")

    # If user chose "same", we’ll slice df_all for the previous month.
    if not prev_month:
        prev_month = infer_prev_month(month, df_all, date_source)

    # ---------- 6) Prepare period-sliced DataFrames for YoY ----------
    df_cur_period  = prepare_period_df(df_all,     corr_map, special_corr, month,       date_source)
    df_prev_source = df_prev_all if df_prev_all is not None else df_all
    df_prev_period = prepare_period_df(df_prev_source, corr_map, special_corr, prev_month, date_source)

    # ---------- 7) Build YoY tables (only if feasible and not skipped) ----------
    if (
        prev_mode != "skip"
        and df_prev_period is not None and not df_prev_period.empty
        and df_cur_period  is not None and not df_cur_period.empty
    ):
        yoy_monthly = build_yoy_monthly(
            df_cur_period, df_prev_period,
            vat_rate, vat_mode,
            date_source,
            nonempty_only, exclude_sundays,
            month
        )
        yoy_daily = build_yoy_daily(
            df_cur_period, df_prev_period,
            vat_rate, vat_mode, date_source
        )
        yoy_warr = build_yoy_warranty(
            df_cur_period, df_prev_period,
            vat_rate, vat_mode,
            date_source,
            nonempty_only, exclude_sundays,
            month
        )


        tables["YoY_Monthly"]  = yoy_monthly
        tables["YoY_Daily"]    = yoy_daily
        tables["YoY_Warranty"] = yoy_warr
    else:
        print("ℹ️ YoY: baseline period not available (no data for previous-period month or 'skip' chosen).")

    # ---------- 8) Export ----------
    base  = args.out_name or (f"monthly_revenue_VAT_{vat_mode}_{month or 'ALL'}")
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M")
    out_path = os.path.join(os.getcwd(), f"{base}__{stamp}.xlsx")
    export_excel(tables, out_path)

    # Output
    base = args.out_name or (f"monthly_revenue_VAT_{vat_mode}_{month or 'ALL'}")
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M")
    out_path = os.path.join(os.getcwd(), f"{base}__{stamp}.xlsx")
    export_excel(tables, out_path)

if __name__ == "__main__":
    main()
