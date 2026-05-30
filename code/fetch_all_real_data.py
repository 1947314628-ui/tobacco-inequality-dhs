#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_all_real_data.py
======================
Pull ALL real data from multiple open APIs and assemble a tidy analysis dataset.

Sources:
  1. DHS Program API — smoking (AH_TOBC_*), smokeless (AH_TOBU_*),
     any tobacco (AH_TOBA_*), with wealth/education/residence/gender breakdowns
  2. WHO GHO API — national smokeless & smoking prevalence by sex
  3. World Bank API — GDP, Gini, urbanisation, health expenditure, schooling

Output: data/analysis_tidy.csv — one row per country×indicator×dimension×subgroup
"""

from __future__ import annotations
import sys, os, time, json
import requests
import pandas as pd
import numpy as np

# ============================ CONFIG ============================
DHS_BASE = "https://api.dhsprogram.com/rest/dhs/data"
GHO_BASE  = "https://ghoapi.azureedge.net/api"
WB_BASE   = "https://api.worldbank.org/v2"

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
os.makedirs(DATA_DIR, exist_ok=True)

REQUEST_PAUSE = 0.4
PER_PAGE = 5000
YEAR_MIN = 2010

# DHS: (indicator_id, tobacco_type, indicator_abbr)
DHS_INDICATORS = [
    ("AH_TOBC_W_ANY", "smoked", "smoked"),
    ("AH_TOBC_M_ANY", "smoked", "smoked"),
    ("AH_TOBU_W_ASM", "smokeless", "smokeless"),
    ("AH_TOBU_M_ASM", "smokeless", "smokeless"),
    ("AH_TOBA_W_ANY", "any_tobacco", "any_tobacco"),
    ("AH_TOBA_M_ANY", "any_tobacco", "any_tobacco"),
]

# GHO indicators
GHO_INDICATORS = {
    "Adult_curr_smokeless": "smokeless",
}
GHO_SMOKED = {
    "M_Est_smk_curr_std": "smoked",
}

# World Bank
WB_INDICATORS = {
    "gdp_pc_usd":      "NY.GDP.PCAP.CD",
    "gini":            "SI.POV.GINI",
    "urban_pct":       "SP.URB.TOTL.IN.ZS",
    "health_exp_pct":  "SH.XPD.CHEX.GD.ZS",
    "edu_years":       "BAR.SCHL.15UP",
}

# Dimension mapping for DHS
DHS_DIM_MAP = {
    "Wealth quintile": "wealth",
    "Education":        "education",
    "Residence":        "residence",
}

# Order mapping
WEALTH_ORDER = {"Lowest": 1, "Second": 2, "Middle": 3, "Fourth": 4, "Highest": 5}
EDU_ORDER    = {"No education": 1, "Primary": 2, "Secondary": 3, "Higher": 4}
RES_ORDER    = {"Rural": 1, "Urban": 2}
# =================================================================


def dhs_pull_all():
    """Pull all DHS indicator data with all breakdowns."""
    all_rows = []
    for ind_id, _, ind_abbr in DHS_INDICATORS:
        print(f"  DHS: {ind_id} ...")
        for page in range(1, 20):
            r = requests.get(DHS_BASE, params={
                "indicatorIds": ind_id,
                "breakdown": "all",
                "f": "json",
                "perpage": str(PER_PAGE),
                "page": str(page),
                "returnFields": "IndicatorId,CountryName,DHS_CountryCode,"
                                "SurveyYear,Value,CharacteristicLabel,"
                                "CharacteristicCategory,IsPreferred,"
                                "DenominatorWeighted",
            }, timeout=120)
            r.raise_for_status()
            js = r.json()
            batch = js.get("Data", [])
            for item in batch:
                item["_tobacco_type"] = ind_abbr
            all_rows.extend(batch)
            total_pages = js.get("TotalPages", 0)
            n_ret = js.get("RecordsReturned", len(batch))
            if page >= total_pages or n_ret == 0:
                break
            time.sleep(REQUEST_PAUSE)
        time.sleep(REQUEST_PAUSE)
    return pd.DataFrame(all_rows)


def gho_pull(indicator_code: str) -> pd.DataFrame:
    """Pull GHO indicator data (all countries, recent years)."""
    url = f"{GHO_BASE}/{indicator_code}"
    # GHO has a low page size limit; use pagination via $skip
    all_rows = []
    skip = 0
    while True:
        params = {"$filter": "SpatialDimType eq 'COUNTRY' and TimeDim ge 2010",
                  "$top": 2000, "$skip": skip}
        r = requests.get(url, params=params, timeout=120)
        if r.status_code == 400:
            # Try without $skip
            params.pop("$skip", None)
            params["$top"] = 500
            r = requests.get(url, params=params, timeout=120)
        r.raise_for_status()
        js = r.json()
        batch = js.get("value", [])
        all_rows.extend(batch)
        if len(batch) < 500:
            break
        skip += 500
        time.sleep(REQUEST_PAUSE)
    return pd.DataFrame(all_rows)


def wb_pull(indicator_code: str, label: str) -> pd.DataFrame:
    """Pull World Bank indicator across all countries."""
    rows = []
    page = 1
    while True:
        r = requests.get(
            f"{WB_BASE}/country/all/indicator/{indicator_code}",
            params={"format": "json", "per_page": 20000, "page": page,
                    "date": f"{YEAR_MIN}:2024"},
            timeout=120)
        js = r.json()
        if not isinstance(js, list) or len(js) < 2 or js[1] is None:
            break
        meta, data = js[0], js[1]
        total_pages = meta.get("pages", 1)
        for d in data:
            if d.get("value") is not None:
                rows.append({
                    "iso3": d.get("countryiso3code", d.get("country", {}).get("id", "")),
                    "year": int(d["date"]),
                    label: d["value"],
                })
        if page >= total_pages:
            break
        page += 1
        time.sleep(REQUEST_PAUSE)
    return pd.DataFrame(rows)


def process_dhs(df: pd.DataFrame) -> pd.DataFrame:
    """Process raw DHS data into tidy format."""
    # Only preferred estimates
    df = df[df.IsPreferred == 1].copy()

    # Map dimension
    df["dimension"] = df.CharacteristicCategory.map(DHS_DIM_MAP)
    df = df[df.dimension.notna()].copy()

    # Map order
    def get_order(row):
        if row.dimension == "wealth":
            return WEALTH_ORDER.get(row.CharacteristicLabel)
        elif row.dimension == "education":
            return EDU_ORDER.get(row.CharacteristicLabel)
        elif row.dimension == "residence":
            return RES_ORDER.get(row.CharacteristicLabel)
        return None

    df["subgroup"] = df.CharacteristicLabel
    df["order"] = df.apply(get_order, axis=1)
    df = df[df.order.notna()].copy()

    # Parse values
    df["estimate"] = pd.to_numeric(df.Value, errors="coerce")
    df["population"] = pd.to_numeric(df.DenominatorWeighted, errors="coerce")

    # Combine male+female by weighted average
    grp_cols = ["CountryName", "DHS_CountryCode", "SurveyYear",
                "dimension", "subgroup", "order", "_tobacco_type"]

    combined = []
    for keys, g in df.groupby(grp_cols):
        g_valid = g.dropna(subset=["estimate", "population"])
        if len(g_valid) == 0:
            continue
        total_pop = g_valid.population.sum()
        w_est = (g_valid.estimate * g_valid.population).sum() / total_pop if total_pop > 0 else g_valid.estimate.mean()
        d = dict(zip(grp_cols, keys))
        d["estimate"] = w_est
        d["population"] = total_pop
        combined.append(d)

    out = pd.DataFrame(combined)
    out = out.rename(columns={
        "CountryName": "country", "DHS_CountryCode": "iso3",
        "SurveyYear": "year", "_tobacco_type": "indicator_abbr",
    })
    out["data_source"] = "DHS"
    return out


def process_gho(df: pd.DataFrame, label: str) -> pd.DataFrame:
    """Process GHO data into tidy format."""
    if df.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "country": df.get("SpatialDim"),
        "iso3":    df.get("SpatialDim"),
        "year":    pd.to_numeric(df.get("TimeDim"), errors="coerce"),
        "indicator_abbr": label,
        "dimension": "sex",
        "subgroup": df.get("Dim1"),
        "order": df.get("Dim1").map({"SEX_MLE": 2, "SEX_FMLE": 2, "SEX_BTSX": 1}).fillna(1),
        "estimate": pd.to_numeric(df.get("NumericValue"), errors="coerce"),
        "population": np.nan,
        "data_source": "GHO",
    })
    return out.dropna(subset=["estimate"])


def main():
    print("=" * 65)
    print("FULL REAL DATA FETCH")
    print("Smokeless Tobacco Inequality Study")
    print("=" * 65)

    # ---- 1. DHS ----
    print("\n[1/3] DHS Program API — disaggregated data ...")
    dhs_raw = dhs_pull_all()
    print(f"  Raw: {len(dhs_raw)} rows")
    dhs_tidy = process_dhs(dhs_raw)
    print(f"  Tidy: {len(dhs_tidy)} rows, {dhs_tidy.country.nunique()} countries")
    for ind in sorted(dhs_tidy.indicator_abbr.unique()):
        sub = dhs_tidy[dhs_tidy.indicator_abbr == ind]
        print(f"    {ind}: {sub.country.nunique()} countries, {len(sub)} rows, "
              f"dims={sorted(sub.dimension.unique())}")

    # ---- 2. GHO ----
    print("\n[2/3] WHO GHO API — national data ...")
    gho_parts = []
    for code, label in GHO_INDICATORS.items():
        print(f"  GHO: {code}")
        raw = gho_pull(code)
        tidy = process_gho(raw, label)
        gho_parts.append(tidy)
        print(f"    -> {len(tidy)} rows")
    for code, label in GHO_SMOKED.items():
        print(f"  GHO: {code}")
        raw = gho_pull(code)
        tidy = process_gho(raw, label)
        gho_parts.append(tidy)
        print(f"    -> {len(tidy)} rows")

    gho_all = pd.concat([p for p in gho_parts if not p.empty], ignore_index=True)

    # ---- 3. World Bank ----
    print("\n[3/3] World Bank API — covariates ...")
    wb_parts = []
    for label, code in WB_INDICATORS.items():
        print(f"  WB: {label} ({code})")
        wb_df = wb_pull(code, label)
        if not wb_df.empty:
            wb_parts.append(wb_df)
            print(f"    -> {len(wb_df)} rows")
        else:
            print(f"    -> empty, skipped")
        time.sleep(REQUEST_PAUSE)

    if wb_parts:
        wb_merged = wb_parts[0]
        for wb2 in wb_parts[1:]:
            wb_merged = wb_merged.merge(wb2, on=["iso3", "year"], how="outer")
    else:
        wb_merged = pd.DataFrame()

    # ---- Merge WB covariates onto DHS & GHO ----
    if not wb_merged.empty:
        for data_part in [dhs_tidy]:
            data_part["_merge_year"] = data_part["year"]
            wb_merged["_merge_year"] = wb_merged["year"]
            merged = pd.merge_asof(
                data_part.sort_values("_merge_year"),
                wb_merged.sort_values("_merge_year"),
                left_on="_merge_year", right_on="_merge_year",
                by="iso3", direction="nearest", tolerance=3)
            merged = merged.drop(columns=["_merge_year"], errors="ignore")
            for col in WB_INDICATORS:
                if col in merged.columns:
                    dhs_tidy[col] = merged[col]

    # ---- Combine & save ----
    all_tidy = pd.concat([dhs_tidy, gho_all], ignore_index=True)

    final_cols = ["country", "iso3", "year", "indicator_abbr", "dimension",
                  "subgroup", "order", "estimate", "population", "data_source"]
    final_cols += [c for c in WB_INDICATORS if c in all_tidy.columns]
    final_cols = [c for c in final_cols if c in all_tidy.columns]
    all_tidy = all_tidy[final_cols].drop_duplicates().reset_index(drop=True)

    out_path = os.path.join(DATA_DIR, "analysis_tidy.csv")
    all_tidy.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"\n{'='*65}")
    print(f"Master dataset: {len(all_tidy)} rows -> {out_path}")
    print(f"Countries: {all_tidy.country.nunique()}")
    print(f"Indicators: {sorted(all_tidy.indicator_abbr.unique())}")
    print(f"Dimensions: {sorted(all_tidy.dimension.unique())}")
    print(f"\nData rows per indicator:")
    for ind in sorted(all_tidy.indicator_abbr.unique()):
        sub = all_tidy[all_tidy.indicator_abbr == ind]
        print(f"  {ind}: {len(sub)} rows, {sub.country.nunique()} countries")
    print(f"\nData rows per dimension:")
    for dim in sorted(all_tidy.dimension.unique()):
        sub = all_tidy[all_tidy.dimension == dim]
        print(f"  {dim}: {len(sub)} rows")
    print("=" * 65)

    return all_tidy


if __name__ == "__main__":
    main()
