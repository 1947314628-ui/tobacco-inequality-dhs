#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analysis_real.py
================
Compute inequality measures (SII, RII, CI, Wagstaff CI) from REAL data.
All data loaded from data/analysis_master.csv (DHS + GHO + World Bank).

Output: output/tables/inequality_results.csv
"""

from __future__ import annotations
import os, sys
import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(__file__))
DATA_PATH = os.path.join(ROOT, "data", "analysis_master.csv")
OUT_DIR  = os.path.join(ROOT, "data")
os.makedirs(OUT_DIR, exist_ok=True)


def _ridit_ranks(weights: np.ndarray) -> np.ndarray:
    w = weights / weights.sum()
    cum = np.cumsum(w)
    return cum - w / 2.0


def inequality_measures(estimate, order, population=None) -> dict:
    est = np.asarray(estimate, float)
    order = np.asarray(order, float)
    n = len(est)
    if population is None or np.all(pd.isna(population)):
        pop = np.ones(n)
    else:
        pop = np.asarray(pd.to_numeric(pd.Series(population), errors="coerce"), float)
        pop = np.where(np.isfinite(pop) & (pop > 0), pop, np.nan)
        if np.all(np.isnan(pop)):
            pop = np.ones(n)
        pop = np.where(np.isnan(pop), np.nanmean(pop), pop)

    idx = np.argsort(order)
    est, pop = est[idx], pop[idx]
    w = pop / pop.sum()
    R = _ridit_ranks(pop)
    mu = np.sum(w * est)
    if mu == 0 or n < 2:
        return dict(mean=mu, SII=np.nan, SII_se=np.nan, RII=np.nan,
                    CI=np.nan, CI_se=np.nan, CI_Wagstaff=np.nan, n_sub=n)

    X = np.column_stack([np.ones(n), R])
    W = np.diag(w)
    XtWX = X.T @ W @ X
    try:
        beta = np.linalg.solve(XtWX, X.T @ W @ est)
    except np.linalg.LinAlgError:
        return dict(mean=mu, SII=np.nan, SII_se=np.nan, RII=np.nan,
                    CI=np.nan, CI_se=np.nan, CI_Wagstaff=np.nan, n_sub=n)

    resid = est - X @ beta
    dof = max(n - 2, 1)
    sigma2 = np.sum(w * resid**2) / dof
    cov_beta = sigma2 * np.linalg.inv(XtWX)
    SII = beta[1]
    SII_se = np.sqrt(max(cov_beta[1, 1], 0))
    y0, y1 = beta[0], beta[0] + beta[1]
    RII = (y1 / y0) if y0 != 0 else np.nan

    var_R = np.sum(w * (R - 0.5)**2)
    CI = (2.0 / mu) * np.sum(w * est * (R - 0.5))
    ystar = 2 * var_R * est / mu
    Xc = np.column_stack([np.ones(n), R])
    try:
        b_c = np.linalg.solve(Xc.T @ W @ Xc, Xc.T @ W @ ystar)
        res_c = ystar - Xc @ b_c
        s2_c = np.sum(w * res_c**2) / dof
        covc = s2_c * np.linalg.inv(Xc.T @ W @ Xc)
        CI_se = np.sqrt(max(covc[1, 1], 0))
    except np.linalg.LinAlgError:
        CI_se = np.nan

    p = mu / 100.0 if mu > 1 else mu
    CI_W = CI / (1 - p) if (0 < p < 1) else np.nan

    return dict(mean=mu, SII=SII, SII_se=SII_se, RII=RII,
                CI=CI, CI_se=CI_se, CI_Wagstaff=CI_W, n_sub=n)


def compute_all(df, group_cols, est_col="estimate",
                order_col="order", pop_col="population"):
    rows = []
    for keys, g in df.groupby(list(group_cols)):
        if g[est_col].notna().sum() < 2:
            continue
        m = inequality_measures(
            g[est_col].values, g[order_col].values,
            g[pop_col].values if pop_col in g else None)
        rows.append({**dict(zip(group_cols, keys)), **m})
    return pd.DataFrame(rows)


def main():
    print("=" * 60)
    print("Inequality Analysis — REAL DATA")
    print("=" * 60)

    df = pd.read_csv(DATA_PATH)
    print(f"\nLoaded: {len(df)} rows, {df.country.nunique()} countries")

    # Filter: DHS primary analysis (wealth/education/residence breakdowns for smoking)
    dhs_smoke = df[(df.data_source == "DHS") &
                   (df.indicator_abbr.isin(["smoked", "any_tobacco"])) &
                   (df.dimension.isin(["wealth", "education", "residence"]))].copy()
    print(f"DHS disaggregated subset: {len(dhs_smoke)} rows, "
          f"{dhs_smoke.country.nunique()} countries")

    # Compute inequality by (country, indicator, dimension, year)
    res = compute_all(dhs_smoke, group_cols=(
        "country", "iso3", "indicator_abbr", "dimension", "year"))

    # Add covariates (per-country most recent)
    cov_cols = ["gdp_pc_usd", "gini", "urban_pct", "health_exp_pct"]
    for col in cov_cols:
        if col not in df.columns:
            continue
        latest = df.dropna(subset=[col]).groupby("iso3").last().reset_index()
        res = res.merge(latest[["iso3", col]].drop_duplicates("iso3"),
                        on="iso3", how="left")

    # Save
    out_path = os.path.join(OUT_DIR, "inequality_results.csv")
    res.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"\nResults saved: {out_path}")
    print(f"  {len(res)} inequality estimates")
    print(f"  {res.country.nunique()} countries")

    # Quick summary
    for ind in sorted(res.indicator_abbr.unique()):
        for dim in sorted(res.dimension.unique()):
            sub = res[(res.indicator_abbr == ind) & (res.dimension == dim)]
            if len(sub) == 0:
                continue
            sig = (abs(sub.CI / sub.CI_se) > 1.96).sum()
            mean_ci = sub.CI.mean()
            mean_sii = sub.SII.mean()
            print(f"\n  [{ind}] [{dim}] n={len(sub)} countries")
            print(f"    Mean CI: {mean_ci:.4f}  (SD={sub.CI.std():.4f})")
            print(f"    Mean SII: {mean_sii:.2f}  (SD={sub.SII.std():.2f})")
            print(f"    CI significant (p<0.05): {sig}/{len(sub)}")
            print(f"    CI < 0 (pro-disadvantaged): {(sub.CI < 0).sum()}/{len(sub)}")

    # Core comparison: any_tobacco vs smoked (paired)
    a = res[(res.indicator_abbr == "any_tobacco") & (res.dimension == "wealth")][
        ["iso3", "country", "CI"]].rename(columns={"CI": "CI_any"})
    b = res[(res.indicator_abbr == "smoked") & (res.dimension == "wealth")][
        ["iso3", "CI"]].rename(columns={"CI": "CI_smoked"})
    paired = a.merge(b, on="iso3")
    if len(paired) > 0:
        paired["delta"] = paired.CI_any - paired.CI_smoked
        print(f"\n  === Core comparison: any_tobacco vs smoked (wealth) ===")
        print(f"  Paired countries: {len(paired)}")
        print(f"  Mean CI_any: {paired.CI_any.mean():.4f}")
        print(f"  Mean CI_smoked: {paired.CI_smoked.mean():.4f}")
        print(f"  Mean delta (any - smoked): {paired.delta.mean():.4f}")
        print(f"  CI_any < CI_smoked: {(paired.CI_any < paired.CI_smoked).sum()}/{len(paired)}")
        if len(paired) > 2:
            from scipy import stats
            t, p = stats.ttest_rel(paired.CI_any, paired.CI_smoked)
            print(f"  Paired t-test: t={t:.3f}, p={p:.4f}")

    print("\nDone.")


if __name__ == "__main__":
    main()
