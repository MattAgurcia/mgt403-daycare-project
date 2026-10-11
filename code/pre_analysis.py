"""
MGT-403 Data Project - Pre-Analysis Project Data Work: Prices and Costs
=======================================================================
One script, end to end:
  Q1  merge NDCP county prices onto the PS5 year x county metro panel, deflate to 2022 $,
      define the analysis dataset              -> data/processed/analysis_dataset.csv
  Q2  2022 summary table (mean / SD / N)       -> output/tables/summary_2022.{csv,md}
  Q3  2022 scatter plots with OLS fit          -> output/figures/scatter_price_vs_*.png
  Q4  2022 regressions, HC1 robust SEs         -> output/tables/regressions_2022.{csv,md}
  write-up                                     -> output/pre_analysis_writeup.md

Run from anywhere:   python3 code/pre_analysis.py
Requires: pandas, numpy, matplotlib, openpyxl.  (statsmodels is optional: if installed it is used
only to cross-check the hand-coded OLS/HC1 results; all reported numbers come from the numpy code
below, so output is identical on any machine.)

Inputs (in the project folder, i.e. the parent of code/):
  NDCP2022.xlsx                   U.S. DOL National Database of Childcare Prices, county level
  ps5_county_panel_2015_2022.csv  PS5 output: year x county metro panel (OEWS 39-9011 wages + ACS),
                                  built by build_childcare_wages.py -> build_acs_county.py ->
                                  build_ps5_panel.py
  county_oews_crosswalk.csv       course-provided crosswalk (used here only to explain non-matches)

CPI-U, U.S. city average, all items (CUUR0000SA0), annual averages, 1982-84=100. Source: BLS,
"Historical Consumer Price Index for All Urban Consumers (CPI-U): U.S. city average, all items",
https://www.bls.gov/cpi/tables/supplemental-files/historical-cpi-u-202212.pdf
"""
import json
import math
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# ------------------------------------------------------------------ paths
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IN = lambda f: os.path.join(ROOT, f)
D_PROC = os.path.join(ROOT, "data", "processed")
D_TAB = os.path.join(ROOT, "output", "tables")
D_FIG = os.path.join(ROOT, "output", "figures")
for d_ in (D_PROC, D_TAB, D_FIG):
    os.makedirs(d_, exist_ok=True)

LOG_LINES = []
def log(*a):
    s = " ".join(str(x) for x in a)
    LOG_LINES.append(s)
    print(s)

# ------------------------------------------------------------------ constants
CPI = {2015: 237.017, 2016: 240.007, 2017: 245.120, 2018: 251.107,
       2019: 255.657, 2020: 258.811, 2021: 270.970, 2022: 292.655}
YEARS = list(range(2015, 2023))
PRICE = "price_preschool"
WAGE = "childcare_wage_a_mean"
RENT = "median_gross_rent"
PCI = "per_capita_income"
PREDICTORS = [WAGE, PCI, "pop_total", "kids_under6", "kids_under6_allpar_lf", RENT,
              "female_college_share", "female_ftyr_share", "under6_share", "under6_careneed_share"]
DOLLARS = [PRICE, WAGE, RENT, PCI]
R = "_real2022"
LOW_OUTLIER = ("29063", 2018)          # DeKalb County, MO 2018 ($37.87/wk): kept, flagged

# ==================================================================================
# Q1. MERGE AND DEFINE THE ANALYSIS DATASET
# ==================================================================================
def build_analysis_dataset():
    res = {}
    log("=" * 78 + "\nQ1 - MERGE NDCP onto PS5 on [year, county_fips]\n" + "=" * 78)

    # ---- NDCP: one sheet; price = MCPRESCHOOL (median weekly, center-based, 36-54 months)
    ndcp = pd.read_excel(IN("NDCP2022.xlsx"), sheet_name="County_LevelNDCP_v8_update2008_",
                         usecols=["STATE_ABBREVIATION", "COUNTY_NAME", "COUNTY_FIPS_CODE",
                                  "STUDYYEAR", "MCPRESCHOOL", "MCPreschool_flag"],
                         dtype={"COUNTY_FIPS_CODE": str}, engine="openpyxl")
    ndcp = ndcp[ndcp["STUDYYEAR"].between(2015, 2022)].copy()
    ndcp["year"] = ndcp["STUDYYEAR"].astype("int64")
    ndcp["county_fips"] = ndcp["COUNTY_FIPS_CODE"].str.strip().str.split(".").str[0].str.zfill(5)
    ndcp = ndcp.rename(columns={"MCPRESCHOOL": PRICE, "MCPreschool_flag": "price_preschool_aggflag",
                                "COUNTY_NAME": "ndcp_county_name", "STATE_ABBREVIATION": "ndcp_state"})
    ndcp = ndcp[["year", "county_fips", "ndcp_state", "ndcp_county_name", PRICE,
                 "price_preschool_aggflag"]]
    assert ndcp["county_fips"].str.fullmatch(r"\d{5}").all()
    assert not ndcp.duplicated(["year", "county_fips"]).any()

    # ---- PS5 panel
    ps5 = pd.read_csv(IN("ps5_county_panel_2015_2022.csv"), dtype={"county_fips": str, "oews_area": str})
    ps5["year"] = ps5["year"].astype("int64")
    ps5["county_fips"] = ps5["county_fips"].str.zfill(5)
    assert ps5["county_fips"].str.fullmatch(r"\d{5}").all()
    assert not ps5.duplicated(["year", "county_fips"]).any()

    res["ndcp_rows"], res["ndcp_counties"] = len(ndcp), ndcp["county_fips"].nunique()
    res["ps5_rows"], res["ps5_counties"] = len(ps5), ps5["county_fips"].nunique()
    log(f"NDCP rows 2015-2022: {res['ndcp_rows']:,} ({res['ndcp_counties']:,} counties)")
    log(f"PS5 rows  2015-2022: {res['ps5_rows']:,} ({res['ps5_counties']:,} counties)")

    m = ps5.merge(ndcp, on=["year", "county_fips"], how="outer", indicator=True, validate="1:1")
    res["matched"] = int((m["_merge"] == "both").sum())
    res["ps5_unmatched"] = int((m["_merge"] == "left_only").sum())
    res["ndcp_unmatched"] = int((m["_merge"] == "right_only").sum())
    log(f"matched {res['matched']:,} | PS5 not in NDCP {res['ps5_unmatched']:,} | "
        f"NDCP not in PS5 {res['ndcp_unmatched']:,}")

    # explain unmatched NDCP rows with the OEWS crosswalk
    xw = pd.read_csv(IN("county_oews_crosswalk.csv"), dtype=str)
    xw["year"] = xw["year"].astype("int64")
    xw["county_fips"] = xw["county_fips"].str.zfill(5)
    xw["oews_area"] = xw["oews_area"].str.zfill(7)
    xw = xw.drop_duplicates(["year", "county_fips"])
    r = m[m["_merge"] == "right_only"][["year", "county_fips", "ndcp_state", "ndcp_county_name"]].merge(
        xw[["year", "county_fips", "oews_area", "area_title"]], on=["year", "county_fips"], how="left")
    r["reason"] = np.where(r["oews_area"].isna(), "not in OEWS crosswalk",
                  np.where(r["oews_area"].str.startswith("00"), "metro in crosswalk but not in PS5",
                           "nonmetro OEWS area (outside metro sample)"))
    res["ndcp_unmatched_reasons"] = r["reason"].value_counts().to_dict()
    res["ndcp_unmatched_examples"] = (r.sort_values(["year", "county_fips"]).groupby("reason").head(4)
                                      [["year", "county_fips", "ndcp_state", "ndcp_county_name",
                                        "oews_area", "area_title", "reason"]])
    log("unmatched NDCP reasons:", res["ndcp_unmatched_reasons"])

    d = m[m["_merge"] != "right_only"].drop(columns="_merge").copy()
    d["acs_ct_2021_carryforward"] = d["acs_ct_2021_carryforward"].fillna(0).astype(int)

    # ---- deflate to 2022 $. Deflating year = panel year for NDCP and OEWS; = ACS 5-year
    # vintage end year for ACS dollars (ACS reports 5-yr estimates in vintage-end-year dollars).
    d["acs_vintage"] = np.where(d["acs_ct_2021_carryforward"] == 1, 2021, d["year"])
    d["cpi_u"] = d["year"].map(CPI)
    for v in [PRICE, WAGE]:
        d[v + "_nominal"] = d[v]
        d[v + R] = d[v] * CPI[2022] / d["year"].map(CPI)
    for v in [RENT, PCI]:
        d[v + "_nominal"] = d[v]
        d[v + R] = d[v] * CPI[2022] / d["acs_vintage"].map(CPI)
    res["cpi_ct_rows"] = int(d["acs_ct_2021_carryforward"].sum())

    # ---- Puerto Rico check: does NDCP's municipio name agree with the ACS name for the same FIPS?
    import re, unicodedata
    def _norm(x):
        x = unicodedata.normalize("NFKD", str(x).split(",")[0]).encode("ascii", "ignore").decode().lower()
        for w in ("county", "municipio", "parish", "city", "borough"):
            x = x.replace(w, "")
        return re.sub(r"[^a-z]", "", x)
    both_ = d[d["ndcp_county_name"].notna()]
    mism = both_[both_["county_name"].map(_norm) != both_["ndcp_county_name"].map(_norm)]
    res["name_mismatch_fips"] = sorted(mism["county_fips"].unique())
    res["name_mismatch_states"] = sorted(mism["county_fips"].str[:2].unique())
    pr = d[d["county_fips"].str.startswith("72")]
    res["pr_price_years"] = sorted(pr.loc[pr[PRICE].notna(), "year"].unique().tolist())
    log(f"NDCP vs ACS name mismatches (FIPS): {res['name_mismatch_fips']}; PR price years: {res['pr_price_years']}")
    assert res["name_mismatch_states"] == ["72"], "name mismatches outside Puerto Rico - inspect"

    # ---- where are prices missing? (metro sample excl. PR)
    nopr = d[~d["county_fips"].str.startswith("72")]
    mp = nopr[nopr[PRICE].isna()]
    res["miss_price_by_state"] = pd.crosstab(mp["state_abbr"], mp["year"])
    res["state_counties"] = nopr.groupby(["state_abbr", "year"]).size().unstack()

    # ---- save the full metro panel BEFORE missing-data filters (excl. Puerto Rico), so the final
    # analysis can fill short interior gaps by straight-line interpolation
    full_cols = (["year", "county_fips", "county_name", "state_abbr", "oews_area", "area_title",
                  "acs_vintage", "cpi_u"] + [v + "_nominal" for v in DOLLARS]
                 + ["pop_total", "kids_under6", "kids_under6_allpar_lf", "female_college_share",
                    "female_ftyr_share", "under6_share", "under6_careneed_share", "acs_ct_2021_carryforward"])
    d[~d["county_fips"].str.startswith("72")][full_cols].sort_values(["county_fips", "year"]).to_csv(
        os.path.join(D_PROC, "panel_all_county_years.csv"), index=False)

    # ---- filters
    steps, cur = [], d
    def step(label, mask):
        nonlocal cur
        before, cur = cur, cur[mask(cur)]
        row = {"step": label, "dropped": len(before) - len(cur), "rows_left": len(cur)}
        row.update({y: int((before["year"] == y).sum() - (cur["year"] == y).sum()) for y in YEARS})
        steps.append(row)
    steps.append({"step": "merged PS5 metro panel", "dropped": 0, "rows_left": len(d),
                  **{y: 0 for y in YEARS}})
    step("drop Puerto Rico (state FIPS 72)", lambda x: ~x["county_fips"].str.startswith("72"))
    step("drop missing preschool price", lambda x: x[PRICE].notna())
    for v in PREDICTORS:
        step(f"drop missing {v}", lambda x, v=v: x[v].notna())
    a = cur.copy()
    assert a[[PRICE] + PREDICTORS].notna().all().all()
    res["steps"] = pd.DataFrame(steps)
    log(res["steps"].to_string(index=False))

    a["flag_low_price_outlier"] = ((a["county_fips"] == LOW_OUTLIER[0]) &
                                   (a["year"] == LOW_OUTLIER[1])).astype(int)
    a["wage_from_divisions"] = a["wage_from_divisions"].fillna(0).astype(int)
    res["per_year"] = a.groupby("year").agg(counties=("county_fips", "nunique"),
                                            states=("state_abbr", "nunique"),
                                            metro_areas=("oews_area", "nunique"))
    res["n_rows"], res["n_counties"] = len(a), a["county_fips"].nunique()
    lo = a[(a["county_fips"] == LOW_OUTLIER[0]) & (a["year"] == LOW_OUTLIER[1])]
    res["low_outlier_price"] = float(lo[PRICE].iloc[0])
    res["low_outlier_name"] = str(lo["county_name"].iloc[0])
    res["n_price_under_50"] = int((a[PRICE] < 50).sum())
    res["flags"] = {"ct": int(a["acs_ct_2021_carryforward"].sum()),
                    "low_outlier": int(a["flag_low_price_outlier"].sum()),
                    "divisions": int(a["wage_from_divisions"].sum())}
    log(res["per_year"].to_string())

    cols = (["year", "county_fips", "county_name", "state_abbr", "oews_area", "area_title",
             "acs_vintage", "cpi_u"]
            + [v + s for v in DOLLARS for s in ("_nominal", R)]
            + ["childcare_wage_h_mean", "childcare_tot_emp", "pop_total", "kids_under6",
               "kids_under6_allpar_lf", "female_college_share", "female_ftyr_share",
               "under6_share", "under6_careneed_share", "price_preschool_aggflag",
               "wage_from_divisions", "acs_ct_2021_carryforward", "flag_low_price_outlier"])
    a = a[cols].sort_values(["year", "county_fips"]).reset_index(drop=True)
    out = os.path.join(D_PROC, "analysis_dataset.csv")
    a.to_csv(out, index=False)
    chk = pd.read_csv(out, dtype={"county_fips": str, "oews_area": str})
    assert len(chk) == len(a) and chk["county_fips"].str.fullmatch(r"\d{5}").all()
    log(f"saved {out} ({len(a):,} rows)")
    return a, res


# ==================================================================================
# Statistics helpers (numpy only, so the script runs without scipy/statsmodels)
# ==================================================================================
def _betacf(a, b, x, itmax=500, eps=3e-16):
    """Continued fraction for the regularized incomplete beta (Numerical Recipes 6.4)."""
    fpmin = 1e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    d = 1.0 / (d if abs(d) > fpmin else fpmin)
    h = d
    for m in range(1, itmax + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d; d = fpmin if abs(d) < fpmin else d
        c = 1.0 + aa / c; c = fpmin if abs(c) < fpmin else c
        d = 1.0 / d; h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d; d = fpmin if abs(d) < fpmin else d
        c = 1.0 + aa / c; c = fpmin if abs(c) < fpmin else c
        d = 1.0 / d; de = d * c; h *= de
        if abs(de - 1.0) < eps:
            break
    return h


def betainc_reg(a, b, x):
    if x <= 0: return 0.0
    if x >= 1: return 1.0
    lbt = math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log(1 - x)
    if x < (a + 1) / (a + b + 2):
        return math.exp(lbt) * _betacf(a, b, x) / a
    return 1.0 - math.exp(lbt) * _betacf(b, a, 1 - x) / b


def t_pvalue_2sided(t, df):
    """Two-sided p-value from Student's t with df degrees of freedom."""
    return betainc_reg(df / 2.0, 0.5, df / (df + t * t))


def chi2_1_sf(x):
    """Survival function of chi-square with 1 df."""
    return math.erfc(math.sqrt(x / 2.0))


def ols(y, X, names):
    """OLS with classical and HC1 standard errors. X must include the constant column."""
    y, X = np.asarray(y, float), np.asarray(X, float)
    n, k = X.shape
    XtX_inv = np.linalg.inv(X.T @ X)
    b = XtX_inv @ X.T @ y
    e = y - X @ b
    sse, sst = e @ e, ((y - y.mean()) ** 2).sum()
    r2 = 1 - sse / sst
    df = n - k
    V_cl = XtX_inv * (sse / df)
    meat = (X * (e ** 2)[:, None]).T @ X
    V_hc1 = XtX_inv @ meat @ XtX_inv * n / df
    se = np.sqrt(np.diag(V_hc1))
    t = b / se
    p = np.array([t_pvalue_2sided(ti, df) for ti in t])
    return {"names": names, "b": b, "se": se, "t": t, "p": p, "se_classical": np.sqrt(np.diag(V_cl)),
            "n": n, "k": k, "df": df, "r2": r2, "adj_r2": 1 - (1 - r2) * (n - 1) / df,
            "resid": e, "fitted": X @ b}


def breusch_pagan(res, X):
    """Koenker/studentized Breusch-Pagan LM = n * R^2 from regressing e^2 on X."""
    e2 = res["resid"] ** 2
    aux = ols(e2, X, ["c"] + ["x"] * (X.shape[1] - 1))
    lm = res["n"] * aux["r2"]
    return lm


def fmt_p(p):
    return "<0.001" if p < 0.001 else f"{p:.3f}"


# ==================================================================================
# Q2. SUMMARY TABLE (2022)
# ==================================================================================
LABELS = [
    (PRICE + R, "Median weekly price, center-based preschool", "weekly $, 2022 $"),
    (WAGE + R, "Childcare-worker mean annual wage (OEWS 39-9011)", "annual $, 2022 $"),
    (RENT + R, "Median gross rent (rent + utilities)", "monthly $, 2022 $"),
    (PCI + R, "Per-capita income", "annual $, 2022 $"),
    ("pop_total", "Total population", "persons"),
    ("kids_under6", "Own children under 6", "children"),
    ("kids_under6_allpar_lf", "Children under 6, all resident parents in labor force", "children"),
    ("female_college_share", "Share of women 25+ with bachelor's degree or higher", "proportion (0-1)"),
    ("female_ftyr_share", "Share of women 16-64 working full-time, year-round", "proportion (0-1)"),
    ("under6_share", "Children under 6 / total population", "proportion (0-1)"),
    ("under6_careneed_share", "Children under 6 w/ all parents in LF / total population", "proportion (0-1)"),
]


def summary_table(a22):
    rows = []
    for v, lab, unit in LABELS:
        s = a22[v]
        rows.append({"variable": v, "label": lab, "units": unit, "mean": s.mean(), "sd": s.std(ddof=1),
                     "n": int(s.notna().sum()), "min": s.min(), "max": s.max()})
    t = pd.DataFrame(rows)
    t.to_csv(os.path.join(D_TAB, "summary_2022.csv"), index=False)

    def f(v, x):
        if "share" in v:
            return f"{x:.3f}"
        if v in (PRICE + R,):
            return f"{x:,.2f}"
        return f"{x:,.0f}"
    md = ["| Variable | Units | Mean | SD | N | Min | Max |", "|---|---|---:|---:|---:|---:|---:|"]
    for _, r in t.iterrows():
        md.append(f"| {r['label']} | {r['units']} | {f(r['variable'], r['mean'])} | "
                  f"{f(r['variable'], r['sd'])} | {r['n']:,} | {f(r['variable'], r['min'])} | "
                  f"{f(r['variable'], r['max'])} |")
    md = "\n".join(md)
    with open(os.path.join(D_TAB, "summary_2022.md"), "w") as fh:
        fh.write(md + "\n")
    return t, md


# ==================================================================================
# Q3 + Q4. SCATTERS AND REGRESSIONS (2022)
# ==================================================================================
def label_county(row):
    name = str(row["county_name"]).split(",")[0]
    return f"{name}, {row['state_abbr']}"


def scatter(a22, xvar, xlab, fname, title, fit, box_x=0.02):
    x, y = a22[xvar].to_numpy(), a22["y"].to_numpy()
    fig, ax = plt.subplots(figsize=(8, 5.6), dpi=150)
    ax.scatter(x, y, s=14, alpha=0.45, color="#2f6db5", edgecolor="none", label="County (2022)")
    xs = np.linspace(x.min(), x.max(), 100)
    ax.plot(xs, fit["b"][0] + fit["b"][1] * xs, color="#c0392b", lw=2, label="OLS fit")
    tmp = a22.assign(res=fit["resid"])
    top = tmp.loc[tmp["res"].abs().sort_values(ascending=False).index].head(5)
    xmax = x.max()
    for i, (_, r) in enumerate(top.sort_values("y").iterrows()):
        right = r[xvar] > x.min() + 0.8 * (xmax - x.min())
        ax.annotate(label_county(r), (r[xvar], r["y"]), fontsize=7,
                    xytext=(-5 if right else 5, 4 if i % 2 == 0 else -10),
                    ha="right" if right else "left", textcoords="offset points", color="#333333")
    txt = (f"slope = {fit['b'][1]:.2f}  (HC1 SE {fit['se'][1]:.2f})\n"
           f"R² = {fit['r2']:.3f}    N = {fit['n']:,}")
    ax.text(box_x, 0.97, txt, transform=ax.transAxes, va="top", fontsize=10,
            bbox=dict(boxstyle="round", fc="white", ec="#999999"))
    ax.set_xlabel(xlab); ax.set_ylabel("Median weekly preschool price, center-based (2022 \\$)")
    ax.set_title(title, fontsize=12); ax.grid(alpha=0.25); ax.legend(loc="lower right", fontsize=9)
    fig.tight_layout()
    path = os.path.join(D_FIG, fname)
    fig.savefig(path); plt.close(fig)
    return path, top


def regressions(a22):
    a22 = a22.copy()
    a22["y"] = a22[PRICE + R]
    a22["wage_k"] = a22[WAGE + R] / 1000.0          # $1,000s per year
    a22["rent_h"] = a22[RENT + R] / 100.0           # $100s per month
    a22["col_pp"] = a22["female_college_share"] * 100.0  # women 25+ with BA+, in percentage points
    n = len(a22)
    one = np.ones(n)
    Xa = np.column_stack([one, a22["wage_k"]]); Xb = np.column_stack([one, a22["rent_h"]])
    Xc = np.column_stack([one, a22["wage_k"], a22["rent_h"]])
    ma = ols(a22["y"], Xa, ["Intercept", "wage_k"])
    mb = ols(a22["y"], Xb, ["Intercept", "rent_h"])
    mc = ols(a22["y"], Xc, ["Intercept", "wage_k", "rent_h"])
    Xd = np.column_stack([one, a22["col_pp"]])
    Xe = np.column_stack([one, a22["wage_k"], a22["rent_h"], a22["col_pp"]])
    md_ = ols(a22["y"], Xd, ["Intercept", "col_pp"])
    me = ols(a22["y"], Xe, ["Intercept", "wage_k", "rent_h", "col_pp"])

    # cross-check slope/classical SE against an independent closed form (univariate)
    for m_, xv in ((ma, "wage_k"), (mb, "rent_h"), (md_, "col_pp")):
        xx, yy = a22[xv].to_numpy(), a22["y"].to_numpy()
        sxx = ((xx - xx.mean()) ** 2).sum()
        slope = ((xx - xx.mean()) * (yy - yy.mean())).sum() / sxx
        se_cl = math.sqrt((m_["resid"] @ m_["resid"]) / (len(xx) - 2) / sxx)
        assert abs(slope - m_["b"][1]) < 1e-9 and abs(se_cl - m_["se_classical"][1]) < 1e-9
    try:   # optional cross-check with statsmodels if available
        import statsmodels.api as sm
        for m_, X in ((ma, Xa), (mb, Xb), (mc, Xc), (md_, Xd), (me, Xe)):
            f = sm.OLS(a22["y"].to_numpy(), X).fit(cov_type="HC1")
            assert np.allclose(f.params, m_["b"]) and np.allclose(f.bse, m_["se"])
        log("statsmodels cross-check: coefficients and HC1 SEs match")
    except ImportError:
        log("statsmodels not installed: skipped optional cross-check (numpy results used)")

    # correlation and OVB decomposition
    corr = float(np.corrcoef(a22["wage_k"], a22["rent_h"])[0, 1])
    aux_rw = ols(a22["rent_h"], Xa, ["Intercept", "wage_k"])     # rent on wage
    aux_wr = ols(a22["wage_k"], Xb, ["Intercept", "rent_h"])     # wage on rent
    bw_c, br_c = mc["b"][1], mc["b"][2]
    ovb = {
        "corr": corr,
        "delta_rent_on_wage": aux_rw["b"][1], "delta_wage_on_rent": aux_wr["b"][1],
        "wage_a": ma["b"][1], "wage_c": bw_c, "bias_wage": br_c * aux_rw["b"][1],
        "rent_b": mb["b"][1], "rent_c": br_c, "bias_rent": bw_c * aux_wr["b"][1],
    }
    # in-sample OLS identity: short coef = long coef + (other long coef) x (aux slope)
    assert abs(ovb["wage_a"] - (bw_c + ovb["bias_wage"])) < 1e-8
    assert abs(ovb["rent_b"] - (br_c + ovb["bias_rent"])) < 1e-8
    ovb["pred_sign_wage"] = np.sign(br_c) * np.sign(aux_rw["b"][1])
    ovb["pred_sign_rent"] = np.sign(bw_c) * np.sign(aux_wr["b"][1])
    ovb["wage_change_matches"] = bool(np.sign(ovb["wage_a"] - bw_c) == ovb["pred_sign_wage"])
    ovb["rent_change_matches"] = bool(np.sign(ovb["rent_b"] - br_c) == ovb["pred_sign_rent"])
    log("OVB:", {k: (round(v, 4) if isinstance(v, float) else v) for k, v in ovb.items()})

    # female_college_share: (1) model (d) vs (e): short coef = long coef + sum_j(long coef_j x slope of
    # omitted j on college share). (2) model (c) vs (e): c coef = e coef + (college coef in e) x
    # (coef of wage/rent in regression of college share on wage + rent). Both exact OLS identities.
    aux_wk = ols(a22["wage_k"], Xd, ["Intercept", "col_pp"])
    aux_rk = ols(a22["rent_h"], Xd, ["Intercept", "col_pp"])
    aux_c = ols(a22["col_pp"], Xc, ["Intercept", "wage_k", "rent_h"])
    kids = {
        "corr_wage": float(np.corrcoef(a22["col_pp"], a22["wage_k"])[0, 1]),
        "corr_rent": float(np.corrcoef(a22["col_pp"], a22["rent_h"])[0, 1]),
        "kids_d": md_["b"][1], "kids_e": me["b"][3],
        "delta_wage_on_kids": aux_wk["b"][1], "delta_rent_on_kids": aux_rk["b"][1],
        "via_wage": me["b"][1] * aux_wk["b"][1], "via_rent": me["b"][2] * aux_rk["b"][1],
        "sd": a22["col_pp"].std(), "mean": a22["col_pp"].mean(), "p10": a22["col_pp"].quantile(0.10),
        "p90": a22["col_pp"].quantile(0.90), "min": a22["col_pp"].min(), "max": a22["col_pp"].max(),
        "g_wage": aux_c["b"][1], "g_rent": aux_c["b"][2],
        "bias_c_wage": me["b"][3] * aux_c["b"][1], "bias_c_rent": me["b"][3] * aux_c["b"][2],
    }
    assert abs(kids["kids_d"] - (kids["kids_e"] + kids["via_wage"] + kids["via_rent"])) < 1e-8
    assert abs(mc["b"][1] - (me["b"][1] + kids["bias_c_wage"])) < 1e-8
    assert abs(mc["b"][2] - (me["b"][2] + kids["bias_c_rent"])) < 1e-8
    log("female_college_share:", {k: round(v, 4) for k, v in kids.items()})

    # heteroskedasticity (univariate fits)
    bp = {"wage": breusch_pagan(ma, Xa), "rent": breusch_pagan(mb, Xb)}
    bp = {k: {"lm": v, "p": chi2_1_sf(v)} for k, v in bp.items()}

    # ---- regression table
    rows, md = [], []
    terms = [("Intercept", "Intercept"), ("wage_k", "Childcare wage ($1,000s/yr, 2022 $)"),
             ("rent_h", "Median gross rent ($100s/mo, 2022 $)"),
             ("col_pp", "Female college share (pct. points)")]
    for lab, m_ in (("(a)", ma), ("(b)", mb), ("(c)", mc), ("(d)", md_), ("(e)", me)):
        for i, nm in enumerate(m_["names"]):
            rows.append({"model": lab, "term": nm, "coef": m_["b"][i], "se_hc1": m_["se"][i],
                         "t": m_["t"][i], "p": m_["p"][i], "se_classical": m_["se_classical"][i]})
        rows.append({"model": lab, "term": "N", "coef": m_["n"]})
        rows.append({"model": lab, "term": "R2", "coef": m_["r2"]})
        rows.append({"model": lab, "term": "adj_R2", "coef": m_["adj_r2"]})
    pd.DataFrame(rows).to_csv(os.path.join(D_TAB, "regressions_2022.csv"), index=False)

    models = [ma, mb, mc, md_, me]
    md.append("| | (a) Wage only | (b) Rent only | (c) Wage + rent | (d) College share only | (e) Wage + rent + college share |")
    md.append("|---|---:|---:|---:|---:|---:|")
    for key, lab in terms:
        cells = {"b": [], "se": [], "t": [], "p": []}
        for m_ in models:
            if key in m_["names"]:
                i = m_["names"].index(key)
                cells["b"].append(f"**{m_['b'][i]:.3f}**"); cells["se"].append(f"({m_['se'][i]:.3f})")
                cells["t"].append(f"t = {m_['t'][i]:.2f}"); cells["p"].append(f"p {'< 0.001' if m_['p'][i] < 0.001 else '= ' + format(m_['p'][i], '.3f')}")
            else:
                for c in cells: cells[c].append("")
        md.append(f"| {lab} | " + " | ".join(cells["b"]) + " |")
        md.append("| &nbsp;&nbsp;robust SE (HC1) | " + " | ".join(cells["se"]) + " |")
        md.append("| &nbsp;&nbsp;t-stat | " + " | ".join(cells["t"]) + " |")
        md.append("| &nbsp;&nbsp;p-value | " + " | ".join(cells["p"]) + " |")
    md.append("| N | " + " | ".join(f"{m_['n']:,}" for m_ in models) + " |")
    md.append("| R² | " + " | ".join(f"{m_['r2']:.3f}" for m_ in models) + " |")
    md.append("| Adjusted R² | " + " | ".join(f"{m_['adj_r2']:.3f}" for m_ in models) + " |")
    md = "\n".join(md)
    with open(os.path.join(D_TAB, "regressions_2022.md"), "w") as fh:
        fh.write(md + "\n")

    # ---- scatters
    pa, top_w = scatter(a22.assign(x=a22["wage_k"]), "wage_k",
                        "Childcare-worker mean annual wage (\\$1,000s, 2022 \\$)",
                        "scatter_price_vs_wage_2022.png",
                        "Preschool price vs. childcare-worker wage, metro counties, 2022", ma)
    pb, top_r = scatter(a22, "rent_h", "Median gross rent (\\$100s per month, 2022 \\$)",
                        "scatter_price_vs_rent_2022.png",
                        "Preschool price vs. median gross rent, metro counties, 2022", mb)
    pk, top_k = scatter(a22, "col_pp",
                        "Women 25+ with a bachelor's degree or higher (% of women 25+)",
                        "scatter_price_vs_female_college_share_2022.png",
                        "Preschool price vs. female college share, metro counties, 2022", md_)
    return {"top_k": top_k, "a": ma, "b": mb, "c": mc, "d": md_, "e": me, "kids": kids, "ovb": ovb, "bp": bp, "md": md,
            "top_w": top_w, "top_r": top_r, "a22": a22}


# ==================================================================================
# PREDICTOR SCREEN (2022): price on each summary-table variable, one at a time,
# and each non-cost variable added to wage + rent
# ==================================================================================
SCREEN = [  # (column, label, scale divisor, unit of the coefficient)
    (WAGE + R, "Childcare-worker mean annual wage", 1000, "per $1,000/yr"),
    (RENT + R, "Median gross rent", 100, "per $100/mo"),
    (PCI + R, "Per-capita income", 1000, "per $1,000/yr"),
    ("female_college_share", "Share of women 25+ with bachelor's+", 0.01, "per pct. point"),
    ("kids_under6_allpar_lf", "Children <6, all parents in labor force", 1000, "per 1,000 children"),
    ("pop_total", "Total population", 100000, "per 100,000 people"),
    ("kids_under6", "Own children under 6", 1000, "per 1,000 children"),
    ("female_ftyr_share", "Share of women 16-64 full-time year-round", 0.01, "per pct. point"),
    ("under6_careneed_share", "Children <6 w/ all parents in LF / population", 0.01, "per pct. point"),
    ("under6_share", "Children under 6 / population", 0.01, "per pct. point"),
]


def predictor_screen(a22):
    y = a22[PRICE + R].to_numpy()
    one = np.ones(len(a22))
    w, r = a22[WAGE + R] / 1000.0, a22[RENT + R] / 100.0
    base = ols(y, np.column_stack([one, w, r]), ["c", "w", "r"])
    rows = []
    for col, lab, div, unit in SCREEN:
        x = a22[col] / div
        u = ols(y, np.column_stack([one, x]), ["c", "x"])
        row = {"variable": col, "label": lab, "coef_unit": unit, "uni_coef": u["b"][1], "uni_se_hc1": u["se"][1],
               "uni_p": u["p"][1], "uni_r2": u["r2"]}
        if col not in (WAGE + R, RENT + R):
            m = ols(y, np.column_stack([one, w, r, x]), ["c", "w", "r", "x"])
            row.update({"wr_coef": m["b"][3], "wr_se_hc1": m["se"][3], "wr_p": m["p"][3],
                        "wr_r2": m["r2"], "wr_delta_r2": m["r2"] - base["r2"]})
        rows.append(row)
    t = pd.DataFrame(rows).sort_values("uni_r2", ascending=False).reset_index(drop=True)
    t.insert(0, "rank_uni_r2", range(1, len(t) + 1))
    t.to_csv(os.path.join(D_TAB, "predictor_screen_2022.csv"), index=False)
    pv = lambda p: "< 0.001" if p < 0.001 else f"{p:.3f}"
    md = ["| Rank | Variable | Coef. unit | Alone: coef (HC1 SE) | Alone: p | **Alone: R²** | With wage + rent: coef (HC1 SE) | p | R² | **ΔR² vs. wage + rent** |",
          "|---:|---|---|---:|---:|---:|---:|---:|---:|---:|"]
    for _, q in t.iterrows():
        tail = ("| — | — | — | — |" if pd.isna(q.get("wr_r2")) else
                f"| {q['wr_coef']:.3f} ({q['wr_se_hc1']:.3f}) | {pv(q['wr_p'])} | {q['wr_r2']:.3f} | **{q['wr_delta_r2']:+.3f}** |")
        md.append(f"| {q['rank_uni_r2']} | {q['label']} | {q['coef_unit']} | {q['uni_coef']:.3f} ({q['uni_se_hc1']:.3f}) | "
                  f"{pv(q['uni_p'])} | **{q['uni_r2']:.3f}** " + tail)
    md = "\n".join(md)
    with open(os.path.join(D_TAB, "predictor_screen_2022.md"), "w") as fh:
        fh.write(md + "\n")
    log(t.to_string(index=False))
    return {"table": t, "md": md, "base_r2": base["r2"]}


# ==================================================================================
def main():
    a, q1 = build_analysis_dataset()
    a22 = a[a["year"] == 2022].copy()
    log("\n" + "=" * 78 + "\nQ2 - SUMMARY TABLE 2022\n" + "=" * 78)
    summ, summ_md = summary_table(a22)
    log(summ.to_string(index=False))
    log("\n" + "=" * 78 + "\nQ3/Q4 - SCATTERS AND REGRESSIONS 2022\n" + "=" * 78)
    reg = regressions(a22)
    reg["screen"] = predictor_screen(a22)
    log(reg["md"])
    for k in ("top_w", "top_r"):
        log(k, reg[k][["county_fips", "county_name", "area_title", "y", "wage_k", "rent_h", "res"]].to_string(index=False))
    log("BP:", reg["bp"])
    # extra descriptives used in the write-up
    a22r = reg["a22"]
    extra = {
        "hc1_vs_classical": {m: (reg[m]["se"][1] / reg[m]["se_classical"][1]) for m in ("a", "b")},
        "sd_wage_k": a22r["wage_k"].std(), "sd_rent_h": a22r["rent_h"].std(), "sd_y": a22r["y"].std(),
    }
    log("extra:", extra)
    if "--json" in sys.argv:
        print(json.dumps({"extra": extra}, default=float))
    write_writeup(q1, summ, summ_md, reg, extra)
    with open(os.path.join(ROOT, "output", "pre_analysis_log.txt"), "w") as fh:
        fh.write("\n".join(LOG_LINES) + "\n")


def write_writeup(q1, summ, summ_md, reg, extra):
    """Every number below is pulled from objects computed above. The asserts guard the
    qualitative statements in the text (signs, directions), so if the inputs change the
    script fails rather than writing an interpretation the numbers no longer support."""
    ma, mb, mc, ovb, bp, a22 = reg["a"], reg["b"], reg["c"], reg["ovb"], reg["bp"], reg["a22"]
    md_, me, kd = reg["d"], reg["e"], reg["kids"]
    pv = lambda p: "< 0.001" if p < 0.001 else "= " + format(p, ".3f")
    sc = reg["screen"]; tb = sc["table"].set_index("variable")
    nc = tb.dropna(subset=["wr_r2"]).sort_values("wr_delta_r2", ascending=False)
    top_u, top_i = tb.index[:4].tolist(), nc.index[:2].tolist()
    assert top_i == [PCI + R, "female_college_share"] and (nc["wr_delta_r2"].iloc[2] < 0.01)
    assert top_u[:4] == [WAGE + R, PCI + R, RENT + R, "female_college_share"]
    screen_txt = f"""Each variable in the summary table was regressed on its own against the 2022 weekly preschool price (N = {len(a22):,}). Each non-cost variable was also added to the wage + rent model (c), which has R² = {sc['base_r2']:.3f}. The table is sorted by the stand-alone R². The **ΔR²** column shows how much a variable adds **beyond** wages and rent.

{sc['md']}

*Table file: `output/tables/predictor_screen_2022.csv`.* Scaling changes the coefficients but not R².

**Highest stand-alone R².** The top four are:

1. childcare wage ({tb.loc[WAGE + R, 'uni_r2']:.3f})
2. per-capita income ({tb.loc[PCI + R, 'uni_r2']:.3f})
3. median rent ({tb.loc[RENT + R, 'uni_r2']:.3f})
4. female college share ({tb.loc['female_college_share', 'uni_r2']:.3f})

Everything else explains under {100 * tb.iloc[4]['uni_r2'] + 1:.0f}% of the variation on its own.

**Best additions to wage + rent.**

- **Per-capita income:** R² rises from {sc['base_r2']:.3f} to **{tb.loc[PCI + R, 'wr_r2']:.3f}** (ΔR² = {tb.loc[PCI + R, 'wr_delta_r2']:+.3f}). Coefficient: ${tb.loc[PCI + R, 'wr_coef']:.2f}/week per $1,000 of income, p {pv(tb.loc[PCI + R, 'wr_p'])}.
- **Female college share:** R² rises to **{tb.loc['female_college_share', 'wr_r2']:.3f}** (ΔR² = {tb.loc['female_college_share', 'wr_delta_r2']:+.3f}). Coefficient: ${tb.loc['female_college_share', 'wr_coef']:.2f}/week per percentage point, p {pv(tb.loc['female_college_share', 'wr_p'])}.
- **The rest:** every other variable adds less than 0.01 to R². That includes the population and child measures tried earlier. Female college share is the variable used in models (d) and (e) above.

**What this means:** prices are explained by local costs plus how affluent and educated a county is, not by how many people or young children it has. These are associations, not causal effects. Income and college share are also correlated with each other and with wages and rent, so their coefficients would change if both were added together."""
    pv = lambda p: "< 0.001" if p < 0.001 else "= " + format(p, ".3f")
    s = summ.set_index("variable")
    st = q1["steps"]
    py = q1["per_year"]
    drop = lambda lab: int(st.loc[st["step"] == lab, "dropped"].iloc[0])

    # guards for the narrative
    assert ma["b"][1] > 0 and mb["b"][1] > 0 and mc["b"][1] > 0 and mc["b"][2] > 0
    assert ovb["corr"] > 0 and ovb["wage_change_matches"] and ovb["rent_change_matches"]
    assert ma["b"][1] > mc["b"][1] and mb["b"][1] > mc["b"][2]
    assert max(ma["p"][1], mb["p"][1], mc["p"][1], mc["p"][2]) < 0.001
    assert bp["wage"]["p"] < 0.05 and bp["rent"]["p"] < 0.05
    assert ma["r2"] > mb["r2"] and mc["r2"] > ma["r2"]
    # guards for the female-college-share text
    assert md_["b"][1] > 0 and me["b"][3] > 0 and md_["p"][1] < 0.001 and me["p"][3] < 0.001
    assert kd["corr_wage"] > 0 and kd["corr_rent"] > kd["corr_wage"] and kd["via_wage"] > 0 and kd["via_rent"] > 0
    assert kd["kids_d"] > kd["kids_e"] and kd["bias_c_rent"] > 0 and kd["bias_c_wage"] < 0
    assert me["b"][2] < mc["b"][2] and me["b"][1] > mc["b"][1] and me["p"][2] < 0.001
    assert me["r2"] > mc["r2"]
    assert mc["b"][1] * extra["sd_wage_k"] > mc["b"][2] * extra["sd_rent_h"]
    assert (mb["b"][1] - mc["b"][2]) / mb["b"][1] > (ma["b"][1] - mc["b"][1]) / ma["b"][1]
    top5w = set(reg["top_w"]["county_fips"]); top5r = set(reg["top_r"]["county_fips"])
    assert {"51013", "51510", "27145"} <= top5w and {"51013", "51510", "27145"} <= top5r

    def ctab(df):
        cols = list(df.columns)
        out = ["| " + " | ".join(str(c) for c in cols) + " |", "|" + "---|" * len(cols)]
        for _, r in df.iterrows():
            out.append("| " + " | ".join(
                str(v) if c in ("Year", "State") else (f"{v:,}" if isinstance(v, (int, np.integer)) else str(v))
                for c, v in zip(cols, r.values)) + " |")
        return "\n".join(out)

    steps_md = st.rename(columns={y: str(y) for y in YEARS}).copy()
    steps_md.columns = ["Step", "Dropped", "Rows left"] + [str(y) for y in YEARS]
    py_md = py.reset_index().rename(columns={"year": "Year", "counties": "Counties",
                                             "states": "States", "metro_areas": "OEWS metro areas"})
    ex = q1["ndcp_unmatched_examples"].head(4)
    ex_txt = "; ".join(f"{r.ndcp_county_name}, {r.ndcp_state} ({r.county_fips}, {r.year}: "
                       f"{r.area_title}, area {r.oews_area})" for r in ex.itertuples())
    reasons = ", ".join(f"{v:,} {k}" for k, v in q1["ndcp_unmatched_reasons"].items())

    tw = a22.assign(res=ma["resid"]); tr = a22.assign(res=mb["resid"]); tc = a22.assign(res=mc["resid"])
    def names(df, n=5, high=True):
        d_ = df.sort_values("res", ascending=not high).head(n)
        return "; ".join(f"{label_county(r)} ({r['area_title']}): ${r['y']:.2f}/wk, "
                         f"{'+' if r['res'] >= 0 else '−'}${abs(r['res']):.0f} vs. fitted"
                         for _, r in d_.iterrows())
    n_wage_vals = a22["wage_k"].nunique()
    sd_w, sd_r = extra["sd_wage_k"], extra["sd_rent_h"]
    p_pos_w = (tw["res"] > 0).mean()
    resid_q = {k: (np.percentile(v["res"], 1), np.percentile(v["res"], 99)) for k, v in (("w", tw), ("r", tr))}
    small = a22.nsmallest(1, "pop_total").iloc[0]
    mps, sc = q1["miss_price_by_state"], q1["state_counties"]
    full = [(stt, y) for stt in mps.index for y in mps.columns if mps.loc[stt, y] == sc.loc[stt, y]]
    full_all = [stt for stt in mps.index if all(mps.loc[stt, y] == sc.loc[stt, y] for y in YEARS if y in mps.columns)
                and set(mps.columns) >= set(YEARS)]
    tot_full = sum(int(mps.loc[stt, y]) for stt, y in full)
    miss_txt = (f"{tot_full:,} of the {int(mps.values.sum()):,} missing-price county-years ({100 * tot_full / mps.values.sum():.0f}%) "
                f"are state-years where NDCP has no metro-county price at all"
                + (f"; states with no price in any year: {', '.join(full_all)}" if full_all else ""))
    miss_tab = mps[mps.sum(axis=1) >= 20].copy()
    miss_tab["Total"] = miss_tab.sum(axis=1)
    miss_tab = miss_tab.sort_values("Total", ascending=False).reset_index().rename(columns={"state_abbr": "State"})
    miss_tab.columns = [str(c) for c in miss_tab.columns]
    pct = lambda new, old: 100 * (old - new) / old

    md = f"""# MGT-403 Data Project — Pre-Analysis Project Data Work: Prices and Costs

*Generated by `code/pre_analysis.py`. Every number in this document is written by the script from its own output; re-running the script regenerates this file, all tables, and all figures.*

**Sample:** U.S. counties in OEWS **metropolitan** areas (OEWS area code beginning `00`), excluding Puerto Rico. **Price** = NDCP `MCPRESCHOOL`, the median weekly full-time price of center-based care for preschoolers (36–54 months). **Dollar values** are in constant 2022 dollars (CPI-U).

---

## 1. Analysis dataset

The NDCP county prices (2015–2022) were merged onto the Problem Set 5 year × county metro panel (OEWS childcare-worker wages + ACS variables) on **[year, county FIPS]**, with FIPS as a 5-character zero-padded string and year as an integer in both files.

| | Rows |
|---|---:|
| NDCP, 2015–2022 | {q1['ndcp_rows']:,} ({q1['ndcp_counties']:,} counties) |
| PS5 metro panel | {q1['ps5_rows']:,} ({q1['ps5_counties']:,} counties) |
| Matched | {q1['matched']:,} |
| PS5 rows not found in NDCP | {q1['ps5_unmatched']:,} |
| NDCP rows not found in PS5 | {q1['ndcp_unmatched']:,} |

Every PS5 county-year found a match in NDCP. The unmatched NDCP rows break down as: {reasons}. These are excluded by design, because the analysis is restricted to metropolitan counties. Examples: {ex_txt}.

The **analysis dataset** is every year × county row where the preschool price **and** all ten predictors (childcare-worker mean annual wage, per-capita income, total population, children under 6, children under 6 with all parents in the labor force, median gross rent, female college share, female full-time-year-round share, under-6 share, under-6 care-need share) are non-missing. Rows dropped at each step (year columns = rows dropped in that year):

{ctab(steps_md)}

Result: **{q1['n_rows']:,} county-years** covering **{q1['n_counties']:,} distinct counties**.

{ctab(py_md)}

Missing prices cluster by state and year rather than being scattered across counties: {miss_txt}. Below are the states with at least 20 missing-price county-years (counts are metro county-years with no NDCP price; Puerto Rico excluded):

{ctab(miss_tab)}

The large 2021 gap is why 2021 has the fewest counties. Missing wages ({drop('drop missing ' + WAGE)} rows) reflect BLS suppression of individual area-years; missing rent ({drop('drop missing ' + RENT)} rows) reflects ACS suppression. No other predictor had missing values.

---

## 2. Summary statistics, 2022 analysis dataset

N = {len(a22):,} metro counties. SD is the sample standard deviation. Dollar values are 2022 dollars (2022 is the base year, so 2022 values equal nominal values, except Connecticut's carried-forward rent and income; see Data handling).

{summ_md}

*Table file: `output/tables/summary_2022.csv`.* Note that the minimum population ({int(small['pop_total']):,}) is {label_county(small)}, a very small county that sits inside the {small['area_title']} metro area; it meets the sample definition and is kept.

The typical metro county charges **${s.loc[PRICE + R, 'mean']:.2f}/week** for center-based preschool (SD ${s.loc[PRICE + R, 'sd']:.2f}), or about ${52 * s.loc[PRICE + R, 'mean']:,.0f} per year at 52 weeks. Childcare workers earn on average **${s.loc[WAGE + R, 'mean']:,.0f}** per year (SD ${s.loc[WAGE + R, 'sd']:,.0f}). Median gross rent averages **${s.loc[RENT + R, 'mean']:,.0f}/month** (SD ${s.loc[RENT + R, 'sd']:,.0f}). Prices vary a lot: the most expensive county charges {s.loc[PRICE + R, 'max'] / s.loc[PRICE + R, 'min']:.1f}× the cheapest (${s.loc[PRICE + R, 'max']:.2f} vs. ${s.loc[PRICE + R, 'min']:.2f}).

---

## 3. Scatter plots, 2022

In both plots the line is the OLS fit. Slopes use wages in **$1,000s per year** and rent in **$100s per month**. SEs are heteroskedasticity-robust (HC1). The five largest absolute residuals are labeled.

![Price vs wage](figures/scatter_price_vs_wage_2022.png)

**(a) Price vs. childcare-worker wage.** The relationship is **positive and strong**. Each $1,000 of higher annual childcare wage goes with **${ma['b'][1]:.2f} higher weekly price** (HC1 SE {ma['se'][1]:.2f}), or about ${52 * ma['b'][1]:,.0f} per child per year. A one-SD difference in wages (${1000 * sd_w:,.0f}) corresponds to roughly **${ma['b'][1] * sd_w:.0f}/week**. Wages alone explain **{100 * ma['r2']:.1f}%** of the cross-county variation in price (R² = {ma['r2']:.3f}).

- **Wage varies by metro area, not by county.** The 2022 sample has only {n_wage_vals} distinct wage values for {len(a22):,} counties, so the points form vertical stripes. Every county in a metro area shares one wage but can have a very different price.
- **Outliers.** The largest outliers are counties priced far **above** the line: {names(tw, 5, True)}. The most under-priced counties relative to the line are {names(tw, 3, False)}.
- **The residuals are right-skewed.** The 99th-percentile residual is +${resid_q['w'][1]:.0f}; the 1st percentile is −${abs(resid_q['w'][0]):.0f}.
- **Heteroskedasticity.** The scatter fans out as wages rise. A Breusch–Pagan test rejects constant variance (LM = {bp['wage']['lm']:.1f}, p {('< 0.001' if bp['wage']['p'] < 0.001 else '= ' + format(bp['wage']['p'], '.3f'))}). The robust SE on the slope is {extra['hc1_vs_classical']['a']:.2f}× the classical SE.

![Price vs rent](figures/scatter_price_vs_rent_2022.png)

**(b) Price vs. median gross rent.** The relationship is also **positive and strong**. Each $100/month of higher median rent goes with **${mb['b'][1]:.2f} higher weekly price** (HC1 SE {mb['se'][1]:.2f}). A one-SD difference in rent (${100 * sd_r:,.0f}/month) corresponds to roughly **${mb['b'][1] * sd_r:.0f}/week**. Rent alone explains **{100 * mb['r2']:.1f}%** of the variation (R² = {mb['r2']:.3f}), less than wages do.

- **Outliers.** The largest positive outliers are {names(tr, 5, True)}. The most under-priced counties relative to the line are {names(tr, 3, False)}.
- **The residuals are right-skewed.** The 99th-percentile residual is +${resid_q['r'][1]:.0f}; the 1st percentile is −${abs(resid_q['r'][0]):.0f}.
- **Heteroskedasticity.** The spread widens at high rents (Breusch–Pagan LM = {bp['rent']['lm']:.1f}, p {('< 0.001' if bp['rent']['p'] < 0.001 else '= ' + format(bp['rent']['p'], '.3f'))}). The robust SE is {extra['hc1_vs_classical']['b']:.2f}× the classical SE.

**Interpretation.** Higher-cost markets charge more, which fits the idea that prices partly reflect labor and space costs. These are associations, not causal effects. High-wage and high-rent counties are also richer, more educated, and denser, so they likely have stronger demand. A wage effect of ${52 * ma['b'][1]:,.0f} per child per year for each $1,000 of annual wage per worker looks larger than a pure cost pass-through would produce, since each worker cares for several children. So the wage slope probably also picks up demand and local affluence.

For the location question, the counties far above the line (Arlington and Alexandria, VA, and St. Cloud, MN, for example) are early candidates for "prices high relative to costs." Because wages are measured at the metro level, residuals within one metro area mostly reflect county-level price differences.

---

## 4. Regressions, 2022

**Dependent variable:** median weekly preschool price (2022 $). **Wage** is in **$1,000s per year** and **rent** is in **$100s per month**, both in 2022 dollars. **Female college share** (`female_college_share` = women 25+ with a bachelor's degree or higher ÷ all women 25+, ACS) is in **percentage points**. Models (d) and (e) add this demand-side variable, chosen because it had the highest stand-alone R² of the non-cost variables in the predictor screen below, to the three cost models the assignment asks for. SEs are heteroskedasticity-robust **HC1** (White, with the n/(n−k) small-sample correction). p-values use the t distribution with n − k degrees of freedom.

{reg['md']}

*Table files: `output/tables/regressions_2022.csv` and `.md`.*

**Correlation between childcare wage and median gross rent (2022):** r = **{ovb['corr']:.3f}**.

**(a) Wage only.** A $1,000 higher annual childcare-worker wage is associated with a **${ma['b'][1]:.2f}** higher weekly preschool price (t = {ma['t'][1]:.1f}, p {('< 0.001' if ma['p'][1] < 0.001 else '= ' + format(ma['p'][1], '.3f'))}). The intercept ({ma['b'][0]:.1f}) is the extrapolated price at a $0 wage. It has no economic meaning because the lowest observed wage is ${1000 * a22['wage_k'].min():,.0f}.

**(b) Rent only.** A $100/month higher median gross rent is associated with a **${mb['b'][1]:.2f}** higher weekly price (t = {mb['t'][1]:.1f}, p {('< 0.001' if mb['p'][1] < 0.001 else '= ' + format(mb['p'][1], '.3f'))}).

**(c) Wage and rent together.**

- **Wage, holding rent constant:** a $1,000 higher annual childcare wage is associated with a **${mc['b'][1]:.2f}** higher weekly price (t = {mc['t'][1]:.1f}).
- **Rent, holding wage constant:** a $100/month higher median rent is associated with a **${mc['b'][2]:.2f}** higher weekly price (t = {mc['t'][2]:.1f}).
- **Fit:** both remain highly significant (p < 0.001), and R² rises to **{mc['r2']:.3f}** from {ma['r2']:.3f} (wage only) and {mb['r2']:.3f} (rent only).
- **Equal-SD comparison:** a one-SD increase in wage goes with ${mc['b'][1] * sd_w:.0f}/week and a one-SD increase in rent with ${mc['b'][2] * sd_r:.0f}/week, so wages are the stronger of the two cost signals.

**Why the coefficients change: omitted variable bias.** Each univariate model leaves out the other cost variable. The coefficient on the included variable therefore also picks up part of the omitted variable's effect:

  short coefficient = long coefficient + (omitted variable's long coefficient) × (slope of omitted on included)

- **Sign of the bias:** wages and rent are **positively correlated** (r = {ovb['corr']:.3f}), and each has a **positive** effect in model (c). Both univariate coefficients should therefore be biased **upward**, and both should fall when the other variable is added.
- **Wage coefficient:** regressing rent ($100s) on wage ($1,000s) gives a slope of {ovb['delta_rent_on_wage']:.3f}. The predicted bias is {ovb['rent_c']:.3f} × {ovb['delta_rent_on_wage']:.3f} = **+{ovb['bias_wage']:.3f}**. Indeed, {ovb['wage_a']:.3f} (a) = {ovb['wage_c']:.3f} (c) + {ovb['bias_wage']:.3f}. The wage coefficient falls by {pct(ovb['wage_c'], ovb['wage_a']):.0f}%.
- **Rent coefficient:** regressing wage on rent gives a slope of {ovb['delta_wage_on_rent']:.3f}. The predicted bias is {ovb['wage_c']:.3f} × {ovb['delta_wage_on_rent']:.3f} = **+{ovb['bias_rent']:.3f}**. Indeed, {ovb['rent_b']:.3f} (b) = {ovb['rent_c']:.3f} (c) + {ovb['bias_rent']:.3f}. The rent coefficient falls by {pct(ovb['rent_c'], ovb['rent_b']):.0f}%.
- **Check:** the observed direction of change matches the OVB prediction for both coefficients (both fall). The script verifies the decomposition holds exactly in the data.

The rent coefficient shrinks by more than the wage coefficient. A large part of the rent-only slope reflects the fact that high-rent counties are also high-wage counties, while wages keep most of their association with price after controlling for rent.

**(d) Female college share, alone.** A county where the share of women 25+ with a bachelor's degree is one percentage point higher has a **${md_['b'][1]:.2f}** higher weekly preschool price on average (HC1 SE {md_['se'][1]:.2f}, t = {md_['t'][1]:.2f}, p {pv(md_['p'][1])}). College share alone explains **{100 * md_['r2']:.1f}%** of the variation in price (R² = {md_['r2']:.3f}). That is less than wage ({ma['r2']:.3f}) or rent ({mb['r2']:.3f}) alone, but far more than the population and child measures tried earlier.

![Price vs female college share](figures/scatter_price_vs_female_college_share_2022.png)

- **Size of a typical difference:** the share averages {kd['mean']:.1f}% (SD {kd['sd']:.1f} points). A one-SD difference goes with about **${md_['b'][1] * kd['sd']:.0f}/week**. Moving from the 10th percentile ({kd['p10']:.1f}%) to the 90th ({kd['p90']:.1f}%) implies ${md_['b'][1] * (kd['p90'] - kd['p10']):.0f}/week more.
- **What the plot shows:** a clear upward slope. The spread widens at higher college shares, the same fan shape seen in the wage and rent plots.

**(e) Wage, rent and female college share together.**

- **College share:** **${me['b'][3]:.2f}** per percentage point, holding wage and rent constant (HC1 SE {me['se'][3]:.2f}, t = {me['t'][3]:.2f}, p {pv(me['p'][3])}). About half its stand-alone size, but still highly significant.
- **Wage:** **${me['b'][1]:.2f}** per $1,000, holding rent and college share constant (t = {me['t'][1]:.1f}). Slightly **up** from ${mc['b'][1]:.2f} in model (c).
- **Rent:** **${me['b'][2]:.2f}** per $100/month, holding wage and college share constant (t = {me['t'][2]:.1f}, p {pv(me['p'][2])}). **Down sharply** from ${mc['b'][2]:.2f} in model (c), a {pct(me['b'][2], mc['b'][2]):.0f}% drop.
- **Fit:** R² rises from {mc['r2']:.3f} to **{me['r2']:.3f}**. College share adds explanatory power beyond the two cost variables.

**Why the college coefficient falls from (d) to (e) (OVB).** College share is positively correlated with wages (r = {kd['corr_wage']:.3f}) and especially with rent (r = {kd['corr_rent']:.3f}). Model (d) therefore credits college share with part of the cost effects:

  {kd['kids_d']:.3f} (d) = {kd['kids_e']:.3f} (e) + {me['b'][1]:.3f} × {kd['delta_wage_on_kids']:.4f} (via wage) + {me['b'][2]:.3f} × {kd['delta_rent_on_kids']:.4f} (via rent) = {kd['kids_e']:.3f} + {kd['via_wage']:.3f} + {kd['via_rent']:.3f}

Here {kd['delta_wage_on_kids']:.4f} and {kd['delta_rent_on_kids']:.4f} are the slopes from regressing wage ($1,000s) and rent ($100s) on college share (percentage points). Both terms are positive, so the stand-alone coefficient is biased **upward**. It falls when wage and rent are added, as OVB predicts.

**Why the rent coefficient falls from (c) to (e) (OVB the other way round).** Model (c) omitted college share, which raises price and is correlated with rent. Regressing college share on wage and rent together gives:

- **Rent:** a coefficient of {kd['g_rent']:.3f} percentage points per $100 of rent, holding wage fixed. Positive.
- **Wage:** a coefficient of {kd['g_wage']:.3f} points per $1,000 of wage, holding rent fixed. Slightly negative. Among counties with the same rent, higher-wage ones have marginally *lower* college shares.

The omitted-variable bias in model (c) is therefore:

- **Rent coefficient:** {me['b'][3]:.3f} × {kd['g_rent']:.3f} = **+{kd['bias_c_rent']:.3f}**, so {mc['b'][2]:.3f} (c) = {me['b'][2]:.3f} (e) + {kd['bias_c_rent']:.3f}. Biased upward, and it falls when college share is added.
- **Wage coefficient:** {me['b'][3]:.3f} × ({kd['g_wage']:.3f}) = **{kd['bias_c_wage']:.3f}**, so {mc['b'][1]:.3f} (c) = {me['b'][1]:.3f} (e) + ({kd['bias_c_wage']:.3f}). Biased slightly downward, and it rises a little when college share is added.

Both observed changes match the direction OVB predicts, and the script verifies all three decompositions hold exactly.

**What this means.** Much of what looked like a "rent effect" in model (c) was really an affluence and education effect: high-rent counties are places with many college-educated women, and those places pay more for preschool. Once that is controlled for, childcare wages are the dominant cost signal and rent matters much less. College share is a demand-side variable, so its coefficient reflects both willingness to pay and the kind of care these markets buy. It is an association, not a causal effect.

### Predictor screen: which summary-table variables explain price best? (2022)

{screen_txt}

Both coefficients in model (c) may still be biased by omitted demand variables, such as income and education, that are correlated with both costs and prices.

Wage is measured at the metro level, so counties in the same metro area are not independent observations. Standard errors clustered by metro area would be a natural robustness check. HC1 does not account for this.

---

## Data handling decisions

1. **Sources.**
   - **Prices:** U.S. DOL National Database of Childcare Prices, county workbook `NDCP2022.xlsx`, sheet `County_LevelNDCP_v8_update2008_`.
   - **Wages and ACS variables:** the Problem Set 5 panel `ps5_county_panel_2015_2022.csv`. It contains OEWS SOC 39-9011 mean annual wage by metro area, merged down to counties with the course crosswalk `county_oews_crosswalk.csv`, plus ACS 5-year county variables.
2. **Price variable.** `MCPRESCHOOL`: the median weekly full-time price of **center-based** care for **preschool** age (36–54 months). The NDCP data dictionary calls it "MCPreschool"; the column in the workbook is `MCPRESCHOOL`. `MCPreschool_flag` describes how the age-band price was aggregated (1 = all bands equal, 2 = mode, 3 = highest mode). It is kept for reference and is **not** used to exclude rows.
3. **Years.** 2015–2022 (`STUDYYEAR` in NDCP; `year` in PS5), stored as integers in both files.
4. **Merge keys and FIPS padding.**
   - The merge is on `[year, county_fips]`, one-to-one (validated).
   - NDCP stores FIPS as a number (e.g., `1001`). It was read as text and zero-padded to 5 characters (`01001`). PS5 FIPS were also read as text and padded.
   - OEWS area codes are kept as 7-character text.
   - The output CSV stores FIPS as text, and the script asserts every code is 5 digits on re-read.
5. **Metro sample.** Counties whose year-specific OEWS area code starts with `00`, inherited from PS5. Nonmetro NDCP counties ({q1['ndcp_unmatched']:,} county-years) are therefore not matched.
6. **Puerto Rico dropped** ({drop('drop Puerto Rico (state FIPS 72)')} county-years). NDCP has PR prices only for {", ".join(str(y) for y in q1['pr_price_years'])}. For {len(q1['name_mismatch_fips'])} PR municipios ({", ".join(q1['name_mismatch_fips'])}), the NDCP FIPS code carries a different municipio's name than Census/ACS, so the price-to-county mapping is unreliable. No name mismatches were found outside Puerto Rico.
7. **ACS vintage.** NDCP/OEWS year *t* is matched to the ACS **5-year vintage ending in year *t*** (for example, 2022 ↔ ACS 2018–2022), as in PS5.
8. **Inflation adjustment.**
   - All dollar variables (price, wage, rent, per-capita income) are converted to **constant 2022 dollars** with the CPI-U, U.S. city average, all items (BLS series CUUR0000SA0), annual averages: {", ".join(f"{y} {c:.3f}" for y, c in CPI.items())}.
   - Real value = nominal × CPI₂₀₂₂ / CPI_t.
   - For prices and wages, *t* is the data year.
   - For ACS rent and income, *t* is the **ACS vintage year**, because ACS 5-year dollars are expressed in the vintage's final-year dollars.
   - Source: BLS, Historical CPI-U table (https://www.bls.gov/cpi/tables/supplemental-files/historical-cpi-u-202212.pdf).
9. **Connecticut 2022.** The 2022 ACS reports Connecticut by planning region rather than county. PS5 therefore carries forward 2021-vintage ACS values for the legacy counties (flag `acs_ct_2021_carryforward`). These {q1['flags']['ct']} rows are kept, and their rent and income are deflated with the **2021** CPI, consistent with the vintage rule above.
10. **Missing data.** Rows are dropped if the price or any of the ten predictors is missing. The drop counts at each step are in Section 1. No values are imputed.
11. **Flagged but kept.**
    - {q1['low_outlier_name']} in {LOW_OUTLIER[1]} has a price of ${q1['low_outlier_price']:.2f}/week (nominal). It is {'the only' if q1['n_price_under_50'] == 1 else f"one of {q1['n_price_under_50']}"} analysis county-year{'' if q1['n_price_under_50'] == 1 else 's'} under $50 (flag `flag_low_price_outlier`).
    - {q1['flags']['divisions']} county-years have 2015–2017 wages rebuilt in PS5 as employment-weighted averages of OEWS Metropolitan Division means (flag `wage_from_divisions`).
12. **Regression conventions.** Wages are in $1,000s per year and rent in $100s per month, both in 2022 $. Female college share is in percentage points. OLS with HC1 robust SEs and t-distribution p-values. The OLS/HC1 routine is hand-coded in numpy. It is checked against closed-form univariate formulas, and against statsmodels when that package is installed.

## Files

- `data/processed/analysis_dataset.csv`: analysis dataset, all years, with nominal and 2022-$ columns.
- `output/tables/summary_2022.csv` and `summary_2022.md`: Q2 summary table.
- `output/tables/regressions_2022.csv` and `regressions_2022.md`: Q4 regression table.
- `output/tables/predictor_screen_2022.csv` and `predictor_screen_2022.md`: one-at-a-time R² screen of all summary-table variables.
- `output/figures/scatter_price_vs_wage_2022.png` and `scatter_price_vs_rent_2022.png`: Q3 scatter plots.
- `output/figures/scatter_price_vs_female_college_share_2022.png`: price vs. female college share (model d).
- `output/pre_analysis_log.txt`: full console log of the run.
"""
    with open(os.path.join(ROOT, "output", "pre_analysis_writeup.md"), "w") as fh:
        fh.write(md)


if __name__ == "__main__":
    main()
