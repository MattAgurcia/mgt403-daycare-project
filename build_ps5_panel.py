"""
PS5 step 3: build the year x county metro panel.
  county_oews_crosswalk.csv  (county -> OEWS area, by year)
  + oews_childcare_39-9011_2015_2022_long.csv  (childcare-worker wages by area-year)
  + acs_county_2015_2022.csv  (ACS county demand/cost variables)

Usage (from this folder):  python3 build_ps5_panel.py
Output: ps5_county_panel_2015_2022.csv / .xlsx

Geography fixes applied (documented in the output via flag columns):
  1. Miami-Dade, FL: crosswalk carries both the current FIPS 12086 and the
     retired code 12025 (changed in 1997), which matches nothing in ACS or
     NDCP -> drop the duplicate 12025 rows.
  2. Bedford city, VA (51515): merged into Bedford County (51019) in 2013;
     no longer in ACS or NDCP -> dropped from the crosswalk.
  3. Connecticut 2022: the 2022 ACS reports 9 planning regions (09110-09190)
     instead of the 8 legacy counties used by NDCP and the crosswalk. Medians
     (rent, income) cannot be re-aggregated, so CT 2022 rows use the 2021 ACS
     5-year values (2017-2021) for the legacy counties; flagged with
     acs_ct_2021_carryforward = 1.
  4. OEWS 2015-2017: the 11 largest metros are published only by Metropolitan/
     NECTA Division; MSA-level mean wages are rebuilt as the employment-weighted
     average of division means; flagged with wage_from_divisions = 1.
"""
import os
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
p = lambda f: os.path.join(HERE, f)

# ---------------- crosswalk ----------------
xw = pd.read_csv(p("county_oews_crosswalk.csv"), dtype=str)
xw["year"] = xw["year"].astype(int)
xw["county_fips"] = xw["county_fips"].str.zfill(5)
xw["oews_area"] = xw["oews_area"].str.zfill(7)

n_miami = (xw["county_fips"] == "12025").sum()
assert (xw["county_fips"] == "12086").sum() > 0, "expected 12086 rows in crosswalk"
xw = xw[xw["county_fips"] != "12025"]                                   # fix 1
n_bedford = (xw["county_fips"] == "51515").sum()
xw = xw[xw["county_fips"] != "51515"]                                   # fix 2
assert not xw.duplicated(["year", "county_fips"]).any(), "duplicate county-years in crosswalk"

metro = xw[xw["oews_area"].str.startswith("00")].copy()                # metro analysis sample
print(f"Crosswalk: dropped {n_miami} duplicate Miami-Dade (12025) rows, dropped {n_bedford} Bedford city rows")
print(f"Metro sample: {len(metro):,} county-years")

# ---------------- OEWS wages ----------------
oews = pd.read_csv(p("oews_childcare_39-9011_2015_2022_long.csv"), dtype={"area_code": str})
oews["oews_area"] = oews["area_code"].str.strip().str.zfill(7)        # 7-char, keep leading zeros
oews = oews.rename(columns={"a_mean": "childcare_wage_a_mean",
                            "h_mean": "childcare_wage_h_mean",
                            "a_median": "childcare_wage_a_median",
                            "tot_emp": "childcare_tot_emp"})
oews = oews[["year", "oews_area", "childcare_wage_a_mean", "childcare_wage_h_mean",
             "childcare_wage_a_median", "childcare_tot_emp"]]
oews["wage_from_divisions"] = 0

# Fix 4: in the 2015-2017 OEWS files, the 11 largest MSAs/NECTAs are published
# only by Metropolitan (or NECTA) Division, while the crosswalk uses the parent
# MSA code. Rebuild the MSA-level mean wage as the employment-weighted average
# of its divisions' means (= the MSA mean when all divisions are published).
# Medians cannot be aggregated this way, so a_median is left missing for these.
DIVISIONS = {
    "0016980": ["16974", "20994", "23844", "29404"],                   # Chicago
    "0019100": ["19124", "23104"],                                     # Dallas-Fort Worth
    "0019820": ["19804", "47664"],                                     # Detroit
    "0031080": ["11244", "31084"],                                     # Los Angeles
    "0033100": ["22744", "33124", "48424"],                            # Miami
    "0035620": ["20524", "35004", "35084", "35614"],                   # New York
    "0037980": ["15804", "33874", "37964", "48864"],                   # Philadelphia
    "0041860": ["36084", "41884", "42034"],                            # San Francisco
    "0042660": ["42644", "45104"],                                     # Seattle
    "0047900": ["43524", "47894"],                                     # Washington DC
    "0071650": ["71654", "72104", "73104", "73604", "74204", "74804",  # Boston NECTA
                "74854", "75404", "76524", "78254"],
}
div_rows = []
for msa, divs in DIVISIONS.items():
    divs7 = [d.zfill(7) for d in divs]
    for y in sorted(oews["year"].unique()):
        if ((oews["year"] == y) & (oews["oews_area"] == msa)).any():
            continue  # MSA already published directly (2018+)
        sub = oews[(oews["year"] == y) & oews["oews_area"].isin(divs7)].dropna(
            subset=["childcare_wage_a_mean"])
        if sub.empty:
            continue
        w = sub["childcare_tot_emp"]
        if w.notna().all() and w.sum() > 0:
            a = (sub["childcare_wage_a_mean"] * w).sum() / w.sum()
            h = (sub["childcare_wage_h_mean"] * w).sum() / w.sum()
            emp = w.sum()
        else:  # employment suppressed for a division -> simple average
            a, h, emp = sub["childcare_wage_a_mean"].mean(), sub["childcare_wage_h_mean"].mean(), pd.NA
        div_rows.append({"year": y, "oews_area": msa, "childcare_wage_a_mean": round(a),
                         "childcare_wage_h_mean": round(h, 2), "childcare_wage_a_median": float("nan"),
                         "childcare_tot_emp": emp, "wage_from_divisions": 1,
                         "n_div_used": len(sub), "n_div_total": len(divs)})
div_df = pd.DataFrame(div_rows)
print(f"OEWS: rebuilt {len(div_df)} MSA-years from Metropolitan/NECTA Divisions (2015-2017)")
incomplete = div_df[div_df["n_div_used"] < div_df["n_div_total"]]
if len(incomplete):
    print("  note: some divisions suppressed for:")
    print(incomplete[["year", "oews_area", "n_div_used", "n_div_total"]].to_string(index=False))
oews = pd.concat([oews, div_df.drop(columns=["n_div_used", "n_div_total"])], ignore_index=True)
assert not oews.duplicated(["year", "oews_area"]).any(), "duplicate area-years in OEWS"

# ---------------- ACS ----------------
acs = pd.read_csv(p("acs_county_2015_2022.csv"), dtype={"county_fips": str, "state_fips": str})
acs["acs_ct_2021_carryforward"] = 0
ct_legacy = acs[(acs["year"] == 2021) & acs["county_fips"].str.match(r"^090(0[1-9]|1[0-5])$")].copy()
ct_legacy["year"] = 2022
ct_legacy["acs_ct_2021_carryforward"] = 1                               # fix 3
acs = acs[~((acs["year"] == 2022) & acs["county_fips"].str.startswith("09"))]
acs = pd.concat([acs, ct_legacy], ignore_index=True)
print(f"ACS: carried forward 2021 values for {len(ct_legacy)} CT legacy counties in 2022")

acs_keep = ["year", "county_fips", "county_name", "per_capita_income", "pop_total",
            "kids_under6", "kids_under6_allpar_lf", "median_gross_rent",
            "female_college_share", "female_ftyr_share", "under6_share",
            "under6_careneed_share", "acs_ct_2021_carryforward"]

# ---------------- merge ----------------
panel = (metro.merge(oews, on=["year", "oews_area"], how="left", validate="m:1")
              .merge(acs[acs_keep], on=["year", "county_fips"], how="left",
                     validate="1:1", indicator="_acs"))
panel["acs_ct_2021_carryforward"] = panel["acs_ct_2021_carryforward"].fillna(0).astype(int)

cols = ["year", "county_fips", "county_name", "state_abbr", "oews_area", "area_title",
        "childcare_wage_a_mean", "childcare_wage_h_mean", "childcare_wage_a_median",
        "childcare_tot_emp", "wage_from_divisions"] + acs_keep[3:]
panel = panel.sort_values(["year", "county_fips"])
unmatched_acs = panel[panel["_acs"] == "left_only"]
panel = panel[cols].reset_index(drop=True)

panel.to_csv(p("ps5_county_panel_2015_2022.csv"), index=False)
try:
    panel.to_excel(p("ps5_county_panel_2015_2022.xlsx"), index=False, sheet_name="panel")
except ImportError:
    pass

# ---------------- sanity checks ----------------
print(f"\nWrote ps5_county_panel_2015_2022.csv  ({len(panel):,} rows, "
      f"{panel['county_fips'].nunique():,} counties)")
chk = panel.groupby("year").agg(
    counties=("county_fips", "size"),
    missing_wage=("childcare_wage_a_mean", lambda s: s.isna().sum()),
    missing_rent=("median_gross_rent", lambda s: s.isna().sum()),
    missing_acs=("pop_total", lambda s: s.isna().sum()))
chk["pct_missing_wage"] = (100 * chk["missing_wage"] / chk["counties"]).round(1)
print(chk.to_string())
print(f"\nOverall missing wages: {panel['childcare_wage_a_mean'].isna().mean():.1%} "
      "(assignment expects ~1% from BLS suppression)")
print(f"County-years with no ACS match: {len(unmatched_acs)}")
if len(unmatched_acs):
    print(unmatched_acs[["year", "county_fips", "area_title"]].to_string(index=False))
miss_w = panel[panel["childcare_wage_a_mean"].isna()]
if len(miss_w):
    print("\nArea-years with suppressed/missing wages:")
    print(miss_w.groupby(["year", "oews_area", "area_title"]).size()
          .rename("n_counties").reset_index().to_string(index=False))
