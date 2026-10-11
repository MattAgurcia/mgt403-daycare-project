"""
Pull ACS 5-year county estimates (2015-2022) from the Census API and build
the childcare demand / cost variables defined in the MGT-403 Data Project.

Usage (from this folder, in Terminal):
    python3 build_acs_county.py

Outputs (written next to this script):
    acs_county_2015_2022.csv   - year x county panel, raw ACS codes + derived vars
    acs_county_2015_2022.xlsx  - same data (if openpyxl is installed)

Variables (ACS code -> derived variable), per the assignment table:
    B19301_001E                         -> per_capita_income
    B01001_001E                         -> pop_total
    B23008_002E                         -> kids_under6
    B23008_004E + 010E + 013E           -> kids_under6_allpar_lf
    B25064_001E                         -> median_gross_rent
    (B15002_032E..035E) / B15002_019E   -> female_college_share
    B23022_029E / B23022_026E           -> female_ftyr_share
    kids_under6 / pop_total             -> under6_share
    kids_under6_allpar_lf / pop_total   -> under6_careneed_share

Variable labels verified against the 2015 and 2022 API variable dictionaries:
    B23008_002E  Under 6 years (own children in families/subfamilies)
    B23008_004E  Under 6: living with two parents, both parents in labor force
    B23008_010E  Under 6: living with one parent (father), in labor force
    B23008_013E  Under 6: living with one parent (mother), in labor force
    B15002_019E  Female 25+, total; 032-035E = female bachelor's/master's/
                 professional/doctorate
    B23022_026E  Female 16-64, total; 029E = female worked 35+ hrs/wk, 50-52 wks
"""
import json
import os
import sys
import time
import urllib.request

import pandas as pd

# Census API key: set it in your shell, e.g.  export CENSUS_API_KEY=your_key
# (free signup: https://api.census.gov/data/key_signup.html). Never commit a real key.
API_KEY = os.environ.get("CENSUS_API_KEY")
if not API_KEY:
    sys.exit("Set the CENSUS_API_KEY environment variable first (free key: https://api.census.gov/data/key_signup.html)")
YEARS = range(2015, 2023)
HERE = os.path.dirname(os.path.abspath(__file__))

VARS = [
    "B19301_001E",  # per-capita income
    "B01001_001E",  # total population
    "B23008_002E",  # own children under 6
    "B23008_004E",  # under 6, two parents, both in LF
    "B23008_010E",  # under 6, father only, in LF
    "B23008_013E",  # under 6, mother only, in LF
    "B25064_001E",  # median gross rent
    "B15002_019E",  # female 25+ total
    "B15002_032E",  # female bachelor's
    "B15002_033E",  # female master's
    "B15002_034E",  # female professional
    "B15002_035E",  # female doctorate
    "B23022_026E",  # female 16-64 total
    "B23022_029E",  # female 16-64 full-time year-round
]


def fetch(year):
    url = (f"https://api.census.gov/data/{year}/acs/acs5"
           f"?get=NAME,{','.join(VARS)}&for=county:*&in=state:*&key={API_KEY}")
    for attempt in range(4):
        try:
            with urllib.request.urlopen(url, timeout=120) as r:
                rows = json.loads(r.read().decode("utf-8"))
            df = pd.DataFrame(rows[1:], columns=rows[0])
            df.insert(0, "year", year)
            return df
        except Exception as e:  # network hiccup / rate limit -> retry
            print(f"  {year}: attempt {attempt + 1} failed ({e}); retrying...")
            time.sleep(3 * (attempt + 1))
    sys.exit(f"Could not download {year}. Check your connection / API key.")


frames = []
for y in YEARS:
    print(f"Downloading ACS 5-year {y} ({y - 4}-{y}) ...")
    frames.append(fetch(y))
raw = pd.concat(frames, ignore_index=True)

# FIPS as zero-padded strings (keep leading zeros)
raw["state_fips"] = raw["state"].astype(str).str.zfill(2)
raw["county_fips"] = raw["state_fips"] + raw["county"].astype(str).str.zfill(3)
raw = raw.rename(columns={"NAME": "county_name"}).drop(columns=["state", "county"])

# numeric; Census uses large negative sentinels (e.g. -666666666) for N/A
for v in VARS:
    raw[v] = pd.to_numeric(raw[v], errors="coerce")
    raw.loc[raw[v] < 0, v] = pd.NA
    raw[v] = raw[v].astype("float64")

d = raw
d["per_capita_income"] = d["B19301_001E"]
d["pop_total"] = d["B01001_001E"]
d["kids_under6"] = d["B23008_002E"]
d["kids_under6_allpar_lf"] = d["B23008_004E"] + d["B23008_010E"] + d["B23008_013E"]
d["median_gross_rent"] = d["B25064_001E"]
d["female_college_share"] = (
    d[["B15002_032E", "B15002_033E", "B15002_034E", "B15002_035E"]].sum(axis=1, min_count=4)
    / d["B15002_019E"].where(d["B15002_019E"] > 0)
)
d["female_ftyr_share"] = d["B23022_029E"] / d["B23022_026E"].where(d["B23022_026E"] > 0)
pop = d["pop_total"].where(d["pop_total"] > 0)
d["under6_share"] = d["kids_under6"] / pop
d["under6_careneed_share"] = d["kids_under6_allpar_lf"] / pop

DERIVED = ["per_capita_income", "pop_total", "kids_under6", "kids_under6_allpar_lf",
           "median_gross_rent", "female_college_share", "female_ftyr_share",
           "under6_share", "under6_careneed_share"]
out = d[["year", "county_fips", "state_fips", "county_name"] + DERIVED + VARS]
out = out.sort_values(["year", "county_fips"]).reset_index(drop=True)

csv_path = os.path.join(HERE, "acs_county_2015_2022.csv")
out.to_csv(csv_path, index=False)
print(f"\nWrote {csv_path}  ({len(out):,} rows)")
try:
    xlsx_path = os.path.join(HERE, "acs_county_2015_2022.xlsx")
    out.to_excel(xlsx_path, index=False, sheet_name="acs_county")
    print(f"Wrote {xlsx_path}")
except ImportError:
    print("(openpyxl not installed - skipped .xlsx; the .csv has everything)")

# ---------------- sanity checks ----------------
print("\nCounties per year:")
print(out.groupby("year").size().to_string())
print("\nMissing values per derived variable (all years):")
print(out[DERIVED].isna().sum().to_string())
print("\n2022 summary (all counties):")
print(out[out.year == 2022][DERIVED].describe().T[["count", "mean", "std", "min", "max"]]
      .round(3).to_string())

xw_path = os.path.join(HERE, "county_oews_crosswalk.csv")
if os.path.exists(xw_path):
    xw = pd.read_csv(xw_path, dtype=str)
    xw["year"] = xw["year"].astype(int)
    metro = xw[xw["oews_area"].str.startswith("00")]
    m = metro.merge(out[["year", "county_fips"]], on=["year", "county_fips"],
                    how="left", indicator=True)
    print("\nMetro-sample counties (crosswalk) matched to ACS, by year:")
    print(m.groupby("year")["_merge"].apply(lambda s: f"{(s == 'both').sum()}/{len(s)}")
          .to_string())
    unmatched = m[m["_merge"] == "left_only"]
    if len(unmatched):
        print("Unmatched metro county-years (first 20):")
        print(unmatched[["year", "county_fips", "area_title"]].head(20).to_string(index=False))
        print("Note: from the 2022 ACS on, Connecticut is reported by 9 planning "
              "regions (09110-09190) instead of its 8 legacy counties (09001-09015).")
