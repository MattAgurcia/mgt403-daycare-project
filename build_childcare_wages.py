"""
Build a 2015-2022 panel of BLS OEWS wages for SOC 39-9011 (Childcare Workers)
at the metropolitan (MSA) and nonmetropolitan (BOS) area level.

Usage: python3 build_childcare_wages.py <extracted_oesm_dir> <output_dir>
  <extracted_oesm_dir> contains oesm15ma/ ... oesm22ma/ (unzipped BLS files)
"""
import sys, glob, os, re
import openpyxl
import pandas as pd

SRC, OUT = sys.argv[1], sys.argv[2]
SOC = "39-9011"

FIPS = {"01":"AL","02":"AK","04":"AZ","05":"AR","06":"CA","08":"CO","09":"CT","10":"DE","11":"DC",
        "12":"FL","13":"GA","15":"HI","16":"ID","17":"IL","18":"IN","19":"IA","20":"KS","21":"KY",
        "22":"LA","23":"ME","24":"MD","25":"MA","26":"MI","27":"MN","28":"MS","29":"MO","30":"MT",
        "31":"NE","32":"NV","33":"NH","34":"NJ","35":"NM","36":"NY","37":"NC","38":"ND","39":"OH",
        "40":"OK","41":"OR","42":"PA","44":"RI","45":"SC","46":"SD","47":"TN","48":"TX","49":"UT",
        "50":"VT","51":"VA","53":"WA","54":"WV","55":"WI","56":"WY","66":"GU","72":"PR","78":"VI"}

KEEP = ["TOT_EMP","EMP_PRSE","JOBS_1000","LOC_QUOTIENT","H_MEAN","A_MEAN","MEAN_PRSE",
        "A_PCT10","A_PCT25","A_MEDIAN","A_PCT75","A_PCT90","ANNUAL","HOURLY"]

def norm(h):
    return re.sub(r"[^A-Z0-9]+", "_", str(h).strip().upper())

rows = []
for f in sorted(glob.glob(os.path.join(SRC, "oesm*ma", "*_dl.xlsx"))):
    base = os.path.basename(f)
    if base.startswith("aMSA") or base.startswith("~$"):
        continue
    year = int(re.search(r"M(\d{4})_dl", base).group(1))
    ftype = "metro" if base.startswith("MSA") else "nonmetro"
    wb = openpyxl.load_workbook(f, read_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    hdr = [norm(h) for h in next(it)]
    hdr = ["AREA_TITLE" if h == "AREA_NAME" else h for h in hdr]
    idx = {h: i for i, h in enumerate(hdr)}
    oc = idx["OCC_CODE"]
    n = 0
    for r in it:
        if r[oc] != SOC:
            continue
        rec = {h: r[i] for h, i in idx.items()}
        # cross-industry, all-ownership only (2019+ files carry these fields)
        if "NAICS" in rec and str(rec["NAICS"]) != "000000":
            continue
        area = str(rec["AREA"]).strip()
        st = rec.get("PRIM_STATE")
        if not st:
            if ftype == "nonmetro":
                st = FIPS.get(area.zfill(7)[:2])
            else:
                m = re.search(r",\s*([A-Z]{2})", str(rec["AREA_TITLE"]))
                st = m.group(1) if m else None
        out = {"year": year, "area_type": ftype, "area_code": area,
               "area_title": rec["AREA_TITLE"], "prim_state": st,
               "occ_code": rec["OCC_CODE"], "occ_title": rec["OCC_TITLE"]}
        for k in KEEP:
            out[k.lower()] = rec.get(k)
        rows.append(out); n += 1
    wb.close()
    print(f"{base}: {n} rows")

df = pd.DataFrame(rows)

# BLS suppression codes: '*' / '**' = not released, '#' = wage >= top-code ($100/hr or $208k/yr)
num_cols = [k.lower() for k in KEEP if k not in ("ANNUAL", "HOURLY")]
df["a_mean_flag"] = df["a_mean"].where(df["a_mean"].astype(str).str.strip().isin(["*", "**", "#", "~"]))
for c in num_cols:
    df[c] = pd.to_numeric(df[c].replace({"*": None, "**": None, "#": None, "~": None}), errors="coerce")

df = df.sort_values(["area_type", "prim_state", "area_title", "year"]).reset_index(drop=True)

# Wide: one row per area, a_mean by year
wide = (df.pivot_table(index=["area_type", "area_code"], columns="year", values="a_mean", aggfunc="first")
          .reset_index())
latest = df.sort_values("year").groupby(["area_type", "area_code"])[["area_title", "prim_state"]].last().reset_index()
wide = latest.merge(wide, on=["area_type", "area_code"])
wide.columns = [f"a_mean_{c}" if isinstance(c, int) else c for c in wide.columns]
wide = wide.sort_values(["area_type", "prim_state", "area_title"])

summary = (df.groupby(["year", "area_type"])
             .agg(n_areas=("area_code", "nunique"), n_with_a_mean=("a_mean", "count"),
                  mean_of_a_mean=("a_mean", "mean"), median_of_a_mean=("a_mean", "median"),
                  min_a_mean=("a_mean", "min"), max_a_mean=("a_mean", "max"))
             .round(0).reset_index())

os.makedirs(OUT, exist_ok=True)
df.to_csv(os.path.join(OUT, "oews_childcare_39-9011_2015_2022_long.csv"), index=False)
with pd.ExcelWriter(os.path.join(OUT, "oews_childcare_39-9011_2015_2022.xlsx")) as xw:
    df.to_excel(xw, "panel_long", index=False)
    wide.to_excel(xw, "a_mean_wide", index=False)
    summary.to_excel(xw, "summary_by_year", index=False)
print(summary.to_string())
