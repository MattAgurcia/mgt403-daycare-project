# MGT-403 Data Project — Where Should We Open Our Next Daycare?

Yale SOM · MGT 403 Probability Modeling and Statistics · Fall 2026

We build a county-level **price / cost / demand** dataset from three public sources and use it to recommend **10 U.S. metro counties** where a center-based preschool can charge prices that are high relative to its costs (labor + rent).

> **Teammates:** all results are in [`output/`](output/). Start with [`output/final_analysis_writeup.md`](output/final_analysis_writeup.md) for the recommendation and [`output/pre_analysis_writeup.md`](output/pre_analysis_writeup.md) for the Pre-Analysis questions 1–4. Clean data are in [`data/processed/`](data/processed/).

---

## What the assignment asks for

| Part | Deliverable | Status |
|---|---|---|
| Problem Set 4 | Distribution of childcare prices (NDCP) | separate PS submission |
| Problem Set 5 | ACS + OEWS county panel, wage tests, bivariate regression | data built here (`ps5_county_panel_2015_2022.csv`) |
| **Pre-Analysis: Prices and Costs** | Q1 analysis dataset · Q2 2022 summary table · Q3 price-vs-wage/rent scatters · Q4 regressions | ✅ `output/pre_analysis_writeup.md` |
| **Final project** | ≤ 6 slides / 10 min: (a) top-10 counties + visual, (b) methodology & reliability, (c) market characteristics, (d) stability of price premiums over time, (e) missing factors | analysis ✅ `output/final_analysis_writeup.md` · **slides: to do** |
| **Appendix** | Document detailing important data-handling decisions | draft ✅ "Data handling" sections of both write-ups · **final doc: to do** |
| **Code / prompts** | GitHub repo (or zip) with code, **or** prompt history + outputs | ✅ this repo + `Project_AI_Prompt_Log.docx` |

Grading: data construction · analysis & method · recommendation & business reasoning · communication.

## Headline result (2022)

Model, estimated separately for each year 2015–2022 across metro counties: weekly preschool price (2022 $) ~ childcare-worker wage + median gross rent + female college share (2022 R² = 0.657). Residual = actual − predicted price; the largest residuals are the recommended markets.

| Rank | County | Metro area | Actual $/wk | Predicted $/wk | Residual |
|---:|---|---|---:|---:|---:|
| 1 | Arlington County, Virginia | Washington-Arlington-Alexandria, DC-VA-MD-WV | $496.30 | $329.28 | +$167.02 |
| 2 | Stearns County, Minnesota | St. Cloud, MN | $336.67 | $173.03 | +$163.64 |
| 3 | San Francisco County, California | San Francisco-Oakland-Hayward, CA | $477.04 | $337.90 | +$139.14 |
| 4 | Winnebago County, Wisconsin | Oshkosh-Neenah, WI | $258.75 | $141.72 | +$117.03 |
| 5 | Alexandria city, Virginia | Washington-Arlington-Alexandria, DC-VA-MD-WV | $424.29 | $307.66 | +$116.63 |
| 6 | Belmont County, Ohio | Wheeling, WV-OH | $232.00 | $120.98 | +$111.02 |
| 7 | Jefferson County, Ohio | Weirton-Steubenville, WV-OH | $232.00 | $127.43 | +$104.57 |
| 8 | La Crosse County, Wisconsin | La Crosse-Onalaska, WI-MN | $258.75 | $155.94 | +$102.81 |
| 9 | Norfolk County, Massachusetts | Boston-Cambridge-Nashua, MA-NH | $420.00 | $317.21 | +$102.79 |
| 10 | Cumberland County, New Jersey | Vineland-Bridgeton, NJ | $276.77 | $178.14 | +$98.63 |

Premiums are persistent: 2019 vs 2022 residual correlation r = 0.78 (983 counties); 69% of 2019 top-quintile counties remain top quintile in 2022. See the write-up for caveats (regional NDCP prices, straight-line gap fills).

## Repository layout

```
.
├── README.md
├── Project_AI_Prompt_Log.docx          # AI prompt history (required if submitting prompts)
├── code/
│   ├── pre_analysis.py                 # NDCP + PS5 merge → analysis dataset, Q2–Q4 tables/figures, write-up
│   └── final_analysis.py               # yearly models, residual ranking, top 10, stability, write-up
├── build_childcare_wages.py            # PS5: OEWS 39-9011 wages 2015–2022 from oesmYYma.zip
├── build_acs_county.py                 # PS5: ACS 5-yr county variables via Census API
├── build_ps5_panel.py                  # PS5: crosswalk + OEWS + ACS → year × county metro panel
├── county_oews_crosswalk.csv           # course-provided county → OEWS area crosswalk
├── oews_childcare_39-9011_2015_2022*   # OEWS extract (csv/xlsx)
├── acs_county_2015_2022.*              # ACS extract (csv/xlsx)
├── ps5_county_panel_2015_2022.*        # PS5 merged panel (csv/xlsx)
├── data/processed/
│   ├── analysis_dataset.csv            # observed analysis sample, all years (FIPS stored as text)
│   └── panel_all_county_years.csv      # full metro panel before missing-data drops (for gap fill)
└── output/
    ├── pre_analysis_writeup.md         # Pre-Analysis answers 1–4 + data-handling decisions
    ├── final_analysis_writeup.md       # Final project analysis (a)–(e) + appendix notes
    ├── tables/                         # summary, regressions, predictor screen, top 10, stability …
    └── figures/                        # scatters, top-10 bar chart, 2019-vs-2022, top-10 over time
```

## Reproducing the results

Requirements: Python 3.9+, `pandas`, `numpy`, `matplotlib`, `openpyxl` (statsmodels optional — only used as a cross-check).

```bash
# 0) raw data (not in the repo — too large; see below)
#    put NDCP2022.xlsx and oesm15ma.zip … oesm22ma.zip in the repo root

# 1) PS5 panel (already included as ps5_county_panel_2015_2022.csv; rebuild only if needed)
mkdir -p /tmp/oesm && for y in 15 16 17 18 19 20 21 22; do unzip -oq oesm${y}ma.zip -d /tmp/oesm; done
python3 build_childcare_wages.py /tmp/oesm .
export CENSUS_API_KEY=your_key          # free: https://api.census.gov/data/key_signup.html
python3 build_acs_county.py
python3 build_ps5_panel.py

# 2) Pre-analysis (needs NDCP2022.xlsx) → data/processed/, output/tables, output/figures, write-up
python3 code/pre_analysis.py

# 3) Final analysis (uses data/processed/) → output/final_* and write-up
python3 code/final_analysis.py
```

Every number in the two write-ups is generated by these scripts.

## Raw data (download links)

| Source | File | Link |
|---|---|---|
| U.S. DOL Women's Bureau — National Database of Childcare Prices | `NDCP2022.xlsx` (county workbook, "Data and Documentation") | https://www.dol.gov/agencies/wb/topics/featured-childcare |
| NDCP Technical Report (Sept 2024) | PDF | linked from the page above |
| BLS OEWS metro/nonmetro area files | `oesm15ma.zip` … `oesm22ma.zip` | `https://www.bls.gov/oes/special-requests/oesmYYma.zip` |
| U.S. Census ACS 5-year (API) | pulled by `build_acs_county.py` | https://api.census.gov/data.html |
| CPI-U, CUUR0000SA0 annual averages | hardcoded in `code/pre_analysis.py` | https://www.bls.gov/cpi/tables/supplemental-files/historical-cpi-u-202212.pdf |

## Key data-handling decisions (short version)

- **Price:** NDCP `MCPRESCHOOL` — median weekly full-time price, center-based care, preschool (36–54 months).
- **Sample:** counties in OEWS **metro** areas (code starts `00`); Puerto Rico dropped. FIPS = 5-char text, OEWS area = 7-char text.
- **Merge:** NDCP ↔ PS5 panel on `[year, county_fips]`; NDCP year *t* ↔ ACS 5-year vintage ending in *t*.
- **Dollars:** constant 2022 $ via CPI-U (ACS dollars deflated by vintage year).
- **Final analysis gap fill:** interior gaps ≤ 2 years (e.g., VA/CA 2021 NDCP prices, suppressed OEWS wages) filled by straight-line interpolation — the method NDCP itself uses between survey years; 2022 never filled; all fills flagged.
- **Standard errors:** HC1 (heteroskedasticity-robust).

Full details: "Data handling" sections of both write-ups.

## AI use

Analysis code and write-ups were developed with Claude (Anthropic) via Cowork; prompts are logged in `Project_AI_Prompt_Log.docx`.
