# Direct Link
https://telecom-revenue-rca-dashboard-fcm8mjgkgcvs7khmwuv9za.streamlit.app/

# Telecom Revenue RCA Dashboard

A Python + Streamlit dashboard for **root cause analysis (RCA)** of telecom revenue KPIs.  
The app reads MTD vs LMTD Excel reports, calculates which customer segments drove revenue up or down, and presents the results as **tables, charts, and narrative insights**.  
It supports:

- **Single‑file RCA** – deep dive for one brand/period  
- **Two‑file comparison** – compare drivers between two brands or two time periods  

---

## Features

- Parse multi‑section KPI Excel files (Cluster, Handset Type, ARPU Segment, Usage Category, AoN Bucket, GB Slab, etc.).
- Compute per segment:
  - `Pre`, `Post`, `Absolute Change`, `% Change`
  - Contribution to total revenue change (%)
  - Contribution to Post period (%)
  - Combined impact score and RCA priority ranking
- Generate:
  - RCA results Excel with all segments and metrics
  - Bar charts of top positive and negative revenue drivers
  - Text insights summarizing key positive and negative drivers by KPI family
- Compare two RCA outputs:
  - Side‑by‑side metrics for File A vs File B
  - Deltas in contribution and absolute change
  - Higher brand contribution highlighted in **green**
  - Short textual comparison of top segments


---

## Local Setup and Run (Windows / PowerShell)

1. **Clone the repo**

2. **Create and activate a virtual environment**
  python -m venv .venv
  .venv\Scripts\Activate

3. **Install dependencies**
  pip install -r requirements.txt

4. **Run the Streamlit app**
  streamlit run app.py

5. Open the browser at `http://localhost:8501` if it does not open automatically.

---

## Data Requirements

Your KPI Excel must follow the same structure as `sample.xlsx`:

- A **brand‑level total** row at the top (with `Pre`, `Post`, `Absolute Change`, `% Change`).
- Multiple **sections separated by blank rows**, for example:
- `Clustername`
- `Handset Type`
- `Aon Bucket`
- `Arpu Segment`
- `Usage Category`
- `Gb Slab`
- `Mou Slab`
- `Multisimmer`
- `Base Type`
- `Vc User Category`
- Each section:
- Has a header row: first column is segment name, followed by `Pre`, `Post`, `Absolute Change`, `% Change`.
- Contains one row per segment.

The RCA engine assumes segment‑level `Absolute Change` values within a section sum (or approximately sum) to the brand‑level total change used as denominator for contribution.

---

## Using the Single File RCA Tab

1. In the sidebar, upload a KPI Excel file that matches the sample layout.
2. Enter:
- Sheet index (e.g. `0`) or sheet name (e.g. `Robi`).
- Brand name for the narrative (e.g. `Robi`).
3. Choose the number of **Top drivers in charts** (e.g. 10).
4. Click **“Run RCA (single file)”**.
5. Review:
- **RCA Narrative**: biggest positive and negative segments by KPI family.
- **RCA Results Preview**: table with contributions and RCA priority.
- **Charts**: top positive and negative revenue drivers.
6. Download:
- `rca_results_single.xlsx` – full RCA table.
- `rca_insights_single.txt` – key‑factor text summary.

---

## Using the Compare Two Files Tab

1. Upload **Excel A** and **Excel B** (e.g. Robi and Airtel, or two different months).
2. Provide:
- Sheet index or name for each file.
- Two brand names (e.g. `Robi`, `Airtel`) for labelling contributions.
3. Select **Top segments by change in contribution** (e.g. 10–20).
4. Click **“Run RCA Comparison”**.
5. Inspect:
- **Top segments where contribution changed most**:
  - KPI family and segment name.
  - Each brand’s contribution % and absolute change.
  - Deltas and a `Where` flag (Both / Only Brand A / Only Brand B).
  - Higher brand contribution per row highlighted in **green**.
- **Key comparison insights** text, e.g.  
  `Handset Type: Smartphone – Robi Contrib=X%, Airtel Contrib=Y%, ΔAbsChange=..., ΔContribution=... pts`.
6. Download `rca_comparison.xlsx` for deeper offline or BI analysis.

---

## Common Issues

- **`openpyxl` missing or Excel read error**

Install in the active environment:
pip install openpyxl




