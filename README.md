# 🌍 GHG Emissions Dashboard

A professional Streamlit dashboard for corporate carbon footprint reporting, built on the **GHG Protocol Framework**.

## Features

- 📤 **Upload your Excel file** (Lab1_dashboard.xlsx format) or enter data manually
- 📈 **5-Year Historical Trend** — Scope 1, 2, 3 absolute emissions & YoY change
- 🎯 **Reduction Trajectory** — Actual vs. BAU, Net Zero, SBTi pathways
- 📊 **Carbon Intensity** — tCO₂e per $M Revenue & per Employee over time
- 🥧 **Scope Breakdown** — Pie chart + stacked area chart
- 📥 **CSV Export** — Download a summary report

---

## 🚀 Deploy to Streamlit Community Cloud (Free)

### Step 1 — Push to GitHub
```bash
git init
git add .
git commit -m "Initial GHG dashboard"
git remote add origin https://github.com/YOUR_USERNAME/ghg-dashboard.git
git push -u origin main
```

### Step 2 — Deploy on Streamlit Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with GitHub
3. Click **"New app"**
4. Select your repository → Branch: `main` → Main file: `app.py`
5. Click **Deploy** ✅

Your app will be live at:
`https://YOUR_USERNAME-ghg-dashboard-app-XXXX.streamlit.app`

---

## 🖥️ Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## 📁 File Structure

```
ghg-dashboard/
├── app.py                  # Main Streamlit application
├── requirements.txt        # Python dependencies
├── .streamlit/
│   └── config.toml        # Dark theme configuration
└── README.md
```

---

## 📊 Excel Format Support

The dashboard auto-parses the **Lab1_dashboard.xlsx** format, reading from:
- `Assumptions` sheet — Company name, revenue, employees, reporting year
- `Summary & Intensity` — Current year scope totals & intensity ratios
- `Reduction Trajectory` — 5-year historical data (2020–2024) + projections
- `Trend Analysis` — Reduction targets (Net Zero, SBTi)

---

## 📋 GHG Protocol Compliance

Covers all three scopes per GHG Protocol Corporate Standard:
- **Scope 1** — Direct emissions (stationary combustion, fleet, fugitive)
- **Scope 2** — Purchased energy (market-based AND location-based dual reporting)
- **Scope 3** — Value chain (15 categories, upstream + downstream)

Emission factors: EPA CCCL 2023 · DEFRA 2023 · IEA 2023
