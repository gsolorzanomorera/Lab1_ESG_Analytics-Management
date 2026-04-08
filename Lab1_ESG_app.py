import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import io

# ─── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GHG Emissions Dashboard",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #0f1117; color: #e8eaf6; }
  [data-testid="stSidebar"] { background: #1a1d2e; border-right: 1px solid #2d3561; }
  .metric-card {
    background: linear-gradient(135deg, #1e2140 0%, #252a4a 100%);
    border: 1px solid #3d4a8a;
    border-radius: 12px;
    padding: 20px;
    text-align: center;
    margin-bottom: 12px;
  }
  .metric-value { font-size: 2rem; font-weight: 700; }
  .metric-label { font-size: 0.8rem; color: #9ea7d8; letter-spacing: 0.08em; text-transform: uppercase; margin-top: 4px; }
  .metric-delta { font-size: 0.85rem; margin-top: 6px; }
  .section-header {
    background: linear-gradient(90deg, #2d3561, transparent);
    border-left: 3px solid #6c72cb;
    padding: 10px 16px;
    border-radius: 0 8px 8px 0;
    margin: 24px 0 16px 0;
    font-weight: 600;
    font-size: 1.05rem;
    letter-spacing: 0.04em;
  }
  .stTabs [data-baseweb="tab"] { background: #1a1d2e; border-radius: 8px 8px 0 0; }
  .stTabs [aria-selected="true"] { background: #2d3561; }
  .report-box {
    background: #1e2140;
    border: 1px solid #3d4a8a;
    border-radius: 12px;
    padding: 24px;
    margin: 12px 0;
  }
</style>
""", unsafe_allow_html=True)

# ─── Helpers ─────────────────────────────────────────────────────────────────
SCOPE_COLORS = {
    "Scope 1": "#f97316",
    "Scope 2 (Market)": "#3b82f6",
    "Scope 2 (Location)": "#60a5fa",
    "Scope 3": "#8b5cf6",
    "Total": "#10b981",
    "BAU": "#ef4444",
    "Net Zero": "#10b981",
    "SBTi S1+S2": "#3b82f6",
    "SBTi S3": "#8b5cf6",
}

def fmt(val, unit="tCO₂e"):
    if pd.isna(val) or val == 0:
        return "—"
    if abs(val) >= 1_000_000:
        return f"{val/1_000_000:.2f}M {unit}"
    if abs(val) >= 1_000:
        return f"{val/1_000:.1f}K {unit}"
    return f"{val:.0f} {unit}"

def pct_change(new, old):
    if not old or old == 0:
        return None
    return (new - old) / abs(old) * 100

def delta_badge(pct):
    if pct is None:
        return ""
    arrow = "▲" if pct > 0 else "▼"
    color = "#ef4444" if pct > 0 else "#10b981"
    return f'<span style="color:{color}">{arrow} {abs(pct):.1f}% vs prior year</span>'

# ─── Excel parser ────────────────────────────────────────────────────────────
def parse_excel(file):
    """Extract structured data from the Lab1_dashboard.xlsx format."""
    wb = pd.read_excel(file, sheet_name=None, header=None)
    data = {}

    # --- Assumptions sheet ---
    try:
        df = wb.get("Assumptions", pd.DataFrame())
        raw = df.astype(str)
        for _, row in raw.iterrows():
            row_str = " ".join(row.fillna("").astype(str))
            if "Company Name" in row_str:
                vals = row.dropna()
                data["company"] = str(vals.iloc[1]) if len(vals) > 1 else "Unknown"
            if "Industry" in row_str and "Sector" in row_str:
                vals = row.dropna()
                data["industry"] = str(vals.iloc[1]) if len(vals) > 1 else ""
            if "Annual Revenue" in row_str:
                vals = pd.to_numeric(row, errors="coerce").dropna()
                if len(vals): data["revenue"] = float(vals.iloc[0])
            if "Number of Employees" in row_str:
                vals = pd.to_numeric(row, errors="coerce").dropna()
                if len(vals): data["employees"] = float(vals.iloc[0])
            if "Reporting Year (Current)" in row_str:
                vals = pd.to_numeric(row, errors="coerce").dropna()
                if len(vals): data["current_year"] = int(vals.iloc[0])
    except Exception:
        pass

    # --- Summary sheet ---
    try:
        df = wb.get("Summary & Intensity", pd.DataFrame())
        raw = df.astype(str)
        for _, row in raw.iterrows():
            row_str = " ".join(row.fillna("").astype(str))
            nums = pd.to_numeric(row, errors="coerce").dropna()
            if "Scope 1" in row_str and "Direct" in row_str and len(nums) >= 1:
                data["s1_current"] = float(nums.iloc[0])
            if "Scope 2" in row_str and "Market" in row_str and "PRIMARY" in row_str and len(nums) >= 1:
                data["s2_mb_current"] = float(nums.iloc[0])
            if "Scope 2" in row_str and "Location" in row_str and "SECONDARY" in row_str and len(nums) >= 1:
                data["s2_lb_current"] = float(nums.iloc[0])
            if "Scope 3" in row_str and "Value Chain" in row_str and len(nums) >= 1:
                data["s3_current"] = float(nums.iloc[0])
            if "Carbon Intensity" in row_str and "Revenue" in row_str and "Scope 1" not in row_str and len(nums) >= 1:
                data["intensity_rev"] = float(nums.iloc[0])
            if "Carbon Intensity" in row_str and "Employee" in row_str and len(nums) >= 1:
                data["intensity_emp"] = float(nums.iloc[0])
    except Exception:
        pass

    # --- Reduction Trajectory (historical 5yr) ---
    try:
        df = wb.get("Reduction Trajectory", pd.DataFrame())
        raw = df.astype(str)
        years_row, s1_row, s2lb_row, s2mb_row, s3_row, total_row = None, None, None, None, None, None
        for i, row in raw.iterrows():
            row_str = " ".join(row.fillna("").astype(str))
            if "Scope / Year" in row_str or ("2020" in row_str and "2021" in row_str):
                years_row = i
            if "Scope 1" in row_str and "Direct" in row_str:
                s1_row = i
            if "location based" in row_str.lower() and "Scope 2" in row_str:
                s2lb_row = i
            if "market based" in row_str.lower() and "Scope 2" in row_str:
                s2mb_row = i
            if "Scope 3" in row_str and "Value Chain" in row_str:
                s3_row = i
            if "TOTAL" in row_str and "S1+S2+S3" in row_str:
                total_row = i

        if years_row is not None:
            year_vals = pd.to_numeric(df.iloc[years_row], errors="coerce").dropna()
            years = [int(y) for y in year_vals if 2015 <= y <= 2035]

            def extract_row(row_idx):
                if row_idx is None:
                    return {}
                vals = pd.to_numeric(df.iloc[row_idx], errors="coerce")
                result = {}
                for col_idx, val in enumerate(vals):
                    col_year_val = pd.to_numeric(df.iloc[years_row, col_idx], errors="coerce")
                    if not pd.isna(col_year_val) and not pd.isna(val):
                        y = int(col_year_val)
                        if 2015 <= y <= 2035:
                            result[y] = float(val)
                return result

            data["hist_s1"] = extract_row(s1_row)
            data["hist_s2lb"] = extract_row(s2lb_row)
            data["hist_s2mb"] = extract_row(s2mb_row)
            data["hist_s3"] = extract_row(s3_row)
            data["hist_total"] = extract_row(total_row)
            data["hist_years"] = years
    except Exception:
        pass

    # --- Trend Analysis ---
    try:
        df = wb.get("Trend Analysis", pd.DataFrame())
        raw = df.astype(str)
        traj = {}
        for i, row in raw.iterrows():
            row_str = " ".join(row.fillna("").astype(str))
            if "Net Zero" in row_str:
                nums = pd.to_numeric(row, errors="coerce").dropna()
                if len(nums) >= 3:
                    traj["netzero"] = {"year": int(nums.iloc[0]), "baseline_year": int(nums.iloc[1]),
                                       "baseline": float(nums.iloc[2]), "pct": float(nums.iloc[3]),
                                       "target": float(nums.iloc[4])}
            if "Near-Term" in row_str and "SBTi" in row_str:
                nums = pd.to_numeric(row, errors="coerce").dropna()
                if len(nums) >= 3:
                    traj["sbti_s12"] = {"year": int(nums.iloc[0]), "target": float(nums.iloc[-2])}
            if "Scope 3 Reduction" in row_str:
                nums = pd.to_numeric(row, errors="coerce").dropna()
                if len(nums) >= 3:
                    traj["sbti_s3"] = {"year": int(nums.iloc[0]), "target": float(nums.iloc[-2])}
        data["targets"] = traj
    except Exception:
        pass

    # --- Trajectory scenarios ---
    try:
        df = wb.get("Reduction Trajectory", pd.DataFrame())
        raw = df.astype(str)
        bau_row = nz_row = sbti_row = s3t_row = traj_years_row = None
        for i, row in raw.iterrows():
            row_str = " ".join(row.fillna("").astype(str))
            if "Business as usual" in row_str:
                bau_row = i
            if "Net Zero Commitment" in row_str and bau_row is not None:
                nz_row = i
            if "Near-Term Science-Based" in row_str:
                sbti_row = i
            if "Scope 3 Reduction Target" in row_str:
                s3t_row = i
            if "Category / Year" in row_str:
                traj_years_row = i

        if traj_years_row is not None:
            def extract_traj(row_idx):
                if row_idx is None:
                    return {}
                vals = pd.to_numeric(df.iloc[row_idx], errors="coerce")
                yr_vals = pd.to_numeric(df.iloc[traj_years_row], errors="coerce")
                result = {}
                for c in range(len(vals)):
                    y = yr_vals.iloc[c]
                    v = vals.iloc[c]
                    if not pd.isna(y) and not pd.isna(v) and 2020 <= y <= 2035:
                        result[int(y)] = float(v)
                return result

            data["traj_bau"] = extract_traj(bau_row)
            data["traj_nz"] = extract_traj(nz_row)
            data["traj_sbti"] = extract_traj(sbti_row)
            data["traj_s3"] = extract_traj(s3t_row)
    except Exception:
        pass

    return data

# ─── Default / demo data ─────────────────────────────────────────────────────
DEFAULT = {
    "company": "Your Company",
    "industry": "Technology",
    "revenue": 245100,
    "employees": 228000,
    "current_year": 2024,
    "s1_current": 143510,
    "s2_mb_current": 259090,
    "s2_lb_current": 9955368,
    "s3_current": 15140000,
    "intensity_rev": 63.41,
    "intensity_emp": 68.17,
    "hist_years": [2020, 2021, 2022, 2023, 2024],
    "hist_s1": {2020: 118100, 2021: 123704, 2022: 139413, 2023: 144960, 2024: 143510},
    "hist_s2mb": {2020: 456119, 2021: 429405, 2022: 288029, 2023: 393134, 2024: 259090},
    "hist_s2lb": {2020: 4328916, 2021: 5010667, 2022: 6381250, 2023: 8077403, 2024: 9955368},
    "hist_s3": {2020: 11796000, 2021: 13576000, 2022: 15916000, 2023: 16397000, 2024: 15140000},
    "hist_total": {2020: 12370000, 2021: 14129000, 2022: 16343000, 2023: 16935000, 2024: 15543000},
    "traj_bau": {2024: 12370000, 2025: 16456070, 2026: 17422778, 2027: 18446275, 2028: 19529898, 2029: 20677178, 2030: 21891855},
    "traj_nz": {2024: 12370000, 2025: 11906125, 2026: 11442250, 2027: 10978375, 2028: 10514500, 2029: 10050625, 2030: 9586750},
    "traj_sbti": {2024: 15543000, 2025: 12952500, 2026: 10362000, 2027: 7771500, 2028: 5181000, 2029: 2590500, 2030: 0},
    "traj_s3": {2024: 11796000, 2025: 11206200, 2026: 10616400, 2027: 10026600, 2028: 9436800, 2029: 8847000, 2030: 8257200},
    "targets": {
        "netzero": {"year": 2030, "baseline_year": 2022, "baseline": 12370000, "pct": 0.3, "target": 8659000},
        "sbti_s12": {"year": 2030, "target": 0},
        "sbti_s3": {"year": 2030, "target": 5898000},
    }
}

# ─── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🌍 GHG Dashboard")
    st.markdown("*GHG Protocol Framework*")
    st.divider()

    mode = st.radio("Data Source", ["📤 Upload Excel", "✏️ Manual Entry"], horizontal=False)

    d = dict(DEFAULT)  # working data dict

    if mode == "📤 Upload Excel":
        uploaded = st.file_uploader("Upload your Lab1_dashboard.xlsx", type=["xlsx"])
        if uploaded:
            with st.spinner("Parsing Excel…"):
                parsed = parse_excel(uploaded)
                d.update({k: v for k, v in parsed.items() if v})
            st.success("✅ File loaded!")
        else:
            st.info("Upload your file or use demo data below.")
            st.markdown("**Demo data loaded** (Microsoft 2020–2024)")

    else:  # Manual entry
        st.markdown("### 🏢 Company Info")
        d["company"] = st.text_input("Company Name", d["company"])
        d["industry"] = st.text_input("Industry", d["industry"])
        d["current_year"] = st.number_input("Reporting Year", value=d["current_year"], step=1)
        d["revenue"] = st.number_input("Revenue ($M)", value=float(d["revenue"]), step=1000.0)
        d["employees"] = st.number_input("Employees (FTE)", value=float(d["employees"]), step=1000.0)

        st.markdown("### 📊 Current Year Emissions (tCO₂e)")
        d["s1_current"] = st.number_input("Scope 1", value=float(d["s1_current"]), step=1000.0)
        d["s2_mb_current"] = st.number_input("Scope 2 (Market-Based)", value=float(d["s2_mb_current"]), step=1000.0)
        d["s2_lb_current"] = st.number_input("Scope 2 (Location-Based)", value=float(d["s2_lb_current"]), step=100000.0)
        d["s3_current"] = st.number_input("Scope 3", value=float(d["s3_current"]), step=100000.0)

        st.markdown("### 📅 Historical Data (tCO₂e)")
        st.caption("Enter Scope 1+2+3 total per year")
        for yr in [2020, 2021, 2022, 2023]:
            d["hist_total"][yr] = st.number_input(str(yr), value=float(d["hist_total"].get(yr, 0)), step=100000.0, key=f"ht_{yr}")

    st.divider()
    st.markdown("### 🎯 Reduction Targets")
    nz_year = st.number_input("Net Zero Target Year", value=2030, step=1)
    nz_target = st.number_input("Net Zero Target (tCO₂e)", value=float(d["targets"]["netzero"]["target"]), step=100000.0)

# ─── Compute derived values ──────────────────────────────────────────────────
total_current = d["s1_current"] + d["s2_mb_current"] + d["s3_current"]
cy = d["current_year"]
py = cy - 1

s1_py = d["hist_s1"].get(py, None)
s2_py = d["hist_s2mb"].get(py, None)
s3_py = d["hist_s3"].get(py, None)
total_py = d["hist_total"].get(py, None)

d1_s1 = pct_change(d["s1_current"], s1_py)
d1_s2 = pct_change(d["s2_mb_current"], s2_py)
d1_s3 = pct_change(d["s3_current"], s3_py)
d1_tot = pct_change(total_current, total_py)

int_rev = total_current / d["revenue"] if d["revenue"] else 0
int_emp = total_current / d["employees"] if d["employees"] else 0

s1_pct = d["s1_current"] / total_current * 100 if total_current else 0
s2_pct = d["s2_mb_current"] / total_current * 100 if total_current else 0
s3_pct = d["s3_current"] / total_current * 100 if total_current else 0

# ─── Header ──────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="background:linear-gradient(135deg,#1e2140,#252a4a);border-radius:16px;padding:28px 36px;margin-bottom:24px;border:1px solid #3d4a8a;">
  <h1 style="margin:0;font-size:2rem;">🌍 {d['company']} — Carbon Footprint Dashboard</h1>
  <p style="margin:6px 0 0;color:#9ea7d8;font-size:0.95rem;">
    GHG Protocol Framework &nbsp;·&nbsp; Reporting Year: <b>{cy}</b> &nbsp;·&nbsp;
    Industry: <b>{d['industry']}</b> &nbsp;·&nbsp;
    Revenue: <b>${d['revenue']:,.0f}M</b> &nbsp;·&nbsp;
    Employees: <b>{int(d['employees']):,}</b>
  </p>
</div>
""", unsafe_allow_html=True)

# ─── KPI Cards ───────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)

def kpi(col, label, value, unit, delta_pct, color):
    badge = delta_badge(delta_pct)
    col.markdown(f"""
    <div class="metric-card" style="border-color:{color}40;">
      <div class="metric-value" style="color:{color};">{fmt(value)}</div>
      <div class="metric-label">{label}</div>
      <div class="metric-delta">{badge if badge else '&nbsp;'}</div>
    </div>
    """, unsafe_allow_html=True)

kpi(c1, f"Scope 1 · Direct ({cy})", d["s1_current"], "tCO₂e", d1_s1, "#f97316")
kpi(c2, f"Scope 2 Market-Based ({cy})", d["s2_mb_current"], "tCO₂e", d1_s2, "#3b82f6")
kpi(c3, f"Scope 3 · Value Chain ({cy})", d["s3_current"], "tCO₂e", d1_s3, "#8b5cf6")
kpi(c4, f"Total S1+S2+S3 ({cy})", total_current, "tCO₂e", d1_tot, "#10b981")

# ─── Intensity row ────────────────────────────────────────────────────────────
i1, i2, i3, i4 = st.columns(4)
i1.markdown(f"""<div class="metric-card" style="border-color:#6c72cb40;">
  <div class="metric-value" style="color:#6c72cb;font-size:1.5rem;">{int_rev:.1f}</div>
  <div class="metric-label">tCO₂e per $M Revenue</div>
</div>""", unsafe_allow_html=True)
i2.markdown(f"""<div class="metric-card" style="border-color:#6c72cb40;">
  <div class="metric-value" style="color:#6c72cb;font-size:1.5rem;">{int_emp:.1f}</div>
  <div class="metric-label">tCO₂e per Employee</div>
</div>""", unsafe_allow_html=True)
i3.markdown(f"""<div class="metric-card" style="border-color:#f9731640;">
  <div class="metric-value" style="color:#f97316;font-size:1.5rem;">{s1_pct:.1f}%</div>
  <div class="metric-label">Scope 1 Share</div>
</div>""", unsafe_allow_html=True)
i4.markdown(f"""<div class="metric-card" style="border-color:#8b5cf640;">
  <div class="metric-value" style="color:#8b5cf6;font-size:1.5rem;">{s3_pct:.1f}%</div>
  <div class="metric-label">Scope 3 Share</div>
</div>""", unsafe_allow_html=True)

# ─── Charts ──────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["📈 5-Year Trend", "🎯 Reduction Trajectory", "📊 Intensity", "🥧 Scope Breakdown"])

# ── Tab 1: 5-Year Trend ──────────────────────────────────────────────────────
with tab1:
    st.markdown('<div class="section-header">📈 Scope Emissions — 5-Year Historical Trend</div>', unsafe_allow_html=True)

    hist_years = sorted(d["hist_years"])
    s1_vals = [d["hist_s1"].get(y, None) for y in hist_years]
    s2_vals = [d["hist_s2mb"].get(y, None) for y in hist_years]
    s3_vals = [d["hist_s3"].get(y, None) for y in hist_years]
    tot_vals = [d["hist_total"].get(y, None) for y in hist_years]

    fig = make_subplots(rows=1, cols=2, subplot_titles=("Absolute Emissions (tCO₂e)", "Year-on-Year Change (%)"),
                        column_widths=[0.6, 0.4])

    fig.add_trace(go.Bar(name="Scope 1", x=hist_years, y=s1_vals,
                         marker_color=SCOPE_COLORS["Scope 1"], offsetgroup=0), row=1, col=1)
    fig.add_trace(go.Bar(name="Scope 2 (Market)", x=hist_years, y=s2_vals,
                         marker_color=SCOPE_COLORS["Scope 2 (Market)"], offsetgroup=1), row=1, col=1)
    fig.add_trace(go.Bar(name="Scope 3", x=hist_years, y=s3_vals,
                         marker_color=SCOPE_COLORS["Scope 3"], offsetgroup=2), row=1, col=1)
    fig.add_trace(go.Scatter(name="Total", x=hist_years, y=tot_vals,
                             mode="lines+markers", line=dict(color=SCOPE_COLORS["Total"], width=3),
                             marker=dict(size=8)), row=1, col=1)

    # YoY change
    yoy_s1 = [None] + [pct_change(s1_vals[i], s1_vals[i-1]) for i in range(1, len(s1_vals))]
    yoy_s3 = [None] + [pct_change(s3_vals[i], s3_vals[i-1]) for i in range(1, len(s3_vals))]
    yoy_tot = [None] + [pct_change(tot_vals[i], tot_vals[i-1]) for i in range(1, len(tot_vals))]

    fig.add_trace(go.Scatter(name="S1 YoY", x=hist_years, y=yoy_s1, mode="lines+markers",
                             line=dict(color=SCOPE_COLORS["Scope 1"], dash="dot"), marker=dict(size=6)), row=1, col=2)
    fig.add_trace(go.Scatter(name="S3 YoY", x=hist_years, y=yoy_s3, mode="lines+markers",
                             line=dict(color=SCOPE_COLORS["Scope 3"], dash="dot"), marker=dict(size=6)), row=1, col=2)
    fig.add_trace(go.Scatter(name="Total YoY", x=hist_years, y=yoy_tot, mode="lines+markers",
                             line=dict(color=SCOPE_COLORS["Total"], width=2), marker=dict(size=8)), row=1, col=2)
    fig.add_hline(y=0, line_dash="dash", line_color="#555", row=1, col=2)

    fig.update_layout(
        height=460, barmode="group", plot_bgcolor="#1a1d2e", paper_bgcolor="#0f1117",
        font=dict(color="#e8eaf6"), legend=dict(bgcolor="#1e2140", bordercolor="#3d4a8a"),
        margin=dict(t=50, b=40, l=60, r=40),
    )
    fig.update_xaxes(gridcolor="#2d3561", tickformat="d")
    fig.update_yaxes(gridcolor="#2d3561")
    st.plotly_chart(fig, use_container_width=True)

    # Data table
    with st.expander("📋 View Historical Data Table"):
        table_data = pd.DataFrame({
            "Year": hist_years,
            "Scope 1 (tCO₂e)": s1_vals,
            "Scope 2 MB (tCO₂e)": s2_vals,
            "Scope 3 (tCO₂e)": s3_vals,
            "Total (tCO₂e)": tot_vals,
        })
        st.dataframe(table_data.style.format({
            "Scope 1 (tCO₂e)": "{:,.0f}",
            "Scope 2 MB (tCO₂e)": "{:,.0f}",
            "Scope 3 (tCO₂e)": "{:,.0f}",
            "Total (tCO₂e)": "{:,.0f}",
        }), use_container_width=True)

# ── Tab 2: Reduction Trajectory ───────────────────────────────────────────────
with tab2:
    st.markdown('<div class="section-header">🎯 Emission Reduction Trajectory vs. Targets</div>', unsafe_allow_html=True)

    fig2 = go.Figure()

    # Historical actual
    fig2.add_trace(go.Scatter(
        name="Historical Actual", x=hist_years, y=tot_vals,
        mode="lines+markers", line=dict(color=SCOPE_COLORS["Total"], width=3),
        marker=dict(size=9, symbol="circle"), fill="tozeroy",
        fillcolor="rgba(16,185,129,0.06)"
    ))

    # Trajectories
    def add_traj(name, key, color, dash="dash"):
        td = d.get(key, {})
        if td:
            yrs = sorted(td.keys())
            fig2.add_trace(go.Scatter(
                name=name, x=yrs, y=[td[y] for y in yrs],
                mode="lines+markers", line=dict(color=color, width=2, dash=dash),
                marker=dict(size=6)
            ))

    add_traj("Business As Usual", "traj_bau", SCOPE_COLORS["BAU"], "dot")
    add_traj("Net Zero Pathway", "traj_nz", SCOPE_COLORS["Net Zero"], "dash")
    add_traj("SBTi Scope 1+2 Target", "traj_sbti", SCOPE_COLORS["SBTi S1+S2"], "dashdot")
    add_traj("SBTi Scope 3 Target", "traj_s3", SCOPE_COLORS["SBTi S3"], "longdash")

    # Net zero annotation
    fig2.add_annotation(
        x=nz_year, y=nz_target,
        text=f"🎯 Net Zero Target<br>{nz_year}: {fmt(nz_target)}",
        showarrow=True, arrowhead=2, arrowcolor="#10b981",
        font=dict(color="#10b981", size=11),
        bgcolor="#0f1117", bordercolor="#10b981", borderwidth=1,
        ax=60, ay=-60
    )

    fig2.update_layout(
        height=500, plot_bgcolor="#1a1d2e", paper_bgcolor="#0f1117",
        font=dict(color="#e8eaf6"), legend=dict(bgcolor="#1e2140", bordercolor="#3d4a8a"),
        xaxis_title="Year", yaxis_title="tCO₂e",
        margin=dict(t=30, b=50, l=80, r=40),
    )
    fig2.update_xaxes(gridcolor="#2d3561", tickformat="d")
    fig2.update_yaxes(gridcolor="#2d3561")
    st.plotly_chart(fig2, use_container_width=True)

    # Target summary table
    tgt = d.get("targets", {})
    if tgt:
        st.markdown('<div class="section-header">Target Summary</div>', unsafe_allow_html=True)
        rows = []
        if "netzero" in tgt:
            t = tgt["netzero"]
            rows.append({"Target": "Net Zero Commitment", "Year": t.get("year","—"),
                         "Baseline (tCO₂e)": f"{t.get('baseline',0):,.0f}",
                         "Reduction": f"{t.get('pct',0)*100:.0f}%",
                         "Target (tCO₂e)": f"{t.get('target',0):,.0f}"})
        if "sbti_s12" in tgt:
            rows.append({"Target": "SBTi Near-Term (S1+S2)", "Year": tgt["sbti_s12"].get("year","—"),
                         "Baseline (tCO₂e)": "—", "Reduction": "100%",
                         "Target (tCO₂e)": "0"})
        if "sbti_s3" in tgt:
            rows.append({"Target": "SBTi Scope 3 Target", "Year": tgt["sbti_s3"].get("year","—"),
                         "Baseline (tCO₂e)": "—", "Reduction": "50%",
                         "Target (tCO₂e)": f"{tgt['sbti_s3'].get('target',0):,.0f}"})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Tab 3: Intensity ─────────────────────────────────────────────────────────
with tab3:
    st.markdown('<div class="section-header">📊 Carbon Intensity Ratios</div>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    # Intensity over time (revenue)
    with col_a:
        int_rev_vals = []
        for y in hist_years:
            t = d["hist_total"].get(y)
            int_rev_vals.append(t / d["revenue"] if t and d["revenue"] else None)

        fig3a = go.Figure()
        fig3a.add_trace(go.Scatter(
            name="tCO₂e / $M Revenue", x=hist_years, y=int_rev_vals,
            mode="lines+markers+text",
            text=[f"{v:.1f}" if v else "" for v in int_rev_vals],
            textposition="top center",
            line=dict(color="#6c72cb", width=3), marker=dict(size=10),
            fill="tozeroy", fillcolor="rgba(108,114,203,0.1)"
        ))
        fig3a.update_layout(
            title="Carbon Intensity — Revenue (tCO₂e / $M)",
            height=360, plot_bgcolor="#1a1d2e", paper_bgcolor="#1e2140",
            font=dict(color="#e8eaf6"), margin=dict(t=50, b=40, l=60, r=20),
            showlegend=False
        )
        fig3a.update_xaxes(gridcolor="#2d3561", tickformat="d")
        fig3a.update_yaxes(gridcolor="#2d3561")
        st.plotly_chart(fig3a, use_container_width=True)

    # Intensity over time (employee)
    with col_b:
        int_emp_vals = []
        for y in hist_years:
            t = d["hist_total"].get(y)
            int_emp_vals.append(t / d["employees"] if t and d["employees"] else None)

        fig3b = go.Figure()
        fig3b.add_trace(go.Scatter(
            name="tCO₂e / Employee", x=hist_years, y=int_emp_vals,
            mode="lines+markers+text",
            text=[f"{v:.1f}" if v else "" for v in int_emp_vals],
            textposition="top center",
            line=dict(color="#f59e0b", width=3), marker=dict(size=10),
            fill="tozeroy", fillcolor="rgba(245,158,11,0.1)"
        ))
        fig3b.update_layout(
            title="Carbon Intensity — Employee (tCO₂e / FTE)",
            height=360, plot_bgcolor="#1a1d2e", paper_bgcolor="#1e2140",
            font=dict(color="#e8eaf6"), margin=dict(t=50, b=40, l=60, r=20),
            showlegend=False
        )
        fig3b.update_xaxes(gridcolor="#2d3561", tickformat="d")
        fig3b.update_yaxes(gridcolor="#2d3561")
        st.plotly_chart(fig3b, use_container_width=True)

    # Scope 2 location vs market
    st.markdown('<div class="section-header">Scope 2: Location-Based vs. Market-Based</div>', unsafe_allow_html=True)
    s2lb_vals = [d["hist_s2lb"].get(y) for y in hist_years]
    s2mb_vals = [d["hist_s2mb"].get(y) for y in hist_years]

    fig3c = go.Figure()
    fig3c.add_trace(go.Bar(name="Location-Based", x=hist_years, y=s2lb_vals,
                           marker_color=SCOPE_COLORS["Scope 2 (Location)"], opacity=0.7))
    fig3c.add_trace(go.Bar(name="Market-Based", x=hist_years, y=s2mb_vals,
                           marker_color=SCOPE_COLORS["Scope 2 (Market)"]))
    fig3c.update_layout(
        height=320, barmode="group", plot_bgcolor="#1a1d2e", paper_bgcolor="#0f1117",
        font=dict(color="#e8eaf6"), legend=dict(bgcolor="#1e2140"),
        xaxis_title="Year", yaxis_title="tCO₂e",
        margin=dict(t=20, b=50, l=70, r=30),
        annotations=[dict(text="Location-Based (grid average) vs. Market-Based (REC/PPA adjusted)",
                          x=0.5, y=1.04, xref="paper", yref="paper",
                          showarrow=False, font=dict(color="#9ea7d8", size=12))]
    )
    fig3c.update_xaxes(gridcolor="#2d3561", tickformat="d")
    fig3c.update_yaxes(gridcolor="#2d3561")
    st.plotly_chart(fig3c, use_container_width=True)

# ── Tab 4: Scope Breakdown ───────────────────────────────────────────────────
with tab4:
    st.markdown('<div class="section-header">🥧 Scope Breakdown — Current Year</div>', unsafe_allow_html=True)

    col_p, col_t = st.columns([1, 1])

    with col_p:
        fig4a = go.Figure(go.Pie(
            labels=["Scope 1 — Direct", "Scope 2 — Market-Based", "Scope 3 — Value Chain"],
            values=[d["s1_current"], d["s2_mb_current"], d["s3_current"]],
            hole=0.55,
            marker=dict(colors=[SCOPE_COLORS["Scope 1"], SCOPE_COLORS["Scope 2 (Market)"], SCOPE_COLORS["Scope 3"]],
                        line=dict(color="#0f1117", width=3)),
            textinfo="label+percent",
            textfont=dict(color="#e8eaf6", size=13),
        ))
        fig4a.add_annotation(text=f"<b>{fmt(total_current)}</b><br><span style='font-size:10px'>Total</span>",
                             x=0.5, y=0.5, font=dict(color="#e8eaf6", size=14), showarrow=False)
        fig4a.update_layout(
            title=f"Scope Mix — {cy}", height=420,
            paper_bgcolor="#1e2140", font=dict(color="#e8eaf6"),
            legend=dict(bgcolor="#1e2140", bordercolor="#3d4a8a"),
            margin=dict(t=50, b=20, l=20, r=20)
        )
        st.plotly_chart(fig4a, use_container_width=True)

    with col_t:
        # Stacked area over years
        fig4b = go.Figure()
        fig4b.add_trace(go.Scatter(
            name="Scope 3", x=hist_years, y=[d["hist_s3"].get(y, 0) for y in hist_years],
            stackgroup="one", fillcolor="rgba(139,92,246,0.7)", line=dict(color=SCOPE_COLORS["Scope 3"])
        ))
        fig4b.add_trace(go.Scatter(
            name="Scope 2 (Market)", x=hist_years, y=[d["hist_s2mb"].get(y, 0) for y in hist_years],
            stackgroup="one", fillcolor="rgba(59,130,246,0.7)", line=dict(color=SCOPE_COLORS["Scope 2 (Market)"])
        ))
        fig4b.add_trace(go.Scatter(
            name="Scope 1", x=hist_years, y=[d["hist_s1"].get(y, 0) for y in hist_years],
            stackgroup="one", fillcolor="rgba(249,115,22,0.7)", line=dict(color=SCOPE_COLORS["Scope 1"])
        ))
        fig4b.update_layout(
            title="Stacked Emissions Over Time",
            height=420, plot_bgcolor="#1a1d2e", paper_bgcolor="#1e2140",
            font=dict(color="#e8eaf6"), legend=dict(bgcolor="#1e2140"),
            xaxis_title="Year", yaxis_title="tCO₂e",
            margin=dict(t=50, b=50, l=70, r=20)
        )
        fig4b.update_xaxes(gridcolor="#2d3561", tickformat="d")
        fig4b.update_yaxes(gridcolor="#2d3561")
        st.plotly_chart(fig4b, use_container_width=True)

# ─── Export Report ────────────────────────────────────────────────────────────
st.divider()
st.markdown('<div class="section-header">📄 Export Summary Report</div>', unsafe_allow_html=True)

if st.button("📥 Generate CSV Summary Report"):
    rows = [
        ["Metric", "Value", "Unit"],
        ["Company", d["company"], ""],
        ["Reporting Year", cy, ""],
        ["Revenue", d["revenue"], "$M"],
        ["Employees", d["employees"], "FTE"],
        ["Scope 1", d["s1_current"], "tCO₂e"],
        ["Scope 2 (Market-Based)", d["s2_mb_current"], "tCO₂e"],
        ["Scope 2 (Location-Based)", d["s2_lb_current"], "tCO₂e"],
        ["Scope 3", d["s3_current"], "tCO₂e"],
        ["Total (S1+S2MB+S3)", total_current, "tCO₂e"],
        ["Carbon Intensity — Revenue", round(int_rev, 2), "tCO₂e/$M"],
        ["Carbon Intensity — Employee", round(int_emp, 2), "tCO₂e/FTE"],
        [],
        ["Historical Total Emissions"],
    ]
    for y in hist_years:
        rows.append([y, d["hist_total"].get(y, ""), "tCO₂e"])

    df_export = pd.DataFrame(rows)
    csv = df_export.to_csv(index=False, header=False)
    st.download_button("⬇️ Download CSV", csv, file_name=f"ghg_report_{d['company']}_{cy}.csv", mime="text/csv")

st.markdown("""
<div style="text-align:center;color:#555;font-size:0.8rem;padding:24px 0 8px;">
  GHG Protocol Framework · Emission factors: EPA CCCL 2023, DEFRA 2023, IEA 2023 · Built with Streamlit
</div>
""", unsafe_allow_html=True)
