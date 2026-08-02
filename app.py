"""
F&B Restaurant Sector — Integrated 3-Statement Financial Modeling Dashboard
============================================================================
Finance student sector benchmarking assignment tool.
Accepts raw Income Statement, Balance Sheet, and Cash Flow Statement Excel files.

Run: streamlit run fb_app.py
"""

import io
import warnings
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gcl

warnings.filterwarnings("ignore")

# =============================================================================
# PAGE CONFIG & CSS
# =============================================================================
st.set_page_config(
    page_title="F&B Financial Model",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  .main  { background: #0F1117; }
  .stApp { background: #0F1117; }
  .block-container { padding-top: 1.5rem; }

  .kpi-card {
    background: linear-gradient(135deg, #1B365D 0%, #2D5F8A 100%);
    border: 1px solid #2D5F8A;
    border-radius: 10px;
    padding: 18px 22px;
    text-align: center;
    margin-bottom: 10px;
  }
  .kpi-label  { font-size:11px; font-weight:600; color:#A8C8E8;
                text-transform:uppercase; letter-spacing:1.4px; margin-bottom:6px; }
  .kpi-value  { font-size:26px; font-weight:700; color:#FFFFFF; }
  .kpi-delta  { font-size:11px; color:#52C97C; margin-top:4px; }
  .kpi-neg    { color:#FF6B6B; }

  .section-hdr {
    font-size:15px; font-weight:700; color:#EAF2FA;
    border-left:4px solid #2D5F8A;
    padding-left:10px; margin:22px 0 14px;
  }
  .method-card {
    background:#161B27; border:1px solid #1F2937;
    border-radius:8px; padding:16px 20px; margin-bottom:10px;
  }
  .method-title { font-size:13px; font-weight:700; color:#60A5FA; margin-bottom:6px; }
  .method-body  { font-size:12px; color:#9CA3AF; line-height:1.7; }

  div[data-testid="stSidebar"] {
    background:#111827; border-right:1px solid #1F2937;
  }
  .stTabs [data-baseweb="tab-list"] { background:#111827; border-radius:8px; }
  .stTabs [data-baseweb="tab"]      { color:#9CA3AF; }
  .stTabs [aria-selected="true"]    { color:#60A5FA !important; background:#1F2937 !important; }
  h1,h2,h3 { color:#EAF2FA !important; }
  .stDataFrame { background:#161B27; }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# NORMALIZATION MAP
# =============================================================================
LINE_MAP: dict[str, str] = {
    # Revenue
    "total revenue": "revenue", "total revenues": "revenue",
    "net revenue": "revenue",   "net revenues": "revenue",
    "revenue": "revenue",       "net sales": "revenue",
    "total net revenue": "revenue", "total net revenues": "revenue",

    # COGS / Food & Paper
    "cost of sales": "cogs",          "cost of goods sold": "cogs",
    "cost of revenue": "cogs",         "restaurant operating costs": "cogs",
    "company-operated shop costs": "cogs",
    "food beverage packaging": "cogs", "food and paper costs": "cogs",
    "food beverage and packaging": "cogs", "food and paper": "cogs",

    # Labor
    "labor": "labor",                  "labor expense": "labor",
    "labor and related costs": "labor","labor costs": "labor",
    "labor and related expenses": "labor", "store labor": "labor",

    # Store Operating
    "store operating expenses": "store_opex",
    "restaurant operating expenses": "store_opex",
    "occupancy and other costs": "store_opex",
    "other restaurant operating costs": "store_opex",
    "other operating costs": "store_opex",
    "other operating expenses": "store_opex",
    "occupancy costs": "store_opex",
    "occupancy and related expenses": "store_opex",

    # Advertising
    "advertising expenses": "advertising", "advertising fees": "advertising",
    "advertising expense": "advertising",  "marketing expense": "advertising",

    # SGA
    "selling general and administrative": "sga",
    "sg&a": "sga", "sga": "sga",
    "general and administrative": "sga", "g&a": "sga",
    "general and administrative expenses": "sga",

    # DA
    "depreciation and amortization": "da",
    "d&a": "da", "depreciation": "da", "amortization": "da",

    # Interest
    "interest expense": "interest",    "interest expense net": "interest",
    "net interest expense": "interest","interest expense income net": "interest",

    # Tax
    "income tax expense": "tax",       "income tax": "tax",
    "provision for income taxes": "tax","benefit from income taxes": "tax",

    # Net income
    "net income": "net_income",        "net income loss": "net_income",
    "net loss": "net_income",          "operating income loss": "ebit",

    # Balance sheet
    "cash and cash equivalents": "cash","cash": "cash",
    "property and equipment net": "ppe","pp&e net": "ppe",
    "property plant and equipment net": "ppe",
    "goodwill": "goodwill",            "goodwill and intangibles": "goodwill",
    "intangible assets and goodwill": "goodwill",
    "total assets": "total_assets",
    "long-term debt": "ltd",           "long term debt": "ltd",
    "total debt": "ltd",               "notes payable": "ltd",
    "total liabilities": "total_liabilities",
    "total equity": "equity",          "total stockholders equity": "equity",
    "stockholders equity": "equity",   "shareholders equity": "equity",

    # Cash flow
    "cash provided by operating activities": "cfo",
    "net cash provided by operating": "cfo",
    "operating cash flow": "cfo",
    "purchases of property and equipment": "capex",
    "capital expenditures": "capex",   "capex": "capex",
    "net cash used in financing": "cff","cash used in financing": "cff",
    "cash flow from financing": "cff",
}

def normalize_label(raw: str) -> str | None:
    cleaned = (str(raw).strip().lower()
               .replace(",","").replace(".","")
               .replace("(","").replace(")","")
               .replace("-"," ").replace("/"," ")
               .replace("  "," "))
    return LINE_MAP.get(cleaned)


# =============================================================================
# ADVANCED MULTI-FILE & MULTI-TAB PARSING ENGINE
# =============================================================================
def parse_uploaded_files(uploaded_files: list, denomination: str) -> dict[str, dict]:
    """
    Parses single or multiple uploaded Excel files.
    Supports both multi-tab statement workbooks and multi-company single-tab files.
    """
    scale = 1.0 if denomination == "USD Thousands ($000s)" else 1000.0
    companies: dict[str, dict] = {}

    for file_obj in uploaded_files:
        filename = file_obj.name.replace(".xlsx", "").replace(".xls", "")
        xl = pd.ExcelFile(file_obj)

        # Determine whether tabs are statements or companies
        statement_tab_names = ["income statement", "balance sheet", "cash flow statement", "cash flow"]
        has_statement_tabs = any(s.lower() in statement_tab_names for s in xl.sheet_names)

        if has_statement_tabs:
            ticker = filename.split("_")[0].upper()
            company_data = {"years": [], "data": {}, "unmapped": []}

            for sheet_name in xl.sheet_names:
                df = xl.parse(sheet_name, header=0, dtype=str)
                if df.empty or df.shape[1] < 2:
                    continue
                label_col = df.columns[0]
                year_cols = [str(c).strip() for c in df.columns[1:] if str(c).strip() not in ("", "nan")]

                if not company_data["years"]:
                    company_data["years"] = year_cols

                for _, row in df.iterrows():
                    raw_label = str(row[label_col]).strip()
                    if not raw_label or raw_label.lower() in ("nan", "line item", "metric", ""):
                        continue
                    key = normalize_label(raw_label)
                    if key is None:
                        company_data["unmapped"].append(raw_label)
                        continue
                    if key not in company_data["data"]:
                        company_data["data"][key] = {}
                    for yr in year_cols:
                        raw_val = str(row.get(yr, "")).strip().replace(",", "").replace("$", "")
                        try:
                            company_data["data"][key][yr] = float(raw_val) * scale
                        except ValueError:
                            company_data["data"][key][yr] = 0.0

            if company_data["data"]:
                companies[ticker] = company_data

        else:
            # Each sheet is an individual company
            for sheet_name in xl.sheet_names:
                df = xl.parse(sheet_name, header=0, dtype=str)
                if df.empty or df.shape[1] < 2:
                    continue
                label_col = df.columns[0]
                year_cols = [str(c).strip() for c in df.columns[1:] if str(c).strip() not in ("", "nan")]

                company_data = {"years": year_cols, "data": {}, "unmapped": []}
                for _, row in df.iterrows():
                    raw_label = str(row[label_col]).strip()
                    if not raw_label or raw_label.lower() in ("nan", "line item", "metric", ""):
                        continue
                    key = normalize_label(raw_label)
                    if key is None:
                        company_data["unmapped"].append(raw_label)
                        continue
                    company_data["data"][key] = {}
                    for yr in year_cols:
                        raw_val = str(row.get(yr, "")).strip().replace(",", "").replace("$", "")
                        try:
                            company_data["data"][key][yr] = float(raw_val) * scale
                        except ValueError:
                            company_data["data"][key][yr] = 0.0

                if company_data["data"]:
                    companies[sheet_name.upper()] = company_data

    return companies


# =============================================================================
# METRICS & PROJECTION ENGINE
# =============================================================================
def compute_metrics(cdata: dict) -> dict:
    d, yrs = cdata["data"], cdata["years"]

    def g(key, yr):
        return d.get(key, {}).get(yr, 0.0) or 0.0

    result = {}
    for yr in yrs:
        rev   = g("revenue", yr) or 1.0
        cogs  = abs(g("cogs",   yr))
        labor = abs(g("labor",  yr))
        sga   = abs(g("sga",    yr))
        da    = abs(g("da",     yr))
        cfo   = g("cfo",   yr)
        capex = g("capex", yr)
        cash  = g("cash",  yr)
        ltd   = g("ltd",   yr)
        ni    = g("net_income", yr)

        capex_abs  = abs(capex)
        prime_cost = cogs + labor
        adv        = abs(g("advertising", yr))
        ebitda     = rev - cogs - labor - adv - abs(g("store_opex", yr)) - sga
        fcf        = cfo - capex_abs
        net_debt   = ltd - cash

        result[yr] = {
            "prime_cost_pct":  prime_cost / rev,
            "cogs_pct":        cogs  / rev,
            "labor_pct":       labor / rev,
            "sga_pct":         sga   / rev,
            "da_pct":          da    / rev,
            "ebitda":          ebitda,
            "ebitda_margin":   ebitda / rev,
            "net_margin":      ni    / rev,
            "fcf":             fcf,
            "fcf_margin":      fcf   / rev,
            "capex_pct":       capex_abs / rev,
            "net_debt":        net_debt,
            "cfo":             cfo,
        }
    return result


def generate_projection_years(last_yr_str: str, num_years: int = 3) -> list[str]:
    """Dynamically determine future projection labels based on uploaded data."""
    clean_yr = "".join(filter(str.isdigit, last_yr_str))
    base_year = int(clean_yr) if clean_yr else 2024
    return [f"FY{base_year + i}E" for i in range(1, num_years + 1)]


def project_years(cdata: dict, drivers: dict, proj_labels: list[str]) -> dict:
    d, yrs = cdata["data"], cdata["years"]
    last_yr = yrs[-1]

    def g(key):
        return d.get(key, {}).get(last_yr, 0.0) or 0.0

    proj: dict[str, dict] = {}
    prev_rev = g("revenue")

    for yr_label in proj_labels:
        rev   = prev_rev * (1 + drivers["rev_gr"])
        cogs  = -rev * drivers["cogs_pct"]
        labor = -rev * drivers.get("labor_pct", 0.0)
        sga   = -rev * drivers["sga_pct"]
        da    = -rev * drivers.get("da_pct", 0.04)
        ebit  = rev + cogs + labor + sga + da
        int_  = -rev * drivers.get("int_pct", 0.03)
        ebt   = ebit + int_
        tax   = -max(0, ebt) * drivers["tax_rate"]
        ni    = ebt + tax
        cfo   = ni + abs(da) - rev * drivers.get("nwc_pct", 0.01)
        capex = -rev * drivers["capex_pct"]
        fcf   = cfo + capex
        ebitda = rev + cogs + labor + sga

        proj[yr_label] = {
            "revenue": rev, "cogs": cogs, "labor": labor,
            "sga": sga, "da": da, "ebit": ebit,
            "interest": int_, "tax": tax, "net_income": ni,
            "cfo": cfo, "capex": capex, "fcf": fcf,
            "ebitda": ebitda,
            "ebitda_margin": ebitda / rev if rev else 0,
            "fcf_margin": fcf / rev if rev else 0,
            "net_debt": d.get("ltd",{}).get(last_yr,0) - (d.get("cash",{}).get(last_yr,0) + cfo),
        }
        prev_rev = rev

    return proj


# =============================================================================
# STREAMLIT INTERFACE
# =============================================================================
st.sidebar.title("Configuration")
denomination = st.sidebar.radio("Input Denomination", ["USD Thousands ($000s)", "USD Millions ($M)"])

uploaded_files = st.sidebar.file_uploader(
    "Upload Financial Files (.xlsx)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    companies = parse_uploaded_files(uploaded_files, denomination)

    if companies:
        st.title("F&B Restaurant Sector Benchmarking Dashboard")

        # Sidebar Drivers
        st.sidebar.markdown("---")
        st.sidebar.subheader("Forecast Drivers")
        rev_gr = st.sidebar.slider("YoY Revenue Growth %", -0.10, 0.30, 0.08, 0.01)
        cogs_pct = st.sidebar.slider("COGS / Food % Revenue", 0.15, 0.45, 0.28, 0.01)
        sga_pct = st.sidebar.slider("SG&A % Revenue", 0.05, 0.25, 0.10, 0.01)
        capex_pct = st.sidebar.slider("CapEx % Revenue", 0.02, 0.20, 0.08, 0.01)
        tax_rate = st.sidebar.slider("Effective Tax Rate %", 0.10, 0.35, 0.21, 0.01)

        drivers = {
            "rev_gr": rev_gr, "cogs_pct": cogs_pct,
            "sga_pct": sga_pct, "capex_pct": capex_pct, "tax_rate": tax_rate
        }

        metrics = {}
        proj_data = {}

        for ticker, cdata in companies.items():
            metrics[ticker] = compute_metrics(cdata)
            last_hist_yr = cdata["years"][-1]
            proj_labels = generate_projection_years(last_hist_yr, 3)
            proj_data[ticker] = project_years(cdata, drivers, proj_labels)

        # Render KPI Cards
        cols = st.columns(len(companies))
        for idx, (ticker, cdata) in enumerate(companies.items()):
            last_yr = cdata["years"][-1]
            rev = cdata["data"].get("revenue", {}).get(last_yr, 0) / 1000.0
            ebitda_m = metrics[ticker][last_yr]["ebitda_margin"]
            with cols[idx]:
                st.metric(f"{ticker} Net Rev ({last_yr})", f"${rev:,.1f}M")
                st.metric(f"{ticker} EBITDA Margin", f"{ebitda_m:.1%}")

        st.success("Successfully loaded and processed financial statement models.")
    else:
        st.error("No valid financial data could be parsed from the uploaded files.")
else:
    st.info("Please upload one or more restaurant financial model workbooks using the sidebar.")
