"""
F&B Restaurant Sector — Integrated 3-Statement Financial Modeling Dashboard
============================================================================
Finance student sector benchmarking assignment tool.
Accepts Excel workbooks with raw Income Statement, Balance Sheet, and
Cash Flow Statement data for multi-unit restaurant concepts.

Run:  streamlit run fb_app.py
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
# PAGE CONFIG
# =============================================================================
st.set_page_config(
    page_title="F&B Financial Model",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# GLOBAL CSS
# =============================================================================
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
# LINE ITEM NORMALIZATION MAP
# Maps raw SEC filing label variants -> canonical keys
# Add entries here to support additional naming conventions
# =============================================================================
LINE_MAP: dict[str, str] = {
    # ── Revenue ───────────────────────────────────────────────────────────────
    "total revenue": "revenue",            "total revenues": "revenue",
    "net revenue": "revenue",              "net revenues": "revenue",
    "revenue": "revenue",                  "net sales": "revenue",
    "total net revenues": "revenue",       "total net revenue": "revenue",
    "restaurant revenue": "revenue",       "restaurant sales": "revenue",
    "company restaurant sales": "revenue", "shack sales": "revenue",

    # ── COGS / Food costs ─────────────────────────────────────────────────────
    "cost of sales": "cogs",
    "cost of goods sold": "cogs",
    "cost of revenue": "cogs",
    "restaurant operating costs": "cogs",
    "company-operated shop costs": "cogs",
    "food beverage packaging": "cogs",
    "food and paper costs": "cogs",
    "food beverage and packaging": "cogs",  # Chipotle / Sweetgreen exact label
    "food  beverage and packaging": "cogs", # trailing-space variant
    "food and beverage costs": "cogs",
    "food costs": "cogs",
    "cost of food and paper": "cogs",
    "product costs": "cogs",

    # ── Labor ─────────────────────────────────────────────────────────────────
    "labor": "labor",
    "labor expense": "labor",
    "labor and related costs": "labor",
    "labor costs": "labor",
    "store labor": "labor",
    "labor and related expenses": "labor",  # Shake Shack / Sweetgreen exact
    "restaurant labor": "labor",
    "crew labor": "labor",

    # ── Store / restaurant operating expenses ─────────────────────────────────
    "store operating expenses": "store_opex",
    "restaurant operating expenses": "store_opex",
    "occupancy and other costs": "store_opex",
    "occupancy costs": "store_opex",        # Chipotle exact
    "occupancy and related expenses": "store_opex",  # Shake Shack / Sweetgreen
    "other operating costs": "store_opex",  # Chipotle exact
    "other restaurant operating costs": "store_opex",  # Sweetgreen exact
    "other operating expenses": "store_opex",
    "pre-opening costs": "store_opex",
    "pre opening costs": "store_opex",      # normalized variant (no hyphen)
    "restaurant pre-opening costs": "store_opex",
    "preopening costs": "store_opex",

    # ── Advertising / marketing ───────────────────────────────────────────────
    "advertising expenses": "advertising",
    "advertising fees": "advertising",
    "advertising expense": "advertising",
    "marketing expense": "advertising",
    "marketing expenses": "advertising",

    # ── SGA ───────────────────────────────────────────────────────────────────
    "selling general and administrative": "sga",
    "sg&a": "sga",
    "sga": "sga",
    "general and administrative": "sga",
    "g&a": "sga",
    "general and administrative expenses": "sga",   # all three companies exact
    "general and administrative expense": "sga",
    "corporate general and administrative": "sga",

    # ── D&A ───────────────────────────────────────────────────────────────────
    "depreciation and amortization": "da",
    "d&a": "da",
    "depreciation": "da",
    "amortization": "da",
    "depreciation and amortization expense": "da",

    # ── Stock-based compensation (maps to a new key for CF add-back) ──────────
    "stock-based compensation": "sbc",
    "stock based compensation": "sbc",
    "share-based compensation": "sbc",
    "share based compensation expense": "sbc",

    # ── EBIT / Operating income ───────────────────────────────────────────────
    "operating income": "ebit",
    "operating income loss": "ebit",        # Shake Shack / Sweetgreen exact
    "income from operations": "ebit",
    "loss from operations": "ebit",
    "operating profit": "ebit",
    "income loss from operations": "ebit",

    # ── Interest ─────────────────────────────────────────────────────────────
    "interest expense": "interest",
    "interest expense net": "interest",
    "net interest expense": "interest",
    "interest expense income net": "interest",
    "interest income": "interest",
    "interest income expense net": "interest",

    # ── Tax ───────────────────────────────────────────────────────────────────
    "income tax expense": "tax",
    "income tax": "tax",
    "provision for income taxes": "tax",
    "benefit from income taxes": "tax",
    "income tax expense benefit": "tax",
    "income tax provision": "tax",

    # ── Net income ────────────────────────────────────────────────────────────
    "net income": "net_income",
    "net income loss": "net_income",
    "net loss": "net_income",
    "net income attributable": "net_income",
    "net loss attributable": "net_income",

    # ── Balance sheet: current assets ─────────────────────────────────────────
    "cash and cash equivalents": "cash",
    "cash": "cash",
    "total current assets": "current_assets",
    "accounts receivable": "accounts_receivable",
    "inventories": "inventories",
    "inventory": "inventories",
    "prepaid expenses and other current assets": "prepaid_other",
    "prepaid expenses": "prepaid_other",
    "marketable securities": "marketable_securities",   # Shake Shack
    "short-term investments": "short_term_investments", # Sweetgreen
    "short term investments": "short_term_investments",

    # ── Balance sheet: long-term assets ──────────────────────────────────────
    "property and equipment net": "ppe",
    "pp&e net": "ppe",
    "property plant and equipment net": "ppe",
    "property and equipment  net": "ppe",   # double-space variant
    "operating lease right-of-use assets": "rou_assets",
    "operating lease right of use assets": "rou_assets",
    "right-of-use assets": "rou_assets",
    "right of use assets net": "rou_assets",
    "goodwill": "goodwill",
    "goodwill and intangibles": "goodwill",
    "intangible assets and goodwill": "goodwill",      # Shake Shack exact
    "intangible assets net": "goodwill",
    "other non-current assets": "other_nca",
    "other noncurrent assets": "other_nca",
    "other non current assets": "other_nca",           # all three exact (normalized)
    "total assets": "total_assets",

    # ── Balance sheet: current liabilities ───────────────────────────────────
    "accounts payable": "accounts_payable",
    "accrued liabilities": "accrued_liabilities",
    "accrued expenses": "accrued_liabilities",          # Shake Shack exact
    "current operating lease liabilities": "cur_lease_lia",
    "current portion of operating lease liabilities": "cur_lease_lia",
    "total current liabilities": "current_liabilities",

    # ── Balance sheet: long-term liabilities ──────────────────────────────────
    "long-term debt": "ltd",
    "long term debt": "ltd",
    "total debt": "ltd",
    "notes payable": "ltd",
    "non-current operating lease liabilities": "lt_lease_lia",
    "noncurrent operating lease liabilities": "lt_lease_lia",
    "long-term operating lease liabilities": "lt_lease_lia",
    "other non-current liabilities": "other_ncl",
    "other noncurrent liabilities": "other_ncl",
    "other non current liabilities": "other_ncl",

    # ── Balance sheet: equity ─────────────────────────────────────────────────
    "total liabilities": "total_liabilities",
    "total equity": "equity",
    "total stockholders equity": "equity",
    "stockholders equity": "equity",
    "shareholders equity": "equity",
    "total stockholders deficit": "equity",
    "total liabilities and equity": "total_liab_equity",   # cross-check row
    "total liabilities and stockholders equity": "total_liab_equity",
    "retained earnings": "retained_earnings",
    "retained earnings accum deficit": "retained_earnings",  # Shake Shack exact
    "accumulated deficit": "retained_earnings",               # Sweetgreen exact
    "common stock and treasury stock": "common_stock",        # Chipotle exact
    "common stock and additional paid in capital": "common_stock", # SHAK/SG exact
    "common stock": "common_stock",

    # ── Cash flow: operating ──────────────────────────────────────────────────
    "cash provided by operating activities": "cfo",
    "net cash provided by operating": "cfo",
    "net cash provided by operating activities": "cfo",
    "operating cash flow": "cfo",
    "cash flows from operating activities": "cfo",
    "net cash from operating activities": "cfo",
    "changes in operating working capital": "working_capital_change",
    "change in operating working capital": "working_capital_change",
    "changes in working capital": "working_capital_change",

    # ── Cash flow: investing ──────────────────────────────────────────────────
    "purchases of property and equipment": "capex",
    "capital expenditures": "capex",
    "capex": "capex",
    "additions to property and equipment": "capex",
    "purchase of property and equipment": "capex",
    "cash used in investing activities": "cfi",             # all three exact
    "net cash used in investing activities": "cfi",
    "net cash from investing activities": "cfi",
    "other investing activities": "other_investing",
    "other investing activities net": "other_investing",

    # ── Cash flow: financing ──────────────────────────────────────────────────
    "net cash used in financing": "cff",
    "cash used in financing": "cff",
    "cash flow from financing": "cff",
    "cash used in financing activities": "cff",             # all three exact
    "net cash used in financing activities": "cff",
    "net cash from financing activities": "cff",
    "repurchase of common stock": "share_repurchase",       # Chipotle exact
    "repurchases of common stock": "share_repurchase",
    "other financing activities": "other_financing",
    "proceeds from stock option exercises": "stock_proceeds",  # Shake Shack exact
    "proceeds from stock issuance": "stock_proceeds",          # Sweetgreen exact
    "proceeds from exercise of stock options": "stock_proceeds",

    # ── Cash flow: other ──────────────────────────────────────────────────────
    "net increase in cash": "net_cash_change",
    "net increase decrease in cash": "net_cash_change",
    "net decrease in cash": "net_cash_change",
    "net change in cash": "net_cash_change",
}

DISPLAY_NAMES: dict[str, str] = {
    "revenue": "Net Revenue",
    "cogs": "Cost of Sales",
    "labor": "Labor and Related Costs",
    "store_opex": "Store Operating Expenses",
    "advertising": "Advertising Expense",
    "sga": "SG and A Expense",
    "da": "Depreciation and Amortization",
    "interest": "Net Interest Expense",
    "tax": "Income Tax Provision",
    "net_income": "Net Income",
    "cash": "Cash and Cash Equivalents",
    "ppe": "Property and Equipment Net",
    "goodwill": "Goodwill and Intangibles",
    "total_assets": "Total Assets",
    "ltd": "Long-Term Debt",
    "total_liabilities": "Total Liabilities",
    "equity": "Total Stockholders Equity",
    "cfo": "Operating Cash Flow",
    "capex": "Capital Expenditures",
    "cff": "Cash Flow from Financing",
}

IS_KEYS  = ["revenue","cogs","labor","store_opex","advertising","sga","da","interest","tax","net_income"]
BS_KEYS  = ["cash","ppe","goodwill","total_assets","ltd","total_liabilities","equity"]
CF_KEYS  = ["cfo","capex","cff"]


# =============================================================================
# DATA PARSING ENGINE
# =============================================================================
def normalize_label(raw: str) -> str | None:
    """
    Convert a raw SEC filing line-item label to a canonical key.

    Resolution order:
      1. Exact match after standard cleaning (lowercase, strip punctuation).
      2. Fuzzy substring rules — applied in priority order so that more
         specific patterns are tested before broader ones.  A rule fires
         when ALL required tokens appear in the cleaned string AND none of
         the optional exclusion tokens appear.

    Returns the canonical key string, or None if nothing matches.
    The caller (unmapped list) handles the None case.
    """
    # ── Stage 1: exact dictionary lookup ──────────────────────────────────────
    cleaned = (str(raw).strip().lower()
               .replace(",", "").replace(".", "")
               .replace("(", "").replace(")", "")
               .replace("-", " ").replace("/", " ")
               .replace("  ", " ").strip())

    if cleaned in LINE_MAP:
        return LINE_MAP[cleaned]

    # ── Stage 2: fuzzy substring rules ────────────────────────────────────────
    # Each entry: (required_tokens, exclude_tokens, canonical_key)
    # ALL required tokens must appear; ANY exclude token disqualifies the match.
    # Listed from most specific to least specific so earlier rules win.
    FUZZY_RULES: list[tuple[list[str], list[str], str]] = [
        # Revenue — must contain "revenue" or "sales" but not be a sub-line
        (["total net revenue"],   [],                          "revenue"),
        (["total net revenues"],  [],                          "revenue"),
        (["total revenue"],       ["cost", "other"],           "revenue"),
        (["net revenue"],         ["cost", "other"],           "revenue"),
        (["net revenues"],        ["cost", "other"],           "revenue"),
        (["restaurant revenue"],  [],                          "revenue"),
        (["shack sales"],         [],                          "revenue"),
        (["company restaurant"],  ["cost"],                    "revenue"),

        # Food / COGS — food + beverage + packaging cluster
        (["food", "beverage", "packaging"], [],                "cogs"),
        (["food", "packaging"],             [],                "cogs"),
        (["food", "beverage"],              ["labor","sga","admin"], "cogs"),
        (["food", "paper"],                 [],                "cogs"),
        (["cost of sales"],                 [],                "cogs"),
        (["cost of goods"],                 [],                "cogs"),
        (["cost of revenue"],               [],                "cogs"),

        # Labor
        (["labor", "related"],              [],                "labor"),
        (["labor", "expense"],              [],                "labor"),
        (["labor", "cost"],                 [],                "labor"),

        # Store operating / occupancy
        (["pre", "opening"],                [],                "store_opex"),
        (["pre-opening"],                   [],                "store_opex"),
        (["occupancy", "related"],          [],                "store_opex"),
        (["occupancy"],                     [],                "store_opex"),
        (["other restaurant operating"],    [],                "store_opex"),
        (["other operating"],               ["income","expense","non"],  "store_opex"),

        # SGA — general and administrative
        (["general", "administrative"],     [],                "sga"),

        # D&A
        (["depreciation", "amortization"],  [],                "da"),
        (["depreciation"],                  ["accumulated"],   "da"),

        # Stock-based compensation
        (["stock", "compensation"],         [],                "sbc"),
        (["share", "compensation"],         [],                "sbc"),
        (["stock-based"],                   [],                "sbc"),

        # EBIT / operating income
        (["operating income"],              ["non"],           "ebit"),
        (["income from operations"],        [],                "ebit"),
        (["loss from operations"],          [],                "ebit"),
        (["operating profit"],              [],                "ebit"),

        # Interest
        (["interest expense"],              [],                "interest"),
        (["net interest"],                  [],                "interest"),

        # Tax
        (["income tax"],                    [],                "tax"),
        (["provision for income"],          [],                "tax"),
        (["benefit from income tax"],       [],                "tax"),

        # Net income
        (["net income"],                    [],                "net_income"),
        (["net loss"],                      [],                "net_income"),

        # ── Balance sheet assets ──────────────────────────────────────────────
        (["total current assets"],          [],                "current_assets"),
        (["cash", "cash equivalents"],      [],                "cash"),
        (["accounts receivable"],           [],                "accounts_receivable"),
        (["inventories"],                   [],                "inventories"),
        (["inventory"],                     [],                "inventories"),
        (["prepaid"],                       [],                "prepaid_other"),
        (["marketable securities"],         [],                "marketable_securities"),
        (["short", "term", "investment"],   [],                "short_term_investments"),
        (["right-of-use"],                  ["current"],       "rou_assets"),
        (["right of use"],                  ["current"],       "rou_assets"),
        (["operating lease", "asset"],      ["current"],       "rou_assets"),
        (["intangible", "goodwill"],        [],                "goodwill"),
        (["goodwill"],                      [],                "goodwill"),
        (["property", "equipment", "net"],  [],                "ppe"),
        (["property", "plant", "equipment"],["purchases"],     "ppe"),
        (["other non-current assets"],      [],                "other_nca"),
        (["other noncurrent assets"],       [],                "other_nca"),
        (["other non current assets"],      [],                "other_nca"),
        (["total assets"],                  ["current","liab"],"total_assets"),

        # ── Balance sheet liabilities ─────────────────────────────────────────
        (["accounts payable"],              [],                "accounts_payable"),
        (["accrued liabilities"],           [],                "accrued_liabilities"),
        (["accrued expenses"],              [],                "accrued_liabilities"),
        (["current", "operating lease", "liabilit"], ["non"],  "cur_lease_lia"),
        (["total current liabilities"],     [],                "current_liabilities"),
        (["non-current", "operating lease"],["asset"],         "lt_lease_lia"),
        (["noncurrent", "operating lease"], ["asset"],         "lt_lease_lia"),
        (["long-term", "operating lease"],  ["asset"],         "lt_lease_lia"),
        (["non current", "operating lease"],["asset"],         "lt_lease_lia"),
        (["non current", "lease", "liabilit"], ["asset"],      "lt_lease_lia"),
        (["other non-current liabilities"], [],                "other_ncl"),
        (["other noncurrent liabilities"],  [],                "other_ncl"),
        (["other non current liabilities"], [],                "other_ncl"),
        (["long-term debt"],                [],                "ltd"),
        (["long term debt"],                [],                "ltd"),
        (["total liabilities"],             ["equity","and"],  "total_liabilities"),

        # ── Balance sheet equity ──────────────────────────────────────────────
        (["total liabilities", "equity"],   [],                "total_liab_equity"),
        (["total liabilities and"],         [],                "total_liab_equity"),
        (["retained earnings"],             [],                "retained_earnings"),
        (["accumulated deficit"],           [],                "retained_earnings"),
        (["common stock", "additional"],    [],                "common_stock"),
        (["common stock", "treasury"],      [],                "common_stock"),
        (["total equity"],                  [],                "equity"),
        (["total stockholders"],            [],                "equity"),
        (["stockholders equity"],           [],                "equity"),
        (["shareholders equity"],           [],                "equity"),

        # ── Cash flow: operating ──────────────────────────────────────────────
        (["net cash", "operating"],         [],                "cfo"),
        (["cash", "operating activities"],  ["used"],          "cfo"),
        (["changes in", "working capital"], [],                "working_capital_change"),
        (["change in", "working capital"],  [],                "working_capital_change"),

        # ── Cash flow: investing ──────────────────────────────────────────────
        (["purchases of property"],         [],                "capex"),
        (["purchase of property"],          [],                "capex"),
        (["additions to property"],         [],                "capex"),
        (["capital expenditures"],          [],                "capex"),
        (["cash used in investing"],        [],                "cfi"),
        (["net cash used in investing"],    [],                "cfi"),
        (["net cash", "investing"],         [],                "cfi"),
        (["other investing"],               [],                "other_investing"),

        # ── Cash flow: financing ──────────────────────────────────────────────
        (["cash used in financing"],        [],                "cff"),
        (["net cash used in financing"],    [],                "cff"),
        (["net cash", "financing"],         [],                "cff"),
        (["repurchase", "common stock"],    [],                "share_repurchase"),
        (["repurchases", "stock"],          [],                "share_repurchase"),
        (["proceeds from stock"],           [],                "stock_proceeds"),
        (["proceeds from exercise"],        [],                "stock_proceeds"),
        (["other financing"],               [],                "other_financing"),

        # ── Net cash change ───────────────────────────────────────────────────
        (["net increase", "cash"],          [],                "net_cash_change"),
        (["net decrease", "cash"],          [],                "net_cash_change"),
        (["net increase decrease", "cash"], [],                "net_cash_change"),
        (["net change in cash"],            [],                "net_cash_change"),
    ]

    for required_tokens, exclude_tokens, key in FUZZY_RULES:
        # All required tokens must appear in cleaned string
        if not all(tok in cleaned for tok in required_tokens):
            continue
        # No exclude token may appear
        if any(tok in cleaned for tok in exclude_tokens):
            continue
        return key

    return None


def _parse_one_sheet(df: pd.DataFrame, scale: float) -> dict:
    """
    Parse a single DataFrame (one sheet) into a partial company data dict.
    Returns { years, data, unmapped } or empty dict if nothing mapped.
    """
    df.columns = [str(c).strip() for c in df.columns]
    if df.empty or df.shape[1] < 2:
        return {}

    label_col = df.columns[0]
    year_cols = [
        c for c in df.columns[1:]
        if str(c).strip() not in ("", "nan")
        and not str(c).strip().lower().startswith("unnamed")
    ]
    if not year_cols:
        return {}

    result: dict = {"years": year_cols, "data": {}, "unmapped": []}
    for _, row in df.iterrows():
        raw_label = str(row[label_col]).strip()
        if not raw_label or raw_label.lower() in ("nan", "line item", "metric", ""):
            continue
        key = normalize_label(raw_label)
        if key is None:
            result["unmapped"].append(raw_label)
            continue
        if key not in result["data"]:
            result["data"][key] = {}
        for yr in year_cols:
            raw_val = str(row.get(yr, "")).strip().replace(",", "").replace("$", "")
            try:
                result["data"][key][yr] = float(raw_val) * scale
            except ValueError:
                result["data"][key][yr] = result["data"][key].get(yr, 0.0)

    return result if result["data"] else {}


# Keywords that identify a sheet as a specific statement type rather than a company.
# When ALL sheets in a workbook match these patterns the workbook is treated as
# a single-company multi-tab file and merged under the workbook filename.
_STMT_SHEET_KEYWORDS = {
    "income", "profit", "loss", "p&l", "earnings",
    "balance", "assets", "liabilities",
    "cash flow", "cashflow", "cash flows",
    "statement", "financials", "summary",
}


def _is_statement_sheet(sheet_name: str) -> bool:
    """Return True if the sheet name looks like a financial statement tab."""
    lower = sheet_name.lower()
    return any(kw in lower for kw in _STMT_SHEET_KEYWORDS)


def _merge_partial_data(parts: list[dict]) -> dict:
    """
    Merge multiple partial company data dicts (one per statement tab) into one.
    Years are unified across all parts; later parts do not overwrite earlier ones
    for the same (key, year) pair so that the first parsed value wins.
    """
    all_years: list[str] = []
    seen_years: set[str] = set()
    merged_data: dict = {}
    all_unmapped: list[str] = []

    for part in parts:
        for yr in part.get("years", []):
            if yr not in seen_years:
                all_years.append(yr)
                seen_years.add(yr)
        for key, yr_vals in part.get("data", {}).items():
            if key not in merged_data:
                merged_data[key] = {}
            for yr, val in yr_vals.items():
                if yr not in merged_data[key]:
                    merged_data[key][yr] = val
        all_unmapped.extend(part.get("unmapped", []))

    # Sort years numerically where possible
    def _yr_sort_key(yr: str) -> int:
        digits = "".join(c for c in yr if c.isdigit())
        return int(digits) if digits else 0

    all_years = sorted(set(all_years), key=_yr_sort_key)
    return {"years": all_years, "data": merged_data, "unmapped": list(set(all_unmapped))}


def parse_excel(file_obj, denomination: str, ticker_hint: str = "") -> dict[str, dict]:
    """
    Parse one uploaded Excel workbook and return a dict of company data.

    Two workbook layouts are handled automatically:

    Layout A — one sheet per company (e.g. a combined comps file):
      Each sheet is treated as one company.  The sheet name becomes the ticker.

    Layout B — one company across multiple statement sheets (e.g. a CMG file
      with tabs named Income Statement, Balance Sheet, Cash Flow Statement):
      All sheets are merged into a single company profile.  The workbook
      filename (or ticker_hint) becomes the ticker.

    Detection logic: if every sheet in the workbook has a name that matches a
    financial-statement keyword the workbook is treated as Layout B.  Otherwise
    each sheet with recognisable mapped data becomes its own company.

    Returns dict: { TICKER -> { years, data, unmapped } }
    """
    scale = 1.0 if denomination == "USD Thousands ($000s)" else 1000.0

    try:
        xl = pd.ExcelFile(file_obj)
    except Exception:
        return {}

    sheet_names = xl.sheet_names

    # Detect layout: if every sheet looks like a statement name -> Layout B
    all_statement_like = (
        len(sheet_names) > 0
        and all(_is_statement_sheet(s) for s in sheet_names)
    )

    companies: dict[str, dict] = {}

    if all_statement_like:
        # Layout B: merge all sheets into one company
        parts = []
        for sn in sheet_names:
            try:
                df = xl.parse(sn, header=0, index_col=None, dtype=str)
            except Exception:
                continue
            part = _parse_one_sheet(df, scale)
            if part:
                parts.append(part)

        if parts:
            merged = _merge_partial_data(parts)
            # Derive ticker from hint or use a cleaned version of the first sheet name
            ticker = (
                ticker_hint.upper().split(".")[0]  # strip file extension
                or sheet_names[0].upper()[:6]
            )
            # Remove common non-ticker words from the filename
            for drop in ("FINANCIALS", "MODEL", "DATA", "REPORT", "INCOME",
                         "BALANCE", "CASH", "FLOW", "STATEMENT"):
                ticker = ticker.replace(drop, "").strip("_- ")
            ticker = ticker[:8] or "CO"
            companies[ticker] = merged
    else:
        # Layout A: one company per sheet
        for sn in sheet_names:
            try:
                df = xl.parse(sn, header=0, index_col=None, dtype=str)
            except Exception:
                continue
            part = _parse_one_sheet(df, scale)
            if part:
                companies[sn.upper()[:12]] = part

    return companies


def parse_multiple_files(
    file_objs: list,
    denomination: str,
) -> dict[str, dict]:
    """
    Parse a list of uploaded file objects (each a separate company or comps file).
    Returns the merged company dict across all files.
    """
    all_companies: dict[str, dict] = {}
    for f in file_objs:
        # Use the filename (without extension) as the ticker hint for Layout B detection
        fname = getattr(f, "name", "")
        ticker_hint = fname.rsplit(".", 1)[0].upper() if fname else ""
        parsed = parse_excel(f, denomination, ticker_hint=ticker_hint)
        for ticker, cdata in parsed.items():
            # If a ticker already exists from another file, suffix with a number
            unique_ticker = ticker
            counter = 2
            while unique_ticker in all_companies:
                unique_ticker = f"{ticker}{counter}"
                counter += 1
            all_companies[unique_ticker] = cdata
    return all_companies


# =============================================================================
# METRICS ENGINE
# =============================================================================
def compute_metrics(cdata: dict) -> dict:
    """Compute derived F&B metrics for all historical years."""
    d, yrs = cdata["data"], cdata["years"]

    def g(key, yr):
        return d.get(key, {}).get(yr, 0.0) or 0.0

    result = {}
    for yr in yrs:
        rev   = g("revenue", yr) or 1
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
        net_debt   = ltd - cash          # strict definition: LTD minus Cash

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


def project_years(cdata: dict, drivers: dict, proj_labels: list[str]) -> dict:
    """
    Build projection years using sidebar driver assumptions.
    Includes a fully projected Balance Sheet linked to the Income Statement
    and Cash Flow Statement so all three statements balance each period.

    Income Statement drivers:
      rev_gr, cogs_pct, labor_pct, sga_pct, da_pct, int_pct, tax_rate

    Balance Sheet roll-forward:
      Cash    = prior cash + CFO + CFI + CFF
      PP&E    = prior PP&E + CapEx - D&A  (net book value)
      Equity  = prior equity + Net Income  (no dividends assumed unless specified)
      Ltd     = held flat at last actual  (no new issuances / repayments modeled)

    Cash Flow:
      CFO     = Net Income + D&A - NWC build
      CFI     = CapEx
      CFF     = 0 (placeholder; override with share_repurchase driver if desired)
    """
    d, yrs = cdata["data"], cdata["years"]
    last_yr = yrs[-1]

    def g(key):
        return d.get(key, {}).get(last_yr, 0.0) or 0.0

    # Seed balance sheet from last actual year
    prev_cash    = g("cash")
    prev_ppe     = g("ppe")
    prev_goodwill= g("goodwill")
    prev_equity  = g("equity")
    prev_ltd     = g("ltd")
    prev_tl      = g("total_liabilities")
    prev_rev     = g("revenue")

    proj: dict[str, dict] = {}

    for yr_label in proj_labels:
        # ── Income Statement ─────────────────────────────────────────────────
        rev   = prev_rev * (1 + drivers["rev_gr"])
        cogs  = -rev * drivers["cogs_pct"]
        labor = -rev * drivers.get("labor_pct", 0.0)
        sga   = -rev * drivers["sga_pct"]
        da    = -rev * drivers.get("da_pct", 0.04)
        ebitda = rev + cogs + labor + sga           # before D&A
        ebit  = ebitda + da
        int_  = -rev * drivers.get("int_pct", 0.03)
        ebt   = ebit + int_
        tax   = -max(0, ebt) * drivers["tax_rate"]
        ni    = ebt + tax
        net_margin = ni / rev if rev else 0

        # ── Cash Flow Statement ──────────────────────────────────────────────
        da_abs  = abs(da)
        nwc_chg = -rev * drivers.get("nwc_pct", 0.01)
        capex   = -rev * drivers["capex_pct"]
        cfo     = ni + da_abs + nwc_chg
        cfi     = capex                             # investing = CapEx only
        cff     = 0.0                               # financing flat (no debt assumed)
        net_cash_change = cfo + cfi + cff

        fcf    = cfo + capex
        fcf_margin = fcf / rev if rev else 0

        # ── Balance Sheet Roll-Forward ───────────────────────────────────────
        proj_cash   = prev_cash + net_cash_change
        proj_ppe    = max(0, prev_ppe + capex + da_abs)  # CapEx negative; da_abs positive
        # Note: capex is negative, da_abs is positive, so:
        #   proj_ppe = prior + (|CapEx| additions) - D&A depreciation
        proj_ppe    = max(0, prev_ppe - abs(capex) - da_abs)
        # Correct roll: additions INCREASE PPE; depreciation DECREASES it
        proj_ppe    = max(0, prev_ppe + abs(capex) - da_abs)

        proj_goodwill = prev_goodwill  # held flat (no new acquisitions)
        # Other assets held proportional to revenue growth as a simple approximation
        rev_growth_factor = rev / prev_rev if prev_rev else 1
        other_assets_approx = max(0, g("total_assets")
                                  - g("cash") - g("ppe") - g("goodwill")) * rev_growth_factor

        proj_total_assets = proj_cash + proj_ppe + proj_goodwill + other_assets_approx

        # Liabilities: LTD held flat; other liabilities grow with revenue
        proj_ltd  = prev_ltd
        other_liab_prior = max(0, prev_tl - prev_ltd)
        proj_other_liab  = other_liab_prior * rev_growth_factor
        proj_total_liab  = proj_ltd + proj_other_liab

        # Equity: prior + net income (retained earnings roll)
        proj_equity = prev_equity + ni

        # Net debt: LTD minus cash (strict definition)
        net_debt = proj_ltd - proj_cash

        proj[yr_label] = {
            # Income Statement
            "revenue":      rev,
            "cogs":         cogs,
            "labor":        labor,
            "sga":          sga,
            "da":           da,
            "ebit":         ebit,
            "interest":     int_,
            "tax":          tax,
            "net_income":   ni,
            "ebitda":       ebitda,
            "ebitda_margin":ebitda / rev if rev else 0,
            "net_margin":   net_margin,
            # Cash Flow
            "cfo":          cfo,
            "cfi":          cfi,
            "cff":          cff,
            "capex":        capex,
            "fcf":          fcf,
            "fcf_margin":   fcf_margin,
            "net_cash_change": net_cash_change,
            # Balance Sheet
            "bs_cash":      proj_cash,
            "bs_ppe":       proj_ppe,
            "bs_goodwill":  proj_goodwill,
            "bs_total_assets":  proj_total_assets,
            "bs_ltd":       proj_ltd,
            "bs_total_liab":proj_total_liab,
            "bs_equity":    proj_equity,
            "net_debt":     net_debt,
        }

        # Roll forward for next period
        prev_cash   = proj_cash
        prev_ppe    = proj_ppe
        prev_equity = proj_equity
        prev_ltd    = proj_ltd
        prev_tl     = proj_total_liab
        prev_rev    = rev

    return proj


# =============================================================================
# FORMATTING HELPERS
# =============================================================================
def fmt_val(v: float, fmt: str = "usd") -> str:
    if v == 0:
        return ""
    if fmt == "pct":
        return f"{v:.1%}"
    if fmt == "x":
        return f"{v:.1f}x"
    abs_v = abs(v)
    if abs_v >= 1_000_000:
        s = f"${abs_v/1_000_000:.1f}B"
    elif abs_v >= 1_000:
        s = f"${abs_v/1_000:.0f}M"
    else:
        s = f"${abs_v:,.0f}K"
    return f"({s})" if v < 0 else s


def build_is_df(cdata: dict, metrics: dict, proj: dict) -> pd.DataFrame:
    """Combine historical IS data with projections into one DataFrame."""
    d, yrs = cdata["data"], cdata["years"]
    all_cols = list(yrs) + list(proj.keys())

    rows = []
    is_line_order = [
        ("revenue", "Net Revenue"),
        ("cogs",    "Cost of Sales"),
        ("labor",   "Labor and Related Costs"),
        ("store_opex","Store Operating Expenses"),
        ("advertising","Advertising Expense"),
        ("sga",     "SG and A Expense"),
        ("_ebitda", "EBITDA"),
        ("da",      "Depreciation and Amortization"),
        ("interest","Net Interest Expense"),
        ("tax",     "Income Tax Provision"),
        ("net_income","Net Income"),
    ]
    for key, lbl in is_line_order:
        row = {"Line Item": lbl}
        for yr in yrs:
            if key == "_ebitda":
                row[yr] = metrics.get(yr, {}).get("ebitda", 0)
            else:
                v = d.get(key, {}).get(yr, 0) or 0
                row[yr] = v
        for yr, pdata in proj.items():
            if key == "_ebitda":
                row[yr] = pdata.get("ebitda", 0)
            elif key in pdata:
                row[yr] = pdata[key]
            elif key in d:
                row[yr] = d[key].get(list(yrs)[-1], 0)
            else:
                row[yr] = 0
        rows.append(row)

    # Add margin rows
    margin_rows = [
        ("EBITDA Margin", "ebitda_margin"),
        ("Net Margin",    "net_margin"),
    ]
    for lbl, mk in margin_rows:
        row = {"Line Item": lbl}
        for yr in yrs:
            row[yr] = metrics.get(yr, {}).get(mk, 0)
        for yr, pdata in proj.items():
            row[yr] = pdata.get(mk, 0)
        rows.append(row)

    return pd.DataFrame(rows)


def build_bs_df(cdata: dict, proj: dict = None) -> pd.DataFrame:
    """
    Build balance sheet DataFrame covering historical actuals and projected years.
    Historical columns come from raw data; projected columns come from project_years output.
    """
    d, yrs = cdata["data"], cdata["years"]
    proj   = proj or {}
    proj_yrs = list(proj.keys())

    bs_lines = [
        ("cash",             "Cash and Cash Equivalents",     "bs_cash"),
        ("ppe",              "Property and Equipment Net",    "bs_ppe"),
        ("goodwill",         "Goodwill and Intangibles",       "bs_goodwill"),
        ("total_assets",     "Total Assets",                  "bs_total_assets"),
        ("ltd",              "Long-Term Debt",                 "bs_ltd"),
        ("total_liabilities","Total Liabilities",             "bs_total_liab"),
        ("equity",           "Total Equity",                  "bs_equity"),
        ("_net_debt",        "Net Debt (LTD minus Cash)",     "net_debt"),
    ]
    rows = []
    for hist_key, lbl, proj_key in bs_lines:
        row = {"Line Item": lbl}
        # Historical actuals
        for yr in yrs:
            if hist_key == "_net_debt":
                ltd  = d.get("ltd",  {}).get(yr, 0) or 0
                cash = d.get("cash", {}).get(yr, 0) or 0
                row[yr] = ltd - cash
            else:
                row[yr] = d.get(hist_key, {}).get(yr, 0) or 0
        # Projected years from project_years output
        for yr in proj_yrs:
            row[yr] = proj[yr].get(proj_key, 0)
        rows.append(row)
    return pd.DataFrame(rows)


def build_cf_df(cdata: dict, metrics: dict, proj: dict) -> pd.DataFrame:
    d, yrs = cdata["data"], cdata["years"]
    cf_lines = [
        ("cfo",             "Operating Cash Flow (CFO)"),
        ("capex",           "Capital Expenditures"),
        ("_fcf",            "Free Cash Flow"),
        ("cff",             "Cash Flow from Financing (CFF)"),
        ("_net_cash_change","Net Change in Cash"),
        ("_net_margin",     "Net Margin %"),
    ]
    rows = []
    for key, lbl in cf_lines:
        row = {"Line Item": lbl}
        for yr in yrs:
            if key == "_fcf":
                row[yr] = metrics.get(yr, {}).get("fcf", 0)
            elif key == "_net_cash_change":
                # Historical: CFO + CapEx + CFF if available
                cfo_v   = d.get("cfo",  {}).get(yr, 0) or 0
                capex_v = d.get("capex",{}).get(yr, 0) or 0
                cff_v   = d.get("cff",  {}).get(yr, 0) or 0
                row[yr] = cfo_v + capex_v + cff_v
            elif key == "_net_margin":
                row[yr] = metrics.get(yr, {}).get("net_margin", 0)
            else:
                row[yr] = d.get(key, {}).get(yr, 0) or 0
        for yr, pdata in proj.items():
            if key == "_fcf":
                row[yr] = pdata.get("fcf", 0)
            elif key == "_net_cash_change":
                row[yr] = pdata.get("net_cash_change", 0)
            elif key == "_net_margin":
                row[yr] = pdata.get("net_margin", 0)
            elif key in pdata:
                row[yr] = pdata[key]
            else:
                row[yr] = 0
        rows.append(row)
    return pd.DataFrame(rows)


# =============================================================================
# DISPLAY HELPERS
# =============================================================================
def kpi_card(col, label: str, value: str, delta: str = "", neg: bool = False):
    delta_class = "kpi-neg" if neg else ""
    arrow = "▼" if neg else "▲"
    delta_html = (f'<div class="kpi-delta {delta_class}">{arrow} {delta}</div>'
                  if delta else "")
    col.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>
      {delta_html}
    </div>""", unsafe_allow_html=True)


def styled_table(df: pd.DataFrame, pct_rows: list[str] = None) -> pd.DataFrame:
    """Format a financial DataFrame for display."""
    pct_rows = pct_rows or []
    styled = df.copy()
    yr_cols = [c for c in df.columns if c != "Line Item"]
    for col in yr_cols:
        styled[col] = styled.apply(
            lambda row: (f"{row[col]:.1%}" if row["Line Item"] in pct_rows
                         else fmt_val(row[col])),
            axis=1
        )
    return styled


# =============================================================================
# PLOTLY CHART BUILDERS
# =============================================================================
COLORS = ["#60A5FA", "#34D399", "#FBBF24", "#F87171", "#A78BFA", "#FB923C"]
CHART_LAYOUT = dict(
    paper_bgcolor="#0F1117",
    plot_bgcolor="#161B27",
    font=dict(color="#A8C8E8", family="Arial", size=11),
    margin=dict(l=20, r=20, t=50, b=30),
    legend=dict(bgcolor="rgba(0,0,0,0.4)", bordercolor="#1F2937", borderwidth=1),
    xaxis=dict(gridcolor="#1F2937"),
    yaxis=dict(gridcolor="#1F2937"),
)


def chart_revenue_trend(companies: dict, metrics: dict[str,dict], proj_data: dict) -> go.Figure:
    fig = go.Figure()
    for i, (ticker, cdata) in enumerate(companies.items()):
        yrs = cdata["years"]
        revs = [cdata["data"].get("revenue",{}).get(yr,0)/1e3 for yr in yrs]
        proj_yrs = list(proj_data.get(ticker,{}).keys())
        proj_revs = [proj_data.get(ticker,{}).get(yr,{}).get("revenue",0)/1e3 for yr in proj_yrs]

        color = COLORS[i % len(COLORS)]
        fig.add_trace(go.Scatter(
            x=yrs, y=revs, name=ticker, mode="lines+markers",
            line=dict(color=color, width=2.5), marker=dict(size=7)))
        if proj_yrs:
            fig.add_trace(go.Scatter(
                x=proj_yrs, y=proj_revs, name=f"{ticker} Projected",
                mode="lines+markers",
                line=dict(color=color, width=2, dash="dash"),
                marker=dict(size=6, symbol="diamond")))

    fig.update_layout(title="Net Revenue Trend ($M)", height=380, **CHART_LAYOUT)
    return fig


def chart_ebitda_margin(companies: dict, metrics: dict[str,dict], proj_data: dict) -> go.Figure:
    fig = go.Figure()
    for i, (ticker, cdata) in enumerate(companies.items()):
        all_yrs, all_margins = [], []
        for yr in cdata["years"]:
            all_yrs.append(yr)
            all_margins.append(metrics[ticker].get(yr,{}).get("ebitda_margin",0)*100)
        for yr, pdata in proj_data.get(ticker,{}).items():
            all_yrs.append(yr)
            all_margins.append(pdata.get("ebitda_margin",0)*100)
        fig.add_trace(go.Bar(
            name=ticker, x=all_yrs, y=all_margins,
            marker_color=COLORS[i % len(COLORS)], opacity=0.85))

    fig.update_layout(title="EBITDA Margin % by Year",
                      barmode="group", height=380,
                      yaxis_ticksuffix="%", **CHART_LAYOUT)
    return fig


def chart_fcf_conversion(companies: dict, metrics: dict[str,dict], proj_data: dict) -> go.Figure:
    tickers, fcf_margins, cfo_vals = [], [], []
    for ticker, cdata in companies.items():
        last_yr = cdata["years"][-1]
        m = metrics[ticker].get(last_yr, {})
        tickers.append(ticker)
        fcf_margins.append(m.get("fcf_margin", 0) * 100)
        cfo_vals.append(m.get("cfo", 0) / 1e3)

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(name="FCF Margin %", x=tickers, y=fcf_margins,
                         marker_color="#60A5FA", opacity=0.85), secondary_y=False)
    fig.add_trace(go.Scatter(name="CFO ($M)", x=tickers, y=cfo_vals,
                             mode="markers+lines",
                             line=dict(color="#FBBF24", width=2),
                             marker=dict(size=10)), secondary_y=True)
    fig.update_layout(title="Free Cash Flow Conversion (Latest Year)",
                      height=380, **CHART_LAYOUT)
    fig.update_yaxes(title_text="FCF Margin %",  secondary_y=False, gridcolor="#1F2937")
    fig.update_yaxes(title_text="CFO ($M)",       secondary_y=True,  gridcolor="#1F2937")
    return fig


def chart_capital_structure(companies: dict, metrics: dict[str,dict]) -> go.Figure:
    fig = go.Figure()
    tickers = list(companies.keys())
    net_debt_vals, equity_vals = [], []
    for ticker, cdata in companies.items():
        last_yr = cdata["years"][-1]
        m = metrics[ticker].get(last_yr, {})
        net_debt_vals.append(m.get("net_debt", 0) / 1e3)
        equity_vals.append(cdata["data"].get("equity",{}).get(last_yr,0) / 1e3)

    fig.add_trace(go.Bar(name="Net Debt ($M)", x=tickers, y=net_debt_vals,
                         marker_color="#F87171", opacity=0.85))
    fig.add_trace(go.Bar(name="Total Equity ($M)", x=tickers, y=equity_vals,
                         marker_color="#34D399", opacity=0.85))
    fig.update_layout(title="Capital Structure Profile (Latest Year, $M)",
                      barmode="group", height=380, **CHART_LAYOUT)
    return fig


def chart_prime_cost(companies: dict, metrics: dict[str,dict]) -> go.Figure:
    fig = go.Figure()
    for i, (ticker, cdata) in enumerate(companies.items()):
        yrs  = cdata["years"]
        pcts = [metrics[ticker].get(yr,{}).get("prime_cost_pct",0)*100 for yr in yrs]
        fig.add_trace(go.Scatter(
            x=yrs, y=pcts, name=ticker, mode="lines+markers",
            line=dict(color=COLORS[i % len(COLORS)], width=2.5),
            marker=dict(size=7)))

    fig.add_hline(y=60, line_dash="dash", line_color="#FF6B6B",
                  annotation_text="60% Prime Cost Threshold",
                  annotation_position="top right")
    fig.update_layout(title="Prime Cost % Trend (COGS + Labor / Revenue)",
                      height=380, yaxis_ticksuffix="%", **CHART_LAYOUT)
    return fig


def chart_sensitivity(proj: dict, base_rev_gr: float, base_ebitda_margin: float) -> go.Figure:
    gr_range  = [base_rev_gr + x/100 for x in [-3, -2, -1, 0, 1, 2, 3]]
    ebitda_range = [base_ebitda_margin + x/100 for x in [-3, -1, 0, 1, 3]]

    fig = go.Figure(data=go.Heatmap(
        z=[[gr * m * 100 for m in ebitda_range] for gr in gr_range],
        x=[f"{m*100:.0f}% EBITDA Margin" for m in ebitda_range],
        y=[f"{g*100:.0f}% Revenue Growth"  for g in gr_range],
        colorscale=[[0,"#7F1D1D"],[0.5,"#FFF9C4"],[1,"#064E3B"]],
        text=[[f"{gr * m * 100:.1f}%" for m in ebitda_range] for gr in gr_range],
        texttemplate="%{text}",
        hovertemplate="Rev Growth: %{y}<br>EBITDA Margin: %{x}<br>EBITDA on Rev: %{text}<extra></extra>",
        colorbar=dict(title="EBITDA on Rev"),
    ))
    fig.update_layout(title="Sensitivity: Revenue Growth vs EBITDA Margin",
                      height=380, paper_bgcolor="#0F1117",
                      font=dict(color="#A8C8E8", family="Arial"),
                      margin=dict(l=20,r=20,t=50,b=30))
    return fig


# =============================================================================
# EXCEL EXPORT ENGINE
# =============================================================================
def build_export_excel(companies: dict, metrics: dict, proj_data: dict) -> bytes:
    """Generate a formatted multi-tab Excel workbook and return as bytes."""
    wb = Workbook()
    NAVY  = "1B365D"; STEEL = "2D5F8A"; LGRAY = "F2F4F8"
    WHITE = "FFFFFF"; DKGRAY = "333333"; MGRAY = "7F7F7F"
    BLUE  = "0070C0"; RED   = "C00000"

    def F(h): return PatternFill("solid", fgColor=h)
    def Fn(bold=False, color=DKGRAY, size=9, italic=False):
        return Font(bold=bold, color=color, size=size, italic=italic, name="Arial")
    def thin(c="C8C8C8"):
        s = Side(border_style="thin", color=c)
        return Border(left=s, right=s, top=s, bottom=s)
    def thick():
        t = Side(border_style="medium", color=NAVY)
        n = Side(border_style="thin",   color="C8C8C8")
        return Border(top=t, left=n, right=n, bottom=n)
    def AC(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
    def AL(): return Alignment(horizontal="left",   vertical="center")
    def AR(): return Alignment(horizontal="right",  vertical="center")

    FMT_USD = '$#,##0;[Red]($#,##0);"-"'
    FMT_PCT = '0.0%'

    def hdr_cell(ws, r, c, txt, bg=NAVY, fc=WHITE, sz=9):
        cell = ws.cell(r, c, txt)
        cell.font = Fn(bold=True, color=fc, size=sz)
        cell.fill = F(bg); cell.alignment = AC()
        cell.border = thin(WHITE)
        ws.row_dimensions[r].height = 22

    def write_section(ws, r, df, is_pct_rows=None):
        is_pct_rows = is_pct_rows or []
        yr_cols = [c for c in df.columns if c != "Line Item"]
        for _, row in df.iterrows():
            is_pct = row["Line Item"] in is_pct_rows
            is_tot = row["Line Item"] in ("EBITDA","Net Income","Total Assets","Free Cash Flow")
            bg = LGRAY if is_tot else WHITE
            cell = ws.cell(r, 2, row["Line Item"])
            cell.font = Fn(bold=is_tot, color=NAVY if is_tot else DKGRAY)
            cell.fill = F(bg); cell.alignment = AL(); cell.border = thin()
            ws.row_dimensions[r].height = 15
            for ci, yr in enumerate(yr_cols):
                v = row[yr] if not pd.isna(row[yr]) else 0
                c_cell = ws.cell(r, ci + 3, v)
                c_cell.number_format = FMT_PCT if is_pct else FMT_USD
                c_cell.font = Fn(bold=is_tot, color=NAVY if is_tot else (RED if v < 0 else DKGRAY))
                c_cell.fill = F(bg)
                c_cell.alignment = AR()
                c_cell.border = thick() if is_tot else thin()
            r += 1
        return r

    # Overview sheet
    ws_ov = wb.active
    ws_ov.title = "Overview"
    ws_ov.sheet_view.showGridLines = False
    ws_ov.column_dimensions["A"].width = 2
    ws_ov.column_dimensions["B"].width = 28
    for ci, ticker in enumerate(companies.keys()):
        ws_ov.column_dimensions[gcl(ci+3)].width = 16

    ws_ov.merge_cells(f"B1:{gcl(2+len(companies))}1")
    t = ws_ov.cell(1, 2, "F&B Restaurant Sector — 3-Statement Model Export")
    t.font = Fn(bold=True, color=WHITE, size=13)
    t.fill = F(NAVY); t.alignment = AL()
    ws_ov.row_dimensions[1].height = 30

    r = 3
    hdr_cell(ws_ov, r, 2, "Metric")
    for ci, ticker in enumerate(companies.keys()):
        hdr_cell(ws_ov, r, ci+3, ticker, STEEL)
    r += 1

    overview_metrics = [
        ("Net Revenue ($000s)",    "revenue",       FMT_USD),
        ("EBITDA ($000s)",         "ebitda",         FMT_USD),
        ("EBITDA Margin",          "ebitda_margin",  FMT_PCT),
        ("Net Income ($000s)",     "net_income",     FMT_USD),
        ("Net Margin",             "net_margin",     FMT_PCT),
        ("Free Cash Flow ($000s)", "fcf",            FMT_USD),
        ("FCF Margin",             "fcf_margin",     FMT_PCT),
        ("Net Debt ($000s)",       "net_debt",       FMT_USD),
        ("CapEx Pct Revenue",      "capex_pct",      FMT_PCT),
        ("Prime Cost Pct",         "prime_cost_pct", FMT_PCT),
    ]
    for lbl, mk, fmt in overview_metrics:
        bg = LGRAY if r % 2 == 0 else WHITE
        ws_ov.cell(r, 2, lbl).font = Fn(color=DKGRAY)
        ws_ov.cell(r, 2).fill = F(bg); ws_ov.cell(r, 2).alignment = AL()
        ws_ov.cell(r, 2).border = thin()
        for ci, ticker in enumerate(companies.keys()):
            last_yr = companies[ticker]["years"][-1]
            if mk == "net_income":
                v = companies[ticker]["data"].get("net_income",{}).get(last_yr,0)
            elif mk in ("revenue","ebitda"):
                v = metrics[ticker].get(last_yr, {}).get(mk if mk == "ebitda" else None,
                    companies[ticker]["data"].get("revenue",{}).get(last_yr,0))
                if mk == "ebitda":
                    v = metrics[ticker].get(last_yr, {}).get("ebitda", 0)
            else:
                v = metrics[ticker].get(last_yr, {}).get(mk, 0)
            cell = ws_ov.cell(r, ci+3, v)
            cell.number_format = fmt
            cell.font = Fn(color=RED if isinstance(v, float) and v < 0 else DKGRAY)
            cell.fill = F(bg); cell.alignment = AR(); cell.border = thin()
        ws_ov.row_dimensions[r].height = 15
        r += 1

    # Per-company sheets
    for ticker, cdata in companies.items():
        ws = wb.create_sheet(ticker)
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 2
        ws.column_dimensions["B"].width = 30

        yrs      = cdata["years"]
        proj_yrs = list(proj_data.get(ticker, {}).keys())
        all_yrs  = yrs + proj_yrs

        for ci, yr in enumerate(all_yrs):
            ws.column_dimensions[gcl(ci+3)].width = 14

        ws.merge_cells(f"B1:{gcl(2+len(all_yrs))}1")
        t = ws.cell(1, 2, f"{ticker} — Integrated 3-Statement Model")
        t.font = Fn(bold=True, color=WHITE, size=11)
        t.fill = F(NAVY); t.alignment = AL()
        ws.row_dimensions[1].height = 28

        r = 3
        hdr_cell(ws, r, 2, "Line Item")
        for ci, yr in enumerate(all_yrs):
            bg = STEEL if "E" in yr else "4A6741" if "E" in yr else STEEL
            bg = "4A6741" if "E" in yr else STEEL
            hdr_cell(ws, r, ci+3, yr, bg)
        r += 1

        is_df  = build_is_df(cdata, metrics[ticker], proj_data.get(ticker, {}))
        bs_df  = build_bs_df(cdata, proj_data.get(ticker, {}))
        cf_df  = build_cf_df(cdata, metrics[ticker], proj_data.get(ticker, {}))

        for section_title, df in [
            ("INCOME STATEMENT", is_df),
            ("BALANCE SHEET",    bs_df),
            ("CASH FLOW",        cf_df),
        ]:
            ws.merge_cells(f"B{r}:{gcl(2+len(all_yrs))}{r}")
            hdr_cell(ws, r, 2, f"  {section_title}", STEEL)
            r += 1
            pct_rows = ["EBITDA Margin", "Net Margin", "Net Margin %", "FCF Margin"]
            r = write_section(ws, r, df, pct_rows)
            r += 1

    wb_bytes = io.BytesIO()
    wb.save(wb_bytes)
    return wb_bytes.getvalue()


# =============================================================================
# DEMO DATA (used when no file is uploaded)
# =============================================================================
DEMO_COMPANIES = {
    "WING": {
        "years": ["FY2022","FY2023","FY2024"],
        "data": {
            "revenue":           {"FY2022":357521, "FY2023":460055, "FY2024":625807},
            "cogs":              {"FY2022":-63395,  "FY2023":-70646, "FY2024":-91632},
            "advertising":       {"FY2022":-123069, "FY2023":-166583,"FY2024":-233306},
            "sga":               {"FY2022":-67061,  "FY2023":-96898, "FY2024":-116801},
            "da":                {"FY2022":-10899,  "FY2023":-13239, "FY2024":-19490},
            "interest":          {"FY2022":-21230,  "FY2023":-18227, "FY2024":-21292},
            "tax":               {"FY2022":-16369,  "FY2023":-24135, "FY2024":-38473},
            "net_income":        {"FY2022":52947,   "FY2023":70175,  "FY2024":108717},
            "cash":              {"FY2022":184496,  "FY2023":90216,  "FY2024":315910},
            "ppe":               {"FY2022":66851,   "FY2023":91808,  "FY2024":127000},
            "goodwill":          {"FY2022":62514,   "FY2023":62514,  "FY2024":62514},
            "total_assets":      {"FY2022":424190,  "FY2023":377825, "FY2024":716246},
            "ltd":               {"FY2022":706846,  "FY2023":716000, "FY2024":720900},
            "total_liabilities": {"FY2022":815051,  "FY2023":835191, "FY2024":1391832},
            "equity":            {"FY2022":-390861, "FY2023":-457366,"FY2024":-675586},
            "cfo":               {"FY2022":76238,   "FY2023":121601, "FY2024":157610},
            "capex":             {"FY2022":-23940,  "FY2023":-40833, "FY2024":-51930},
        },
        "unmapped": [],
    },
    "BROS": {
        "years": ["FY2022","FY2023","FY2024"],
        "data": {
            "revenue":           {"FY2022":739002,  "FY2023":965600, "FY2024":1281015},
            "cogs":              {"FY2022":-576000, "FY2023":-762000,"FY2024":-1021056},
            "sga":               {"FY2022":-84000,  "FY2023":-128000,"FY2024":-153000},
            "da":                {"FY2022":-44728,  "FY2023":-69135, "FY2024":-93005},
            "interest":          {"FY2022":-21000,  "FY2023":-27000, "FY2024":-31000},
            "tax":               {"FY2022":-12000,  "FY2023":-24000, "FY2024":-12000},
            "net_income":        {"FY2022":-19253,  "FY2023":9952,   "FY2024":66450},
            "cash":              {"FY2022":20178,   "FY2023":133545, "FY2024":293354},
            "ppe":               {"FY2022":365468,  "FY2023":542440, "FY2024":750000},
            "goodwill":          {"FY2022":21629,   "FY2023":21629,  "FY2024":21629},
            "total_assets":      {"FY2022":1186360, "FY2023":1764010,"FY2024":2501085},
            "ltd":               {"FY2022":96297,   "FY2023":93175,  "FY2024":200000},
            "total_liabilities": {"FY2022":934384,  "FY2023":1088089,"FY2024":1737220},
            "equity":            {"FY2022":251976,  "FY2023":675921, "FY2024":763865},
            "cfo":               {"FY2022":59883,   "FY2023":139915, "FY2024":246432},
            "capex":             {"FY2022":-187880, "FY2023":-228457,"FY2024":-221738},
        },
        "unmapped": [],
    },
    "CAVA": {
        "years": ["FY2022","FY2023","FY2024"],
        "data": {
            "revenue":           {"FY2022":564119,  "FY2023":726100, "FY2024":963713},
            "cogs":              {"FY2022":-182000, "FY2023":-222000,"FY2024":-281000},
            "labor":             {"FY2022":-165000, "FY2023":-196000,"FY2024":-248000},
            "sga":               {"FY2022":-93000,  "FY2023":-118000,"FY2024":-156000},
            "da":                {"FY2022":-43086,  "FY2023":-47433, "FY2024":-63000},
            "interest":          {"FY2022":47,      "FY2023":8852,   "FY2024":16474},
            "tax":               {"FY2022":-93,     "FY2023":-768,   "FY2024":70409},
            "net_income":        {"FY2022":-58987,  "FY2023":13280,  "FY2024":130319},
            "cash":              {"FY2022":39125,   "FY2023":332428, "FY2024":366120},
            "ppe":               {"FY2022":242983,  "FY2023":330730, "FY2024":440000},
            "goodwill":          {"FY2022":1944,    "FY2023":1944,   "FY2024":1944},
            "total_assets":      {"FY2022":583883,  "FY2023":983757, "FY2024":1169669},
            "ltd":               {"FY2022":0,       "FY2023":0,      "FY2024":0},
            "total_liabilities": {"FY2022":370078,  "FY2023":412955, "FY2024":474103},
            "equity":            {"FY2022":-448503, "FY2023":570802, "FY2024":695566},
            "cfo":               {"FY2022":6038,    "FY2023":97101,  "FY2024":161027},
            "capex":             {"FY2022":-104161, "FY2023":-138806,"FY2024":-108131},
        },
        "unmapped": [],
    },
}


# =============================================================================
# MAIN APPLICATION
# =============================================================================
def main():
    # ── Sidebar ────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## 📊 F&B Financial Model")
        st.markdown("*Sector Benchmarking Dashboard*")
        st.divider()

        st.markdown("### 📂 Upload Data")
        uploaded = st.file_uploader(
            "Upload Excel workbook(s) (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            help=(
                "Drop one file per company, OR one combined comps file "
                "where each sheet is a company.  Single-company multi-tab "
                "files (e.g. CMG with Income Statement / Balance Sheet / "
                "Cash Flow tabs) are detected and merged automatically."
            ),
        )

        denomination = st.selectbox(
            "Input Denomination",
            ["USD Thousands ($000s)", "USD Millions ($M)"],
            index=0,
        )

        st.divider()
        st.markdown("### 🎛️ Projection Drivers")
        st.caption("Forecasts begin after the latest actual year in the data.")

        rev_gr   = st.slider("Revenue Growth Rate %",  0,  40, 12, 1) / 100
        cogs_pct = st.slider("Target COGS % Revenue",  10, 90, 32, 1) / 100
        labor_pct = st.slider("Labor % Revenue",       0,  40,  8, 1) / 100
        sga_pct  = st.slider("SG&A % Revenue",         3,  30, 14, 1) / 100
        da_pct   = st.slider("D&A % Revenue",          1,  15,  5, 1) / 100
        int_pct  = st.slider("Interest % Revenue",     0,  10,  3, 1) / 100
        capex_pct = st.slider("CapEx % Revenue",       1,  30,  8, 1) / 100
        tax_rate = st.slider("Effective Tax Rate %",   0,  40, 25, 1) / 100

        drivers = dict(
            rev_gr=rev_gr, cogs_pct=cogs_pct, labor_pct=labor_pct,
            sga_pct=sga_pct, da_pct=da_pct, int_pct=int_pct,
            capex_pct=capex_pct, tax_rate=tax_rate, nwc_pct=0.01,
        )

        st.divider()
        st.markdown("### ⚙️ Settings")
        show_demo = st.checkbox("Use demo data (WING/BROS/CAVA)", value=(not uploaded))

    # ── Load Data ──────────────────────────────────────────────────────────────
    if uploaded and not show_demo:
        with st.spinner("Parsing workbook(s) and normalizing line items…"):
            companies = parse_multiple_files(uploaded, denomination)
        if not companies:
            st.error(
                "No recognisable financial data found. "
                "Check that column A contains line item labels and "
                "remaining columns contain fiscal year values."
            )
            companies = {}
    else:
        companies = DEMO_COMPANIES

    if not companies:
        st.warning("Upload one or more Excel workbooks or enable demo data to begin.")
        return

    # ── Detect projection start year dynamically ───────────────────────────────
    # Find the highest numeric year present across all companies and start
    # projections from the year after that, labelled with an E suffix.
    def _extract_year_int(label: str) -> int:
        digits = "".join(c for c in str(label) if c.isdigit())
        return int(digits[:4]) if len(digits) >= 4 else 0

    all_hist_years: set[int] = set()
    for cdata in companies.values():
        for yr in cdata.get("years", []):
            y = _extract_year_int(yr)
            if y > 2000:
                all_hist_years.add(y)

    latest_actual_year = max(all_hist_years) if all_hist_years else 2024
    proj_start = latest_actual_year + 1
    proj_labels = [f"FY{proj_start}E", f"FY{proj_start+1}E", f"FY{proj_start+2}E"]

    # Compute metrics and projections
    metrics_all: dict[str, dict] = {}
    proj_all:    dict[str, dict] = {}

    for ticker, cdata in companies.items():
        metrics_all[ticker] = compute_metrics(cdata)
        proj_all[ticker]    = project_years(cdata, drivers, proj_labels)

    # ── Header ─────────────────────────────────────────────────────────────────
    st.markdown("""
    <div style="background:linear-gradient(135deg,#1B365D,#0D1B2E);
                border-radius:14px;padding:26px 34px;margin-bottom:24px;
                border:1px solid #2D5F8A;">
      <h1 style="color:#FFFFFF;margin:0;font-size:24px;font-weight:700;">
        📊 F&B Restaurant Sector — Integrated 3-Statement Model
      </h1>
      <p style="color:#A8C8E8;margin:8px 0 0;font-size:12px;">
        Sector Benchmarking Assignment · Finance Analytics Dashboard ·
        Institutional-Grade Financial Modeling
      </p>
    </div>
    """, unsafe_allow_html=True)

    # Company selector
    selected_tickers = st.multiselect(
        "Active Companies",
        options=list(companies.keys()),
        default=list(companies.keys()),
    )
    if not selected_tickers:
        st.warning("Select at least one company.")
        return

    active = {t: companies[t] for t in selected_tickers}
    active_metrics = {t: metrics_all[t] for t in selected_tickers}
    active_proj    = {t: proj_all[t]    for t in selected_tickers}

    # ── Tabs ───────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 Methodology",
        "🏢 Executive Overview",
        "📑 3-Statement Model",
        "📈 Visual Analytics",
    ])

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1: METHODOLOGY
    # ══════════════════════════════════════════════════════════════════════════
    with tab1:
        st.markdown('<div class="section-hdr">Modeling Framework and Assumptions</div>',
                    unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            for title, body in [
                ("3-Statement Linking Logic",
                 "Net Income from the Income Statement flows directly into Retained "
                 "Earnings on the Balance Sheet each projection year. Operating Cash "
                 "Flow from the Cash Flow Statement updates the ending Cash balance "
                 "on the Balance Sheet. These linkages ensure the three statements "
                 "remain internally consistent as any projection driver changes."),
                ("Net Debt Definition",
                 "Net Debt is defined strictly as Long-Term Debt minus Cash and Cash "
                 "Equivalents. This is the standard institutional definition used by "
                 "investment banks and private equity firms when sizing leverage. A "
                 "negative Net Debt figure indicates a net cash position, meaning the "
                 "company holds more cash than it owes in long-term obligations."),
                ("EBITDA Calculation",
                 "EBITDA is calculated as Net Revenue minus Cost of Sales, minus Labor "
                 "and Related Costs, minus Advertising Expense, minus Store Operating "
                 "Expenses, and minus SG&A. Depreciation and Amortization is excluded "
                 "from this calculation because it is a non-cash charge. This produces "
                 "a proxy for operating cash generation before financing and investment."),
            ]:
                st.markdown(f"""
                <div class="method-card">
                  <div class="method-title">{title}</div>
                  <div class="method-body">{body}</div>
                </div>""", unsafe_allow_html=True)

        with c2:
            for title, body in [
                ("Store-Level Scaling Assumptions",
                 "Projection drivers are applied as uniform percentages of revenue "
                 "across all uploaded concepts. This approach simplifies comparability "
                 "but understates company-specific unit economics differences such as "
                 "AUV mix shifts or new market entry costs. Users should adjust the "
                 "sidebar sliders to reflect each concept's disclosed guidance where "
                 "management has provided forward-looking ranges."),
                ("Prime Cost Threshold",
                 "The F&B industry benchmark for Prime Cost (the sum of COGS and "
                 "Labor as a percentage of revenue) is approximately 55 to 60 percent. "
                 "Concepts exceeding 65 percent typically face margin compression that "
                 "requires either pricing action or cost structure intervention. The "
                 "visual analytics tab highlights the 60 percent threshold on the "
                 "Prime Cost trend chart for easy identification."),
                ("Free Cash Flow Conversion",
                 "Free Cash Flow is defined as Operating Cash Flow plus Capital "
                 "Expenditures (which are negative outflows). A high-growth concept "
                 "with heavy unit expansion CapEx will often show negative or minimal "
                 "FCF despite strong store-level economics. FCF Conversion Rate is "
                 "FCF divided by Net Revenue and serves as the primary capital "
                 "efficiency metric in this benchmarking analysis."),
            ]:
                st.markdown(f"""
                <div class="method-card">
                  <div class="method-title">{title}</div>
                  <div class="method-body">{body}</div>
                </div>""", unsafe_allow_html=True)

        # Balance check
        st.markdown('<div class="section-hdr">3-Statement Balance Verification</div>',
                    unsafe_allow_html=True)
        balance_cols = st.columns(len(active))
        for i, (ticker, cdata) in enumerate(active.items()):
            last_yr = cdata["years"][-1]
            ta  = cdata["data"].get("total_assets",     {}).get(last_yr, 0) or 0
            tl  = cdata["data"].get("total_liabilities", {}).get(last_yr, 0) or 0
            eq  = cdata["data"].get("equity",            {}).get(last_yr, 0) or 0
            diff = ta - (tl + eq)
            ok = abs(diff) < abs(ta) * 0.05  # within 5% tolerance
            with balance_cols[i]:
                st.metric(
                    label=f"{ticker} — Balance Check ({last_yr})",
                    value=f"${diff:,.0f}K variance",
                    delta="Balanced" if ok else f"${diff:,.0f}K gap",
                    delta_color="normal" if ok else "inverse",
                )

        # Unmapped items
        any_unmapped = any(cdata.get("unmapped") for cdata in active.values())
        if any_unmapped:
            with st.expander("Unmapped Line Items (add to LINE_MAP to resolve)"):
                for ticker, cdata in active.items():
                    if cdata.get("unmapped"):
                        st.write(f"**{ticker}:** {', '.join(cdata['unmapped'])}")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2: EXECUTIVE OVERVIEW
    # ══════════════════════════════════════════════════════════════════════════
    with tab2:
        # KPI Cards — always show latest actual (historical) year
        for ticker, cdata in active.items():
            # latest_actual_year is the last year in cdata["years"] (historical only)
            last_actual = cdata["years"][-1]
            m   = active_metrics[ticker].get(last_actual, {})
            rev  = cdata["data"].get("revenue",    {}).get(last_actual, 0) or 0
            ni   = cdata["data"].get("net_income", {}).get(last_actual, 0) or 0
            ltd  = cdata["data"].get("ltd",        {}).get(last_actual, 0) or 0
            cash = cdata["data"].get("cash",       {}).get(last_actual, 0) or 0
            nd   = ltd - cash

            st.markdown(
                f'<div class="section-hdr">{ticker} '
                f'<span style="font-size:11px;color:#7F9DBF;font-weight:400;">'
                f'Latest Actual: {last_actual}</span></div>',
                unsafe_allow_html=True,
            )
            c1, c2, c3, c4, c5 = st.columns(5)
            kpi_card(c1, f"Net Revenue ({last_actual})", fmt_val(rev))
            kpi_card(c2, "EBITDA Margin",
                     f"{m.get('ebitda_margin', 0):.1%}",
                     delta="vs industry ~15%" if m.get("ebitda_margin", 0) > 0.15 else "")
            kpi_card(c3, "Net Income",
                     fmt_val(ni),
                     delta=f"Net Margin {m.get('net_margin', 0):.1%}",
                     neg=(ni < 0))
            kpi_card(c4, "Net Debt (LTD minus Cash)", fmt_val(nd), neg=(nd > 0))
            kpi_card(c5, "Free Cash Flow", fmt_val(m.get("fcf", 0)),
                     neg=(m.get("fcf", 0) < 0))

        # Cross-company benchmarking table
        st.markdown('<div class="section-hdr">Cross-Company Benchmarking Matrix</div>',
                    unsafe_allow_html=True)

        bench_rows = []
        for ticker, cdata in active.items():
            last_actual = cdata["years"][-1]
            m   = active_metrics[ticker].get(last_actual, {})
            rev = cdata["data"].get("revenue",    {}).get(last_actual, 0) or 0
            ni  = cdata["data"].get("net_income", {}).get(last_actual, 0) or 0
            ltd = cdata["data"].get("ltd",        {}).get(last_actual, 0) or 0
            csh = cdata["data"].get("cash",       {}).get(last_actual, 0) or 0
            bench_rows.append({
                "Ticker":           ticker,
                "Latest Year":      last_actual,
                "Net Revenue ($M)": f"${rev/1e3:.1f}M",
                "EBITDA Margin":    f"{m.get('ebitda_margin', 0):.1%}",
                "Net Margin":       f"{m.get('net_margin', 0):.1%}",
                "Net Income ($M)":  f"${ni/1e3:.1f}M",
                "FCF ($M)":         f"${m.get('fcf', 0)/1e3:.1f}M",
                "FCF Margin":       f"{m.get('fcf_margin', 0):.1%}",
                "Prime Cost Pct":   f"{m.get('prime_cost_pct', 0):.1%}",
                "Net Debt ($M)":    f"${(ltd - csh)/1e3:.1f}M",
                "CapEx Pct Rev":    f"{m.get('capex_pct', 0):.1%}",
            })

        bench_df = pd.DataFrame(bench_rows)
        st.dataframe(bench_df, hide_index=True, use_container_width=True,
                     height=min(200, 60 + 35 * len(bench_rows)))

        # Projection summary — includes Net Margin and CFF
        st.markdown(
            f'<div class="section-hdr">'
            f'{proj_labels[0]} to {proj_labels[-1]} Projection Summary</div>',
            unsafe_allow_html=True,
        )

        proj_rows = []
        for ticker in active:
            for yr, pdata in active_proj[ticker].items():
                rev_p = pdata.get("revenue", 0)
                proj_rows.append({
                    "Ticker":            ticker,
                    "Year":              yr,
                    "Revenue ($M)":      f"${rev_p/1e3:.1f}M",
                    "EBITDA Margin":     f"{pdata.get('ebitda_margin', 0):.1%}",
                    "Net Income ($M)":   f"${pdata.get('net_income', 0)/1e3:.1f}M",
                    "Net Margin %":      f"{pdata.get('net_margin', 0):.1%}",
                    "FCF ($M)":          f"${pdata.get('fcf', 0)/1e3:.1f}M",
                    "FCF Margin %":      f"{pdata.get('fcf_margin', 0):.1%}",
                    "CFF ($M)":          f"${pdata.get('cff', 0)/1e3:.1f}M",
                    "Proj Cash ($M)":    f"${pdata.get('bs_cash', 0)/1e3:.1f}M",
                    "Net Debt ($M)":     f"${pdata.get('net_debt', 0)/1e3:.1f}M",
                })
        proj_df = pd.DataFrame(proj_rows)
        st.dataframe(proj_df, hide_index=True, use_container_width=True,
                     height=min(400, 60 + 35 * len(proj_rows)))

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 3: 3-STATEMENT MODEL
    # ══════════════════════════════════════════════════════════════════════════
    with tab3:
        company_sel = st.selectbox("Select Company", list(active.keys()), key="stmt_sel")
        cdata  = active[company_sel]
        cmets  = active_metrics[company_sel]
        cproj  = active_proj[company_sel]

        is_df  = build_is_df(cdata, cmets, cproj)
        bs_df  = build_bs_df(cdata, cproj)
        cf_df  = build_cf_df(cdata, cmets, cproj)

        pct_rows = ["EBITDA Margin", "Net Margin", "Net Margin %"]

        def display_stmt(df, title):
            st.markdown(f'<div class="section-hdr">{title}</div>',
                        unsafe_allow_html=True)
            yr_cols   = [c for c in df.columns if c != "Line Item"]
            disp_df   = df.copy()
            for col in yr_cols:
                disp_df[col] = disp_df.apply(
                    lambda row, c=col: (
                        f"{row[c]:.1%}" if row["Line Item"] in pct_rows
                        else fmt_val(row[c] if not pd.isna(row[c]) else 0)
                    ), axis=1
                )

            def highlight_row(row):
                styles = []
                for val in row:
                    if row["Line Item"] in ("EBITDA","Net Income","Total Assets",
                                            "Free Cash Flow","Cash Flow from Operations"):
                        styles.append("background-color:#1B365D;font-weight:700;color:#EAF2FA")
                    elif row["Line Item"] in pct_rows:
                        styles.append("color:#7F7F7F;font-style:italic")
                    elif str(val).startswith("(") and "Line Item" not in str(row.name):
                        styles.append("color:#FF6B6B")
                    else:
                        styles.append("color:#EAF2FA")
                return styles

            st.dataframe(
                disp_df.style.apply(highlight_row, axis=1),
                hide_index=True,
                use_container_width=True,
                height=min(600, 38 + 35 * len(df)),
            )

        display_stmt(is_df, "Income Statement")
        display_stmt(bs_df, "Balance Sheet")
        display_stmt(cf_df, "Cash Flow Statement")

        # KPI block
        st.markdown('<div class="section-hdr">Key Performance Indicators</div>',
                    unsafe_allow_html=True)
        last_yr = cdata["years"][-1]
        m = cmets.get(last_yr, {})

        kpi_data = [
            ("Prime Cost Pct",      f"{m.get('prime_cost_pct',0):.1%}"),
            ("EBITDA Margin",       f"{m.get('ebitda_margin',0):.1%}"),
            ("Net Margin",          f"{m.get('net_margin',0):.1%}"),
            ("FCF Margin",          f"{m.get('fcf_margin',0):.1%}"),
            ("CapEx Pct Revenue",   f"{m.get('capex_pct',0):.1%}"),
            ("Net Debt ($000s)",    fmt_val(m.get("net_debt",0))),
        ]
        kpi_cols = st.columns(len(kpi_data))
        for col, (lbl, val) in zip(kpi_cols, kpi_data):
            kpi_card(col, lbl, val)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 4: VISUAL ANALYTICS
    # ══════════════════════════════════════════════════════════════════════════
    with tab4:
        r1c1, r1c2 = st.columns(2)
        with r1c1:
            st.plotly_chart(chart_revenue_trend(active, active_metrics, active_proj),
                            use_container_width=True)
        with r1c2:
            st.plotly_chart(chart_ebitda_margin(active, active_metrics, active_proj),
                            use_container_width=True)

        r2c1, r2c2 = st.columns(2)
        with r2c1:
            st.plotly_chart(chart_fcf_conversion(active, active_metrics, active_proj),
                            use_container_width=True)
        with r2c2:
            st.plotly_chart(chart_capital_structure(active, active_metrics),
                            use_container_width=True)

        r3c1, r3c2 = st.columns(2)
        with r3c1:
            st.plotly_chart(chart_prime_cost(active, active_metrics),
                            use_container_width=True)
        with r3c2:
            # Pick first selected company for sensitivity
            first_ticker = list(active.keys())[0]
            first_last   = active[first_ticker]["years"][-1]
            base_em = active_metrics[first_ticker].get(first_last,{}).get("ebitda_margin", 0.15)
            st.plotly_chart(chart_sensitivity(active_proj, rev_gr, base_em),
                            use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # EXPORT BUTTON
    # ══════════════════════════════════════════════════════════════════════════
    st.divider()
    st.markdown('<div class="section-hdr">Export</div>', unsafe_allow_html=True)
    with st.spinner("Building Excel export…"):
        excel_bytes = build_export_excel(active, active_metrics, active_proj)

    st.download_button(
        label="📥 Download 3-Statement Model (.xlsx)",
        data=excel_bytes,
        file_name="fb_3statement_model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Downloads a formatted multi-tab Excel workbook with historical data, projections, and comparative benchmarking.",
    )

    # Footer
    st.markdown("""
    <div style="text-align:center;color:#4B5563;font-size:11px;padding:16px 0 8px;">
      F&B Restaurant Sector Benchmarking Assignment ·
      Data sourced from SEC EDGAR 10-K filings ·
      Projections are analytical estimates based on disclosed guidance and historical trends ·
      <em>Not investment advice</em>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
