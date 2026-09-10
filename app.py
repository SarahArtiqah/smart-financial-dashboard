import streamlit as st
import io
import re
import hashlib
from collections import OrderedDict
from pypdf import PdfReader
from docx import Document
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="FinSight", page_icon="📊", layout="wide")

# -----------------------------
# Session state
# -----------------------------
def init_state():
    defaults = {
        "page": "welcome",
        "user": None,
        "workspaces": [],
        "active_workspace": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# -----------------------------
# Styling
# -----------------------------
st.markdown(
    """
<style>
.main-title {
    font-size: 64px;
    font-weight: 800;
    text-align: center;
    margin-top: 45px;
}
.subtitle {
    font-size: 24px;
    text-align: center;
    color: #6b7280;
    margin-bottom: 28px;
}
.tagline {
    font-size: 28px;
    text-align: center;
    font-weight: 600;
    margin-top: 18px;
    line-height: 1.8;
}
.company-card, .summary-card {
    padding: 22px;
    border-radius: 18px;
    border: 1px solid #e5e7eb;
    background-color: #ffffff;
    margin-bottom: 18px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.04);
}
.sticky-notes {
    position: sticky;
    top: 80px;
    padding: 20px;
    border-radius: 18px;
    border: 1px solid #dbeafe;
    background-color: #eff6ff;
    box-shadow: 0 4px 16px rgba(0,0,0,0.05);
}
.pill {
    display: inline-block;
    padding: 6px 13px;
    border-radius: 999px;
    background-color: #f3f4f6;
    margin: 3px;
    font-size: 13px;
}
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------
# Helpers
# -----------------------------
def go(page):
    st.session_state.page = page
    st.rerun()


def sidebar_navigation():
    with st.sidebar:
        st.markdown("## FinSight")

        if st.button("ℹ️ About FinSight", width="stretch"):
            go("about_finsight")

        st.divider()

        if st.button("📂 Workspace", width="stretch"):
            go("home")

        if st.button("📊 Dashboard", width="stretch"):
            if st.session_state.active_workspace is not None:
                go("dashboard")
            else:
                st.warning("Create or open a workspace first.")

        if st.button("📝 Report", width="stretch"):
            go("report")

def back_button(label="← Back to Workspace Summary", target="workspace_summary"):
    if st.button(label, width="content"):
        go(target)


def year_options():
    return list(range(2000, datetime.now().year + 1))


def blank_financial_table(financial_years):
    metrics = [
        "Sales (Revenue)", "COGS", "EBIT", "Net Income", "Interest Expense",
        "Current Assets", "Inventory", "Accounts Receivable", "Current Liabilities",
        "Total Assets", "Total Liabilities", "Long-term Liabilities", "Total Equity",
        "Retained Earnings", "Cash Equivalents", "Marketable Securities", "PPE (Net)",
        "Shares Outstanding", "Dividends per Share", "Share Price",
        "Operating Cash Flow", "Depreciation Expense", "SG&A Expense",
    ]
    df = pd.DataFrame({"Financial Statement Item": metrics})
    for y in financial_years:
        df[str(y)] = None
    return df


# -----------------------------
# Sprint 4: Profitability engine
# -----------------------------
PROFITABILITY_FORMULAS = {
    "Gross Margin": "(Sales − COGS) ÷ Sales × 100",
    "Operating Margin": "EBIT ÷ Sales × 100",
    "Net Profit Margin": "Net Income ÷ Sales × 100",
    "ROA": "Net Income ÷ Average Total Assets × 100",
    "ROE": "Net Income ÷ Average Total Equity × 100",
    "ROCE": "EBIT ÷ Average Capital Employed × 100",
    "EPS": "Net Income ÷ Average Shares Outstanding",
}


def safe_number(value):
    """Convert a data-editor value into float; return None for blanks/invalid values."""
    if value is None or pd.isna(value):
        return None
    if isinstance(value, str):
        value = value.replace(",", "").strip()
        if value == "":
            return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_divide(numerator, denominator, multiplier=1.0):
    numerator = safe_number(numerator)
    denominator = safe_number(denominator)
    if numerator is None or denominator is None or denominator == 0:
        return None
    return (numerator / denominator) * multiplier


def financial_table_for_workspace(ws, company_key):
    """Return the live financial table stored for a company in the active workspace."""
    idx = st.session_state.active_workspace
    key = f"{company_key}_data_{idx}"
    if key not in st.session_state:
        st.session_state[key] = blank_financial_table(ws[company_key]["financial_years"])
    return st.session_state[key]


def table_to_financial_map(df):
    """Convert the input table to {item: {year: value}} for calculation use."""
    result = {}
    if df is None or df.empty or "Financial Statement Item" not in df.columns:
        return result
    for _, row in df.iterrows():
        item = str(row["Financial Statement Item"])
        result[item] = {
            str(col): safe_number(row[col])
            for col in df.columns
            if col != "Financial Statement Item"
        }
    return result


def average_balance(current_value, prior_value, is_first_year):
    """Use closing balance in first year; otherwise average opening and closing balances."""
    current_value = safe_number(current_value)
    prior_value = safe_number(prior_value)
    if current_value is None:
        return None
    if is_first_year or prior_value is None:
        return current_value
    return (prior_value + current_value) / 2


def calculate_profitability_ratios(df, financial_years):
    """Calculate profitability ratios using the locked FinSight methodology."""
    fmap = table_to_financial_map(df)
    years = [str(y) for y in financial_years]
    rows = []

    for i, year in enumerate(years):
        prior_year = years[i - 1] if i > 0 else None
        first_year = i == 0

        sales = fmap.get("Sales (Revenue)", {}).get(year)
        cogs = fmap.get("COGS", {}).get(year)
        ebit = fmap.get("EBIT", {}).get(year)
        net_income = fmap.get("Net Income", {}).get(year)
        total_assets = fmap.get("Total Assets", {}).get(year)
        total_equity = fmap.get("Total Equity", {}).get(year)
        current_liabilities = fmap.get("Current Liabilities", {}).get(year)
        shares = fmap.get("Shares Outstanding", {}).get(year)

        prior_assets = fmap.get("Total Assets", {}).get(prior_year) if prior_year else None
        prior_equity = fmap.get("Total Equity", {}).get(prior_year) if prior_year else None
        prior_current_liabilities = fmap.get("Current Liabilities", {}).get(prior_year) if prior_year else None
        prior_shares = fmap.get("Shares Outstanding", {}).get(prior_year) if prior_year else None

        avg_assets = average_balance(total_assets, prior_assets, first_year)
        avg_equity = average_balance(total_equity, prior_equity, first_year)
        avg_shares = average_balance(shares, prior_shares, first_year)

        capital_employed = None
        if total_assets is not None and current_liabilities is not None:
            capital_employed = total_assets - current_liabilities

        prior_capital_employed = None
        if prior_assets is not None and prior_current_liabilities is not None:
            prior_capital_employed = prior_assets - prior_current_liabilities

        avg_capital_employed = average_balance(
            capital_employed, prior_capital_employed, first_year
        )

        gross_profit = None
        if sales is not None and cogs is not None:
            gross_profit = sales - cogs

        rows.append({
            "Year": year,
            "Gross Margin": safe_divide(gross_profit, sales, 100),
            "Operating Margin": safe_divide(ebit, sales, 100),
            "Net Profit Margin": safe_divide(net_income, sales, 100),
            "ROA": safe_divide(net_income, avg_assets, 100),
            "ROE": safe_divide(net_income, avg_equity, 100),
            "ROCE": safe_divide(ebit, avg_capital_employed, 100),
            "EPS": safe_divide(net_income, avg_shares, 1),
        })

    return pd.DataFrame(rows)



# =========================================================
# Sprint 4 — Layer 1: FinSight Financial Calculator
# =========================================================

LAYER1_FORMULAS = {
    "Profitability": {
        "Gross Margin": "(Sales − COGS) ÷ Sales × 100",
        "Operating Margin": "EBIT ÷ Sales × 100",
        "Net Profit Margin": "Net Income ÷ Sales × 100",
        "ROA": "Net Income ÷ Average Total Assets × 100",
        "ROE": "Net Income ÷ Average Total Equity × 100",
        "ROCE": "EBIT ÷ Average Capital Employed × 100",
        "EPS": "Net Income ÷ Average Shares Outstanding",
    },
    "Liquidity": {
        "Working Capital": "Current Assets − Current Liabilities",
        "Current Ratio": "Current Assets ÷ Current Liabilities",
        "Quick Ratio": "(Current Assets − Inventory) ÷ Current Liabilities",
        "Cash Ratio": "(Cash Equivalents + Marketable Securities) ÷ Current Liabilities",
        "Operating Cash Flow Ratio": "Operating Cash Flow ÷ Current Liabilities",
    },
    "Efficiency": {
        "Asset Turnover": "Sales ÷ Average Total Assets",
        "Inventory Turnover": "COGS ÷ Average Inventory",
        "Days Inventory": "365 ÷ Inventory Turnover",
        "Receivable Turnover": "Sales ÷ Average Accounts Receivable",
        "Collection Period": "365 ÷ Receivable Turnover",
        "PPE Turnover": "Sales ÷ Average Net PPE",
    },
    "Solvency": {
        "Debt Ratio": "Total Liabilities ÷ Total Assets × 100",
        "Debt-to-Equity": "Total Liabilities ÷ Total Equity",
        "Equity Ratio": "Total Equity ÷ Total Assets × 100",
        "Long-term Debt-to-Equity": "Long-term Liabilities ÷ Total Equity",
        "Interest Coverage": "EBIT ÷ Interest Expense",
        "Financial Leverage": "Average Total Assets ÷ Average Total Equity",
    },
    "Market Performance": {
        "EPS": "Net Income ÷ Average Shares Outstanding",
        "Dividend Yield": "Dividends per Share ÷ Share Price × 100",
        "Dividend Payout": "Dividends per Share ÷ EPS × 100",
        "PE Ratio": "Share Price ÷ EPS",
        "Book Value per Share": "Total Equity ÷ Closing Shares Outstanding",
        "Market-to-Book": "Share Price ÷ Book Value per Share",
    },
    "Altman Z Score": {
        "Working Capital / Total Assets": "(Current Assets − Current Liabilities) ÷ Total Assets",
        "Retained Earnings / Total Assets": "Retained Earnings ÷ Total Assets",
        "EBIT / Total Assets": "EBIT ÷ Total Assets",
        "Market Value Equity / Total Liabilities": "(Share Price × Shares Outstanding) ÷ Total Liabilities",
        "Sales / Total Assets": "Sales ÷ Total Assets",
        "Altman Z Score": "1.2X₁ + 1.4X₂ + 3.3X₃ + 0.6X₄ + 1.0X₅",
    },
    "Piotroski F Score": {
        "Positive ROA": "1 point if ROA > 0",
        "Positive CFO": "1 point if Operating Cash Flow > 0",
        "Improving ROA": "1 point if current ROA > prior-year ROA",
        "Accrual Quality": "1 point if Operating Cash Flow > Net Income",
        "Lower Leverage": "1 point if long-term debt ratio decreases",
        "Improving Liquidity": "1 point if Current Ratio increases",
        "No New Shares": "1 point if Shares Outstanding does not increase",
        "Improving Gross Margin": "1 point if Gross Margin increases",
        "Improving Asset Turnover": "1 point if Asset Turnover increases",
        "Piotroski F Score": "Sum of nine binary financial signals (0–9)",
    },
    "Beneish M Score": {
        "DSRI": "(Receivables ÷ Sales)t ÷ (Receivables ÷ Sales)t−1",
        "GMI": "Prior Gross Margin ÷ Current Gross Margin",
        "AQI": "[1 − (Current Assets + Net PPE) ÷ Total Assets]t ÷ prior-year value",
        "SGI": "Salest ÷ Salest−1",
        "DEPI": "Prior Depreciation Rate ÷ Current Depreciation Rate",
        "SGAI": "(SG&A ÷ Sales)t ÷ (SG&A ÷ Sales)t−1",
        "TATA": "(Net Income − Operating Cash Flow) ÷ Total Assets",
        "LVGI": "Current leverage ratio ÷ prior leverage ratio",
        "Beneish M Score": "-4.84 + 0.920DSRI + 0.528GMI + 0.404AQI + 0.892SGI + 0.115DEPI − 0.172SGAI + 4.679TATA − 0.327LVGI",
    },
}


def financial_value(fmap, item, year):
    if year is None:
        return None
    return fmap.get(item, {}).get(str(year))


def calculate_liquidity_ratios(df, financial_years):
    fmap = table_to_financial_map(df)
    rows = []
    for year in [str(y) for y in financial_years]:
        ca = financial_value(fmap, "Current Assets", year)
        inv = financial_value(fmap, "Inventory", year)
        cl = financial_value(fmap, "Current Liabilities", year)
        cash = financial_value(fmap, "Cash Equivalents", year)
        securities = financial_value(fmap, "Marketable Securities", year)
        cfo = financial_value(fmap, "Operating Cash Flow", year)

        working_capital = None if ca is None or cl is None else ca - cl
        quick_assets = None if ca is None or inv is None else ca - inv
        cash_assets = None
        if cash is not None or securities is not None:
            cash_assets = (cash or 0) + (securities or 0)

        rows.append({
            "Year": year,
            "Working Capital": working_capital,
            "Current Ratio": safe_divide(ca, cl),
            "Quick Ratio": safe_divide(quick_assets, cl),
            "Cash Ratio": safe_divide(cash_assets, cl),
            "Operating Cash Flow Ratio": safe_divide(cfo, cl),
        })
    return pd.DataFrame(rows)


def calculate_efficiency_ratios(df, financial_years):
    fmap = table_to_financial_map(df)
    years = [str(y) for y in financial_years]
    rows = []

    for i, year in enumerate(years):
        prior = years[i - 1] if i > 0 else None
        first = i == 0

        sales = financial_value(fmap, "Sales (Revenue)", year)
        cogs = financial_value(fmap, "COGS", year)
        assets = financial_value(fmap, "Total Assets", year)
        inventory = financial_value(fmap, "Inventory", year)
        receivables = financial_value(fmap, "Accounts Receivable", year)
        ppe = financial_value(fmap, "PPE (Net)", year)

        avg_assets = average_balance(
            assets, financial_value(fmap, "Total Assets", prior), first
        )
        avg_inventory = average_balance(
            inventory, financial_value(fmap, "Inventory", prior), first
        )
        avg_receivables = average_balance(
            receivables, financial_value(fmap, "Accounts Receivable", prior), first
        )
        avg_ppe = average_balance(
            ppe, financial_value(fmap, "PPE (Net)", prior), first
        )

        inventory_turnover = safe_divide(cogs, avg_inventory)
        receivable_turnover = safe_divide(sales, avg_receivables)

        rows.append({
            "Year": year,
            "Asset Turnover": safe_divide(sales, avg_assets),
            "Inventory Turnover": inventory_turnover,
            "Days Inventory": safe_divide(365, inventory_turnover),
            "Receivable Turnover": receivable_turnover,
            "Collection Period": safe_divide(365, receivable_turnover),
            "PPE Turnover": safe_divide(sales, avg_ppe),
        })

    return pd.DataFrame(rows)


def calculate_solvency_ratios(df, financial_years):
    fmap = table_to_financial_map(df)
    years = [str(y) for y in financial_years]
    rows = []

    for i, year in enumerate(years):
        prior = years[i - 1] if i > 0 else None
        first = i == 0

        assets = financial_value(fmap, "Total Assets", year)
        liabilities = financial_value(fmap, "Total Liabilities", year)
        long_term_liabilities = financial_value(fmap, "Long-term Liabilities", year)
        equity = financial_value(fmap, "Total Equity", year)
        ebit = financial_value(fmap, "EBIT", year)
        interest = financial_value(fmap, "Interest Expense", year)

        avg_assets = average_balance(
            assets, financial_value(fmap, "Total Assets", prior), first
        )
        avg_equity = average_balance(
            equity, financial_value(fmap, "Total Equity", prior), first
        )

        rows.append({
            "Year": year,
            "Debt Ratio": safe_divide(liabilities, assets, 100),
            "Debt-to-Equity": safe_divide(liabilities, equity),
            "Equity Ratio": safe_divide(equity, assets, 100),
            "Long-term Debt-to-Equity": safe_divide(long_term_liabilities, equity),
            "Interest Coverage": safe_divide(ebit, interest),
            "Financial Leverage": safe_divide(avg_assets, avg_equity),
        })

    return pd.DataFrame(rows)


def calculate_market_ratios(df, financial_years):
    fmap = table_to_financial_map(df)
    years = [str(y) for y in financial_years]
    rows = []

    for i, year in enumerate(years):
        prior = years[i - 1] if i > 0 else None
        first = i == 0

        net_income = financial_value(fmap, "Net Income", year)
        shares = financial_value(fmap, "Shares Outstanding", year)
        dps = financial_value(fmap, "Dividends per Share", year)
        price = financial_value(fmap, "Share Price", year)
        equity = financial_value(fmap, "Total Equity", year)

        avg_shares = average_balance(
            shares, financial_value(fmap, "Shares Outstanding", prior), first
        )
        eps = safe_divide(net_income, avg_shares)
        bvps = safe_divide(equity, shares)

        rows.append({
            "Year": year,
            "EPS": eps,
            "Dividend Yield": safe_divide(dps, price, 100),
            "Dividend Payout": safe_divide(dps, eps, 100),
            "PE Ratio": safe_divide(price, eps),
            "Book Value per Share": bvps,
            "Market-to-Book": safe_divide(price, bvps),
        })

    return pd.DataFrame(rows)


def calculate_altman_z(df, financial_years):
    fmap = table_to_financial_map(df)
    rows = []

    for year in [str(y) for y in financial_years]:
        sales = financial_value(fmap, "Sales (Revenue)", year)
        ebit = financial_value(fmap, "EBIT", year)
        ca = financial_value(fmap, "Current Assets", year)
        cl = financial_value(fmap, "Current Liabilities", year)
        assets = financial_value(fmap, "Total Assets", year)
        liabilities = financial_value(fmap, "Total Liabilities", year)
        retained = financial_value(fmap, "Retained Earnings", year)
        shares = financial_value(fmap, "Shares Outstanding", year)
        price = financial_value(fmap, "Share Price", year)

        wc = None if ca is None or cl is None else ca - cl
        market_equity = None if shares is None or price is None else shares * price

        x1 = safe_divide(wc, assets)
        x2 = safe_divide(retained, assets)
        x3 = safe_divide(ebit, assets)
        x4 = safe_divide(market_equity, liabilities)
        x5 = safe_divide(sales, assets)

        components = [x1, x2, x3, x4, x5]
        z_score = None
        if all(v is not None for v in components):
            z_score = 1.2*x1 + 1.4*x2 + 3.3*x3 + 0.6*x4 + 1.0*x5

        rows.append({
            "Year": year,
            "Working Capital / Total Assets": x1,
            "Retained Earnings / Total Assets": x2,
            "EBIT / Total Assets": x3,
            "Market Value Equity / Total Liabilities": x4,
            "Sales / Total Assets": x5,
            "Altman Z Score": z_score,
        })

    return pd.DataFrame(rows)


def calculate_piotroski_f(df, financial_years):
    fmap = table_to_financial_map(df)
    profitability = calculate_profitability_ratios(df, financial_years)
    liquidity = calculate_liquidity_ratios(df, financial_years)
    efficiency = calculate_efficiency_ratios(df, financial_years)
    years = [str(y) for y in financial_years]
    rows = []

    for i, year in enumerate(years):
        if i == 0:
            rows.append({
                "Year": year,
                "Positive ROA": None,
                "Positive CFO": None,
                "Improving ROA": None,
                "Accrual Quality": None,
                "Lower Leverage": None,
                "Improving Liquidity": None,
                "No New Shares": None,
                "Improving Gross Margin": None,
                "Improving Asset Turnover": None,
                "Piotroski F Score": None,
            })
            continue

        prior = years[i - 1]
        current_row = profitability.iloc[i]
        prior_row = profitability.iloc[i - 1]
        current_liq = liquidity.iloc[i]
        prior_liq = liquidity.iloc[i - 1]
        current_eff = efficiency.iloc[i]
        prior_eff = efficiency.iloc[i - 1]

        cfo = financial_value(fmap, "Operating Cash Flow", year)
        net_income = financial_value(fmap, "Net Income", year)
        shares = financial_value(fmap, "Shares Outstanding", year)
        prior_shares = financial_value(fmap, "Shares Outstanding", prior)
        lt_debt = financial_value(fmap, "Long-term Liabilities", year)
        prior_lt_debt = financial_value(fmap, "Long-term Liabilities", prior)
        assets = financial_value(fmap, "Total Assets", year)
        prior_assets = financial_value(fmap, "Total Assets", prior)

        current_lev = safe_divide(lt_debt, assets)
        prior_lev = safe_divide(prior_lt_debt, prior_assets)

        signals = {
            "Positive ROA": 1 if pd.notna(current_row["ROA"]) and current_row["ROA"] > 0 else 0,
            "Positive CFO": 1 if cfo is not None and cfo > 0 else 0,
            "Improving ROA": 1 if pd.notna(current_row["ROA"]) and pd.notna(prior_row["ROA"]) and current_row["ROA"] > prior_row["ROA"] else 0,
            "Accrual Quality": 1 if cfo is not None and net_income is not None and cfo > net_income else 0,
            "Lower Leverage": 1 if current_lev is not None and prior_lev is not None and current_lev < prior_lev else 0,
            "Improving Liquidity": 1 if pd.notna(current_liq["Current Ratio"]) and pd.notna(prior_liq["Current Ratio"]) and current_liq["Current Ratio"] > prior_liq["Current Ratio"] else 0,
            "No New Shares": 1 if shares is not None and prior_shares is not None and shares <= prior_shares else 0,
            "Improving Gross Margin": 1 if pd.notna(current_row["Gross Margin"]) and pd.notna(prior_row["Gross Margin"]) and current_row["Gross Margin"] > prior_row["Gross Margin"] else 0,
            "Improving Asset Turnover": 1 if pd.notna(current_eff["Asset Turnover"]) and pd.notna(prior_eff["Asset Turnover"]) and current_eff["Asset Turnover"] > prior_eff["Asset Turnover"] else 0,
        }
        score = sum(signals.values())
        signals["Year"] = year
        signals["Piotroski F Score"] = score
        rows.append(signals)

    columns = [
        "Year", "Positive ROA", "Positive CFO", "Improving ROA",
        "Accrual Quality", "Lower Leverage", "Improving Liquidity",
        "No New Shares", "Improving Gross Margin",
        "Improving Asset Turnover", "Piotroski F Score"
    ]
    return pd.DataFrame(rows)[columns]


def calculate_beneish_m(df, financial_years):
    fmap = table_to_financial_map(df)
    years = [str(y) for y in financial_years]
    rows = []

    for i, year in enumerate(years):
        if i == 0:
            rows.append({
                "Year": year, "DSRI": None, "GMI": None, "AQI": None,
                "SGI": None, "DEPI": None, "SGAI": None, "TATA": None,
                "LVGI": None, "Beneish M Score": None
            })
            continue

        prior = years[i - 1]

        def v(item, y):
            return financial_value(fmap, item, y)

        sales_t, sales_p = v("Sales (Revenue)", year), v("Sales (Revenue)", prior)
        cogs_t, cogs_p = v("COGS", year), v("COGS", prior)
        ar_t, ar_p = v("Accounts Receivable", year), v("Accounts Receivable", prior)
        ca_t, ca_p = v("Current Assets", year), v("Current Assets", prior)
        ppe_t, ppe_p = v("PPE (Net)", year), v("PPE (Net)", prior)
        ta_t, ta_p = v("Total Assets", year), v("Total Assets", prior)
        dep_t, dep_p = v("Depreciation Expense", year), v("Depreciation Expense", prior)
        sga_t, sga_p = v("SG&A Expense", year), v("SG&A Expense", prior)
        cl_t, cl_p = v("Current Liabilities", year), v("Current Liabilities", prior)
        ltd_t, ltd_p = v("Long-term Liabilities", year), v("Long-term Liabilities", prior)
        ni_t = v("Net Income", year)
        cfo_t = v("Operating Cash Flow", year)

        dsri = safe_divide(safe_divide(ar_t, sales_t), safe_divide(ar_p, sales_p))

        gm_t = safe_divide(None if sales_t is None or cogs_t is None else sales_t - cogs_t, sales_t)
        gm_p = safe_divide(None if sales_p is None or cogs_p is None else sales_p - cogs_p, sales_p)
        gmi = safe_divide(gm_p, gm_t)

        aq_t = None
        aq_p = None
        if ta_t is not None and ca_t is not None and ppe_t is not None:
            aq_t = 1 - safe_divide(ca_t + ppe_t, ta_t)
        if ta_p is not None and ca_p is not None and ppe_p is not None:
            aq_p = 1 - safe_divide(ca_p + ppe_p, ta_p)
        aqi = safe_divide(aq_t, aq_p)

        sgi = safe_divide(sales_t, sales_p)

        dep_rate_t = safe_divide(dep_t, None if dep_t is None or ppe_t is None else dep_t + ppe_t)
        dep_rate_p = safe_divide(dep_p, None if dep_p is None or ppe_p is None else dep_p + ppe_p)
        depi = safe_divide(dep_rate_p, dep_rate_t)

        sgai = safe_divide(safe_divide(sga_t, sales_t), safe_divide(sga_p, sales_p))
        tata = safe_divide(None if ni_t is None or cfo_t is None else ni_t - cfo_t, ta_t)

        lev_t = safe_divide(None if cl_t is None or ltd_t is None else cl_t + ltd_t, ta_t)
        lev_p = safe_divide(None if cl_p is None or ltd_p is None else cl_p + ltd_p, ta_p)
        lvgi = safe_divide(lev_t, lev_p)

        components = [dsri, gmi, aqi, sgi, depi, sgai, tata, lvgi]
        m_score = None
        if all(x is not None for x in components):
            m_score = (
                -4.84
                + 0.920 * dsri
                + 0.528 * gmi
                + 0.404 * aqi
                + 0.892 * sgi
                + 0.115 * depi
                - 0.172 * sgai
                + 4.679 * tata
                - 0.327 * lvgi
            )

        rows.append({
            "Year": year,
            "DSRI": dsri,
            "GMI": gmi,
            "AQI": aqi,
            "SGI": sgi,
            "DEPI": depi,
            "SGAI": sgai,
            "TATA": tata,
            "LVGI": lvgi,
            "Beneish M Score": m_score,
        })

    return pd.DataFrame(rows)


def calculate_all_layer1(df, financial_years):
    """Single Layer 1 entry point used by workspace, demo, dashboard and future report."""
    return {
        "Profitability": calculate_profitability_ratios(df, financial_years),
        "Liquidity": calculate_liquidity_ratios(df, financial_years),
        "Efficiency": calculate_efficiency_ratios(df, financial_years),
        "Solvency": calculate_solvency_ratios(df, financial_years),
        "Market Performance": calculate_market_ratios(df, financial_years),
        "Altman Z Score": calculate_altman_z(df, financial_years),
        "Piotroski F Score": calculate_piotroski_f(df, financial_years),
        "Beneish M Score": calculate_beneish_m(df, financial_years),
    }


def layer1_output(ws, company_key):
    company = ws[company_key]
    table = financial_table_for_workspace(ws, company_key)
    return calculate_all_layer1(table, company["financial_years"])


def format_layer1_table(df):
    if df is None or df.empty:
        return df
    output = df.copy()
    for col in output.columns:
        if col == "Year":
            continue
        output[col] = output[col].map(
            lambda x: "N/A" if x is None or pd.isna(x) else f"{float(x):,.4f}"
        )
    return output

def profitability_output(ws, company_key):
    company = ws[company_key]
    table = financial_table_for_workspace(ws, company_key)
    return calculate_profitability_ratios(table, company["financial_years"])


def latest_valid_pair(series):
    clean = [float(x) for x in series if x is not None and not pd.isna(x)]
    if len(clean) < 2:
        return None, None
    return clean[-2], clean[-1]


def profitability_signal(result_df):
    """Create a neutral category trend and key flag from actual ratio movements."""
    if result_df is None or result_df.empty:
        return "⚪ Insufficient Data", "Review whether all required profitability inputs have been entered."

    indicators = ["Gross Margin", "Operating Margin", "Net Profit Margin", "ROA", "ROE", "ROCE"]
    changes = []
    for item in indicators:
        previous, latest = latest_valid_pair(result_df[item].tolist())
        if previous is None or latest is None:
            continue
        tolerance = max(abs(previous) * 0.01, 0.05)
        if latest > previous + tolerance:
            changes.append(1)
        elif latest < previous - tolerance:
            changes.append(-1)
        else:
            changes.append(0)

    if len(changes) < 2:
        return "⚪ Insufficient Data", "Review whether all required profitability inputs have been entered."

    positive = sum(x == 1 for x in changes)
    negative = sum(x == -1 for x in changes)
    if positive >= 4 and positive > negative:
        return "🟢 Improving", "Verify whether the improvement is supported by sustainable operating performance."
    if negative >= 4 and negative > positive:
        return "🔴 Requires Attention", "Examine the causes of weakening profitability and whether the decline is temporary or recurring."
    return "🟡 Mixed Trend", "Review the profitability indicators that moved in different directions before forming a conclusion."


def format_ratio_value(item, value):
    if value is None or pd.isna(value):
        return "N/A"
    if item == "EPS":
        return f"{value:,.4f}"
    return f"{value:,.2f}%"


def profitability_guidance(result_df):
    """Generate C3-level guidance from actual latest movements without forming C5 conclusions."""
    if result_df is None or result_df.empty:
        return {
            "observation": "Profitability analysis is unavailable because the required data are incomplete.",
            "evidence": "At least Sales, COGS, EBIT, Net Income, Total Assets, Total Equity, Current Liabilities and Shares Outstanding are required.",
            "investigate": "Review the data-entry table and complete the missing financial statement items.",
            "annual_report": "Refer to the audited financial statements and relevant notes before entering the figures.",
        }

    movements = []
    for item in ["Gross Margin", "Operating Margin", "Net Profit Margin", "ROA", "ROE", "ROCE", "EPS"]:
        previous, latest = latest_valid_pair(result_df[item].tolist())
        if previous is None or latest is None:
            continue
        if latest > previous:
            direction = "increased"
        elif latest < previous:
            direction = "decreased"
        else:
            direction = "remained unchanged"
        movements.append(f"{item} {direction} from {format_ratio_value(item, previous)} to {format_ratio_value(item, latest)}")

    evidence = "; ".join(movements[:3]) + ("." if movements else "Insufficient comparable ratio results are available.")
    signal, _ = profitability_signal(result_df)
    observation = f"The latest profitability pattern is classified as {signal.replace('🟢 ', '').replace('🟡 ', '').replace('🔴 ', '').replace('⚪ ', '').lower()}."

    return {
        "observation": observation,
        "evidence": evidence,
        "investigate": "Review whether the movements were driven by revenue growth, cost control, asset utilisation, financing structure, changes in equity, or one-off gains and losses.",
        "annual_report": "Refer to the statement of profit or loss, statement of financial position, Management Discussion & Analysis, segment information and relevant notes.",
    }


def workspace_name(ws):
    a = ws["company_a"]["company_name"]
    if ws["analysis_type"] == "Compare Two Companies":
        return f"{a} vs {ws['company_b']['company_name']}"
    return a


def progress_for_workspace(ws):
    total = len(ws["company_a"]["financial_years"]) * 20
    if ws["analysis_type"] == "Compare Two Companies":
        total += len(ws["company_b"]["financial_years"]) * 20
    return 0 if total == 0 else min(int(ws.get("completed_cells", 0) / total * 100), 100)

def workflow_progress(ws):
    if ws["analysis_type"] == "Compare Two Companies":
        steps = [
            ("Workspace Created", True),
            ("Company A Data Entered", ws.get("company_a_completed", False)),
            ("Company B Data Entered", ws.get("company_b_completed", False)),
            ("Dashboard Ready", ws.get("company_a_completed", False) and ws.get("company_b_completed", False)),
        ]
    else:
        steps = [
            ("Workspace Created", True),
            ("Company A Data Entered", ws.get("company_a_completed", False)),
            ("Dashboard Ready", ws.get("company_a_completed", False)),
        ]
    completed = sum(1 for _, done in steps if done)
    progress = completed / len(steps)
    return progress, steps


def company_card(prefix, title):
    st.markdown("<div class='company-card'>", unsafe_allow_html=True)
    st.markdown(f"### {title}")
    company_name = st.text_input(f"{title} Name", placeholder="Example: Nestlé Malaysia Berhad", key=f"{prefix}_name")
    activity = st.selectbox(
        f"{title} Primary Business Activity",
        [
            "Manufacturing", "Consumer Products", "Retail", "Service", "Banking",
            "Insurance", "Property Development", "Plantation", "Technology",
            "Healthcare", "Airline / Transportation", "Energy", "Telecommunications",
            "Utilities", "Hospitality", "Others",
        ],
        key=f"{prefix}_activity",
    )
    currency = st.selectbox(
        f"{title} Currency",
        ["RM (MYR)", "USD", "SGD", "GBP", "EUR", "AED", "JPY", "CNY", "THB", "IDR"],
        key=f"{prefix}_currency",
    )
    fye = st.selectbox(
        f"{title} Financial Year-End",
        ["31 December", "31 March", "30 June", "30 September", "Other"],
        key=f"{prefix}_fye",
    )
    if fye == "Other":
        fye = st.text_input(f"{title} Financial Year-End - Other", placeholder="Example: 31 January", key=f"{prefix}_fye_other")
    st.caption("Tip: Refer to the company profile, business review, or principal activities section in the annual report.")
    st.markdown("</div>", unsafe_allow_html=True)
    return {"company_name": company_name, "business_activity": activity, "currency": currency, "financial_year_end": fye}


def confirm_years(label, years, prefix):
    st.markdown(f"#### {label}")
    selected = []
    if years:
        cols = st.columns(min(len(years), 5))
        for i, y in enumerate(years):
            with cols[i % len(cols)]:
                if st.checkbox(str(y), value=True, key=f"{prefix}_year_{y}"):
                    selected.append(y)
    excluded = [y for y in years if y not in selected]
    reason = ""
    if excluded:
        reason = st.selectbox(
            f"Reason for excluded year(s) - {label}",
            ["Change in financial year-end", "Newly listed company", "Company restructuring", "Annual report not available", "Other"],
            key=f"{prefix}_reason",
        )
    return selected, excluded, reason


def notes_panel():
    st.markdown(
        """
        <div class="sticky-notes">
        <h4>📌 Workspace Notes</h4>
        <p><b>Need to update company information?</b><br>
        You can edit the Company Profile before completing financial data entry.</p>
        <hr>
        <p><b>Financial years</b><br>
        The years selected here will automatically appear in the data entry table.</p>
        <hr>
        <p><b>Do not panic 😄</b><br>
        If something looks wrong, go back and amend the Company Profile before entering the financial numbers.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def company_summary_card(company, label):
    st.markdown("<div class='summary-card'>", unsafe_allow_html=True)
    st.markdown(f"### {label}")
    st.markdown(f"## {company['company_name']}")
    st.markdown(
        f"""
        <span class="pill">{company['business_activity']}</span>
        <span class="pill">{company['currency']}</span>
        <span class="pill">{company['financial_year_end']}</span>
        """,
        unsafe_allow_html=True,
    )
    st.write("**Financial Years:**")
    st.write(" • ".join(map(str, company["financial_years"])))
    if company.get("excluded_years"):
        st.warning(f"Excluded year(s): {', '.join(map(str, company['excluded_years']))}. Reason: {company.get('exclusion_reason', '-')}")
    st.markdown("</div>", unsafe_allow_html=True)

# -----------------------------
# Pages
# -----------------------------
def welcome_page():
    st.markdown("<div class='main-title'>FinSight</div>", unsafe_allow_html=True)
    st.markdown("<div class='subtitle'>Smart Financial Analytics Platform</div>", unsafe_allow_html=True)
    st.markdown("<div class='tagline'>Analyse.<br>Compare.<br>Interpret.</div>", unsafe_allow_html=True)
    st.write("")
    _, mid, _ = st.columns([1, 1, 1])
    with mid:
        if st.button("Login", width="stretch"):
            go("login")
        if st.button("Create Account", width="stretch"):
            go("register")


def login_page():
    st.title("Login to FinSight")
    email = st.text_input("Email")
    st.text_input("Password", type="password")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Login", width="stretch"):
            if email:
                st.session_state.user = {"name": email.split("@")[0].title(), "email": email}
                go("home")
            else:
                st.warning("Please enter your email.")
    with c2:
        if st.button("Back", width="stretch"):
            go("welcome")


def register_page():
    st.title("Create Account")
    name = st.text_input("Full Name")
    student_id = st.text_input("Student ID")
    university = st.text_input("University", value="Universiti Poly-Tech Malaysia")
    email = st.text_input("Email")
    password = st.text_input("Password", type="password")
    confirm = st.text_input("Confirm Password", type="password")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Create Account", width="stretch"):
            if not name or not email or not password:
                st.warning("Please complete name, email, and password.")
            elif password != confirm:
                st.error("Passwords do not match.")
            else:
                st.session_state.user = {"name": name, "student_id": student_id, "university": university, "email": email}
                go("home")
    with c2:
        if st.button("Back", width="stretch"):
            go("welcome")


def home_page():
    user = st.session_state.user or {"name": "Student"}
    st.title(f"👋 Welcome back, {user['name']}!")
    st.caption("Your personal financial analysis workspace.")
    c1, c2, c3 = st.columns([1.2, 1.2, 1])
    with c1:
        if st.button("➕ Create New Workspace", width="stretch"):
            go("create_workspace")
    with c2:
        st.button("📥 Import Excel (optional backup)", width="stretch", disabled=True)
    with c3:
        if st.button("Logout", width="stretch"):
            st.session_state.user = None
            go("welcome")
    st.divider()
    st.subheader("My Workspaces")
    if not st.session_state.workspaces:
        st.info("No workspace yet. Create your first financial analysis workspace.")
        return
    for i, ws in enumerate(st.session_state.workspaces):
        progress = progress_for_workspace(ws)
        with st.container(border=True):
            left, right = st.columns([4, 1.2])
            with left:
                st.markdown(f"### 📂 {ws['analysis_type']}")
                st.markdown(f"**{workspace_name(ws)}**")
                st.write(f"**Company A Financial Years:** {', '.join(map(str, ws['company_a']['financial_years']))}")
                if ws["analysis_type"] == "Compare Two Companies":
                    st.write(f"**Company B Financial Years:** {', '.join(map(str, ws['company_b']['financial_years']))}")
                st.progress(progress / 100)
                st.caption(f"Progress: {progress}% | Last saved: {ws['last_saved']}")
            with right:
                if st.button("Continue →", key=f"continue_{i}", width="stretch"):
                    st.session_state.active_workspace = i
                    go("workspace_summary")


def create_workspace_page():
    back_button("← Back to My Workspaces", "home")
    st.title("Create Workspace")
    st.caption("Set up the company context before entering financial data.")
    st.markdown("### 1. Analysis Type")
    analysis_type = st.radio("Choose analysis type", ["One Company", "Compare Two Companies"], horizontal=True, label_visibility="collapsed")
    st.markdown("### 2. Company Information")
    company_a = company_card("company_a", "Company A")
    company_b = None
    if analysis_type == "Compare Two Companies":
        company_b = company_card("company_b", "Company B")
    st.markdown("### 3. Analysis Period")
    years = year_options()
    c1, c2 = st.columns(2)
    with c1:
        start_year = st.selectbox("Start Year", years, index=max(0, len(years) - 5))
    with c2:
        end_year = st.selectbox("End Year", years, index=len(years) - 1)
    generated = list(range(start_year, end_year + 1)) if end_year >= start_year else []
    if not generated:
        st.error("End year must be after start year.")
    elif len(generated) > 8:
        st.warning("This is a long period. For undergraduate analysis, 3 to 5 years is usually sufficient.")
    st.markdown("### 4. Confirm Financial Years")
    st.caption("Confirm reported years separately for each company. Untick missing or irregular years.")
    years_a, excluded_a, reason_a = confirm_years("Company A", generated, "company_a")
    years_b, excluded_b, reason_b = [], [], ""
    if analysis_type == "Compare Two Companies":
        st.divider()
        years_b, excluded_b, reason_b = confirm_years("Company B", generated, "company_b")
    st.markdown("### 5. Review")
    if company_a["company_name"]:
        st.markdown("<div class='summary-card'>", unsafe_allow_html=True)
        st.write(f"**Analysis Type:** {analysis_type}")
        st.write(f"**Company A:** {company_a['company_name']} | {company_a['business_activity']} | {company_a['currency']} | {company_a['financial_year_end']}")
        st.write(f"**Company A Years:** {', '.join(map(str, years_a)) if years_a else '-'}")
        if analysis_type == "Compare Two Companies":
            st.write(f"**Company B:** {company_b['company_name']} | {company_b['business_activity']} | {company_b['currency']} | {company_b['financial_year_end']}")
            st.write(f"**Company B Years:** {', '.join(map(str, years_b)) if years_b else '-'}")
        st.markdown("</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Create Workspace", width="stretch"):
            if not company_a["company_name"]:
                st.error("Please enter Company A.")
            elif analysis_type == "Compare Two Companies" and (not company_b or not company_b["company_name"]):
                st.error("Please enter Company B.")
            elif not years_a:
                st.error("Please select at least one financial year for Company A.")
            elif analysis_type == "Compare Two Companies" and not years_b:
                st.error("Please select at least one financial year for Company B.")
            else:
                company_a.update({"financial_years": years_a, "excluded_years": excluded_a, "exclusion_reason": reason_a})
                if company_b:
                    company_b.update({"financial_years": years_b, "excluded_years": excluded_b, "exclusion_reason": reason_b})
                st.session_state.workspaces.append({
                    "analysis_type": analysis_type,
                    "company_a": company_a,
                    "company_b": company_b,
                    "last_saved": datetime.now().strftime("%d %b %Y, %I:%M %p"),
                    "completed_cells": 0,
                    "company_a_completed": False,
                    "company_b_completed": False,
                })
                st.session_state.active_workspace = len(st.session_state.workspaces) - 1
                go("workspace_summary")
    with c2:
        if st.button("Cancel", width="stretch"):
            go("home")


def workflow_progress_block(ws):
    progress, steps = workflow_progress(ws)
    st.subheader("Overall Progress")
    st.progress(progress)
    st.caption(f"{int(progress * 100)}% completed")
    for label, done in steps:
        st.write(f"✅ {label}" if done else f"○ {label}")

def workspace_summary_page():
    idx = st.session_state.active_workspace
    ws = st.session_state.workspaces[idx]
    main, side = st.columns([3, 1])
    with main:
        st.success("✔ Workspace Created Successfully")
        st.title("Workspace Summary")
        st.caption("Please review the company profile before entering financial data.")
        company_summary_card(ws["company_a"], "Company A")
        if ws["analysis_type"] == "Compare Two Companies":
            company_summary_card(ws["company_b"], "Company B")
        st.subheader("Progress")
        c1, c2 = st.columns(2)
        c1.info("Company A: Not Started" if not ws["company_a_completed"] else "Company A: Completed")
        if ws["analysis_type"] == "Compare Two Companies":
            c2.info("Company B: Not Started" if not ws["company_b_completed"] else "Company B: Completed")
        st.divider()
        b1, b2 = st.columns(2)
        with b1:
            if st.button("Edit Company Profile", width="stretch"):
                go("create_workspace")
        with b2:
            if st.button("Continue to Company A →", width="stretch"):
                go("data_entry_a")
    with side:
        notes_panel()


def data_entry_page(company_key, label, next_page):
    idx = st.session_state.active_workspace
    ws = st.session_state.workspaces[idx]
    company = ws[company_key]
    main, side = st.columns([3, 1])
    with main:
        back_button("← Back to Workspace Summary", "workspace_summary")
        st.title(f"{label}: {company['company_name']}")
        st.caption("Financial years are linked from the Company Profile and shown side by side.")
        company_summary_card(company, label)
        table_key = f"{company_key}_data_{idx}"
        if table_key not in st.session_state:
            st.session_state[table_key] = blank_financial_table(company["financial_years"])
        edited = st.data_editor(st.session_state[table_key], width="stretch", hide_index=True, num_rows="fixed", key=f"editor_{company_key}_{idx}")
        st.session_state[table_key] = edited
        st.divider()
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("💾 Save Progress", width="stretch"):
                count = edited.drop(columns=["Financial Statement Item"]).count().sum()
                ws["completed_cells"] += int(count)
                ws["last_saved"] = datetime.now().strftime("%d %b %Y, %I:%M %p")
                st.success("Progress saved in this prototype session.")
        with c2:
            if st.button(f"✅ Complete {label}", width="stretch"):
                ws[f"{company_key}_completed"] = True
                ws["last_saved"] = datetime.now().strftime("%d %b %Y, %I:%M %p")
                go(next_page)
        with c3:
            if st.button("⬅ Workspace Summary", width="stretch"):
                go("workspace_summary")
    with side:
        st.markdown("""
        <div class="sticky-notes">
        <h4>💡 FinSight Assistant</h4>
        <p>Company information and financial years are taken from the Company Profile.</p>
        <hr>
        <p>Later, you can click the 💡 icon beside each item to view glossary guidance.</p>
        </div>
        """, unsafe_allow_html=True)


def company_a_complete_page():
    idx = st.session_state.active_workspace
    ws = st.session_state.workspaces[idx]
    back_button("← Back to Company A Data Entry", "data_entry_a")

    if ws["analysis_type"] == "Compare Two Companies":
        st.success("✅ Company A completed")
        st.title(ws["company_a"]["company_name"])
        st.write(f"{len(ws['company_a']['financial_years'])} financial years entered.")
        workflow_progress_block(ws)
        st.write("Company B data entry is next.")
        if st.button("Continue to Company B →", width="stretch"):
            go("data_entry_b")
    else:
        st.success("🎉 Analysis Ready")
        st.title("Financial Data Completed")
        st.write(f"{ws['company_a']['company_name']} is ready for analysis.")
        st.write(f"{len(ws['company_a']['financial_years'])} financial years entered.")
        workflow_progress_block(ws)
        if st.button("🚀 Open Dashboard", width="stretch"):
            go("dashboard")

def company_b_complete_page():
    idx = st.session_state.active_workspace
    ws = st.session_state.workspaces[idx]
    back_button("← Back to Company B Data Entry", "data_entry_b")

    st.success("🎉 Analysis Ready")
    st.title("Financial Data Completed")
    st.write("Your financial statements have been successfully entered. FinSight is ready to calculate the analysis.")

    workflow_progress_block(ws)
    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Company A")
        st.write(f"✅ {ws['company_a']['company_name']}")
        st.write(f"{len(ws['company_a']['financial_years'])} financial years entered.")
    with c2:
        st.markdown("### Company B")
        st.write(f"✅ {ws['company_b']['company_name']}")
        st.write(f"{len(ws['company_b']['financial_years'])} financial years entered.")

    st.divider()
    st.markdown("### Next Step")
    st.write("FinSight is now ready to calculate:")
    st.write("✔ Financial Ratios")
    st.write("✔ Trend Analysis")
    st.write("✔ Common Size Analysis")
    st.write("✔ Peer Comparison")

    if st.button("🚀 Open Dashboard", width="stretch"):
        go("dashboard")




def score_status(score):
    if score < 40:
        return "Weak"
    if score < 60:
        return "Fair"
    if score < 80:
        return "Good"
    return "Excellent"

def finsight_score_bar(score=82):
    st.markdown("## Overall Financial Quality")
    st.markdown(
        f"""
        <div style="padding:24px;border-radius:18px;border:1px solid #e5e7eb;background:#ffffff;box-shadow:0 4px 16px rgba(0,0,0,0.04);">
            <div style="font-size:62px;font-weight:850;text-align:center;">{score}</div>
            <div style="text-align:center;color:#6b7280;margin-bottom:16px;font-size:16px;">/ 100 &nbsp;•&nbsp; {score_status(score)}</div>
            <div style="height:24px;border-radius:999px;background:linear-gradient(90deg,#ef4444 0%,#ef4444 39%,#f59e0b 40%,#f59e0b 59%,#facc15 60%,#facc15 79%,#22c55e 80%,#22c55e 100%);position:relative;">
                <div style="position:absolute;left:{score}%;top:-7px;transform:translateX(-50%);font-size:28px;">▲</div>
            </div>
            <div style="display:flex;justify-content:space-between;font-size:12px;color:#6b7280;margin-top:8px;">
                <span>Weak</span><span>Fair</span><span>Good</span><span>Excellent</span>
            </div>
            <p style="font-size:13px;color:#6b7280;margin-top:16px;">
                FinSight Score is an educational composite indicator based on ratio categories and diagnostic models.
                It supports analysis but does not replace professional judgement or student evaluation.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

CATEGORY_DATA = {
    "Profitability": {
        "status": "Improving", "trend": "🟢 Improving", "emoji": "📈",
        "flag": "Verify whether the improvement is supported by sustainable operating performance.",
        "ratios": ["Gross Margin", "Operating Margin", "Net Profit Margin", "ROA", "ROE", "ROCE", "EPS"],
        "observation": "Profitability improved between 2024 and 2025.",
        "evidence": "Operating Margin increased while ROA remained relatively stable.",
        "investigate": "Review whether the improvement was driven by higher profit, better asset utilisation, improved cost control, or one-off gains.",
        "annual_report": "Refer to the Management Discussion & Analysis, segment information, and relevant notes to the financial statements.",
        "positive": ["Operating Margin increased in the latest year."],
        "review": ["ROA remained relatively stable; review asset utilisation."],
        "concern": ["Check whether profit growth is recurring or supported by one-off items."],
    },
    "Liquidity": {
        "status": "Mixed Trend", "trend": "🟡 Mixed Trend", "emoji": "💧",
        "flag": "Review the composition of current assets and current liabilities.",
        "ratios": ["Current Ratio", "Quick Ratio", "Cash Ratio", "Working Capital", "Operating Cash Flow Ratio"],
        "observation": "Liquidity shows a mixed trend.",
        "evidence": "Current Ratio improved, while supporting liquidity indicators did not move uniformly.",
        "investigate": "Review whether the change was driven by cash, receivables, inventory, or lower current liabilities.",
        "annual_report": "Refer to current asset composition, working capital disclosures, and liquidity risk notes.",
        "positive": ["Working capital remains positive in this prototype."],
        "review": ["Identify the main component causing the Current Ratio to increase."],
        "concern": ["Receivables and inventories may weaken liquidity quality if they increased excessively."],
    },
    "Efficiency": {
        "status": "Requires Attention", "trend": "🔴 Requires Attention", "emoji": "⚙️",
        "flag": "Assess whether assets are being utilised efficiently over time.",
        "ratios": ["Asset Turnover", "Inventory Turnover", "Receivable Turnover", "PPE Turnover", "Days Inventory", "Collection Period"],
        "observation": "Efficiency requires further attention.",
        "evidence": "Asset Turnover remained stable while inventory-related indicators weakened.",
        "investigate": "Review whether inventory or receivables increased faster than revenue and whether assets are generating sufficient sales.",
        "annual_report": "Refer to the business review, inventory notes, receivable ageing, and PPE disclosures.",
        "positive": ["Receivable Turnover improved in this prototype."],
        "review": ["Stable Asset Turnover may indicate limited efficiency improvement."],
        "concern": ["Inventory movement may indicate slower turnover or excess stock."],
    },
    "Solvency": {
        "status": "Improving", "trend": "🟢 Improving", "emoji": "🏦",
        "flag": "Review changes in financing structure and debt management.",
        "ratios": ["Debt-to-Equity", "Debt Ratio", "Equity Ratio", "Interest Coverage", "Financial Leverage"],
        "observation": "Solvency generally improved.",
        "evidence": "Debt-to-Equity declined while Interest Coverage increased.",
        "investigate": "Review whether the company reduced borrowings, improved EBIT, or strengthened equity during the period.",
        "annual_report": "Refer to borrowings, finance costs, maturity profiles, and capital management disclosures.",
        "positive": ["Leverage declined and Interest Coverage improved."],
        "review": ["Determine whether debt reduction was supported by stronger operating performance."],
        "concern": ["Assess whether lower borrowing affected expansion or growth capacity."],
    },
    "Market Performance": {
        "status": "Improving", "trend": "🟢 Improving", "emoji": "📊",
        "flag": "Assess whether market indicators are supported by fundamental performance.",
        "ratios": ["EPS", "Dividend Yield", "Dividend Payout", "PE Ratio", "Book Value per Share"],
        "observation": "Market performance generally improved.",
        "evidence": "EPS increased while dividend indicators remained relatively stable.",
        "investigate": "Review whether EPS growth is supported by recurring profit and whether share-price movement is consistent with financial performance.",
        "annual_report": "Refer to EPS, dividend disclosures, share capital notes, and market information.",
        "positive": ["EPS increased in the latest year."],
        "review": ["Review whether EPS growth came from profit growth or changes in shares outstanding."],
        "concern": ["Market ratios may be distorted when share prices move independently of fundamentals."],
    },
}

DIAGNOSTIC_DATA = {
    "Altman Z Score": {
        "status": "Safe Zone", "trend": "🟢 Safe Zone", "emoji": "🟢", "score": "3.62",
        "flag": "Review the five-year trend together with solvency indicators.",
        "ratios": ["Altman Z Score", "Working Capital / Total Assets", "Retained Earnings / Total Assets", "EBIT / Total Assets", "Equity / Liabilities", "Sales / Total Assets"],
        "observation": "Altman Z Score indicates a safe zone in this prototype.",
        "evidence": "The latest score is above the commonly used distress threshold.",
        "investigate": "Review whether the score is supported by liquidity, profitability, retained earnings, and solvency trends.",
        "annual_report": "Refer to going-concern, borrowings, liquidity risk, and capital management disclosures.",
        "positive": ["Latest score is in the safe zone."],
        "review": ["Review the five-year trend rather than relying only on the latest year."],
        "concern": ["A safe score does not remove the need to review other financial and non-financial risks."],
    },
    "Piotroski F Score": {
        "status": "Strong Signals", "trend": "🟢 Strong Signals", "emoji": "🔵", "score": "8 / 9",
        "flag": "Review the individual components supporting the total score.",
        "ratios": ["Piotroski F Score", "Profitability Signals", "Leverage / Liquidity Signals", "Operating Efficiency Signals"],
        "observation": "Piotroski F Score indicates strong financial signals in this prototype.",
        "evidence": "Most profitability, liquidity, leverage, and efficiency signals are positive.",
        "investigate": "Review which components contributed positively and whether the improvement is consistent across the analysis period.",
        "annual_report": "Refer to profitability, operating cash flow, leverage, liquidity, and share-issuance disclosures.",
        "positive": ["The total score suggests strong financial signals."],
        "review": ["Review each component rather than relying only on the total score."],
        "concern": ["A strong score may still conceal weaknesses in an individual area."],
    },
    "Beneish M Score": {
        "status": "Low Manipulation Signal", "trend": "🟢 Low Manipulation Signal", "emoji": "🟣", "score": "-2.85",
        "flag": "Examine earnings quality and unusual financial-statement movements.",
        "ratios": ["Beneish M Score", "DSRI", "GMI", "AQI", "SGI", "DEPI", "SGAI", "TATA", "LVGI"],
        "observation": "Beneish M Score indicates a low manipulation signal in this prototype.",
        "evidence": "The latest score is below the commonly used manipulation threshold.",
        "investigate": "Review accruals, revenue growth, asset quality, depreciation, leverage, and unusual accounting estimates.",
        "annual_report": "Refer to revenue recognition, receivable ageing, accounting estimates, non-current assets, and accrual-related disclosures.",
        "positive": ["The score indicates a low manipulation signal in this prototype."],
        "review": ["Review accruals and revenue quality before concluding."],
        "concern": ["A low signal does not prove that earnings are free from manipulation or estimation risk."],
    },
}

def _divider_html():
    return "<div style='height:3px;border-radius:999px;margin:14px 0;background:linear-gradient(90deg,#2563eb 0%,#60a5fa 55%,rgba(96,165,250,0.08) 100%);'></div>"

def category_card(title, data, key=None, selected=False):
    border = "2px solid #2563eb" if selected else "1px solid #e5e7eb"
    bg = "#eff6ff" if selected else "#ffffff"
    encoded = title.replace(" ", "%20")
    st.markdown(f"""
        <div style="padding:20px;border-radius:18px;border:{border};background:{bg};box-shadow:0 5px 18px rgba(0,0,0,0.05);min-height:305px;">
            <div style="font-size:24px;font-weight:800;">{data['emoji']} {title}</div>
            {_divider_html()}
            <div style="font-size:14px;color:#6b7280;font-weight:700;">Overall Trend</div>
            <div style="font-size:18px;font-weight:800;margin-top:6px;">{data['trend']}</div>
            {_divider_html()}
            <div style="font-size:14px;color:#6b7280;font-weight:700;">🚩 Key Flag</div>
            <div style="font-size:15px;line-height:1.55;margin-top:7px;min-height:72px;">{data['flag']}</div>
            {_divider_html()}
            <a href="?analysis={encoded}" style="color:#1d4ed8;font-weight:800;font-size:16px;text-decoration:underline;text-decoration-thickness:2px;text-underline-offset:4px;">View Analysis →</a>
        </div>""", unsafe_allow_html=True)

def diagnostic_model_card(title, data, key=None, selected=False):
    border = "2px solid #2563eb" if selected else "1px solid #e5e7eb"
    bg = "#eff6ff" if selected else "#ffffff"
    encoded = title.replace(" ", "%20")
    st.markdown(f"""
        <div style="padding:20px;border-radius:18px;border:{border};background:{bg};box-shadow:0 5px 18px rgba(0,0,0,0.05);min-height:345px;">
            <div style="font-size:24px;font-weight:800;">{data['emoji']} {title}</div>
            <div style="font-size:34px;font-weight:850;margin-top:8px;">{data['score']}</div>
            {_divider_html()}
            <div style="font-size:14px;color:#6b7280;font-weight:700;">Overall Trend</div>
            <div style="font-size:18px;font-weight:800;margin-top:6px;">{data['trend']}</div>
            {_divider_html()}
            <div style="font-size:14px;color:#6b7280;font-weight:700;">🚩 Key Flag</div>
            <div style="font-size:15px;line-height:1.55;margin-top:7px;min-height:72px;">{data['flag']}</div>
            {_divider_html()}
            <a href="?analysis={encoded}" style="color:#1d4ed8;font-weight:800;font-size:16px;text-decoration:underline;text-decoration-thickness:2px;text-underline-offset:4px;">View Analysis →</a>
        </div>""", unsafe_allow_html=True)

def dummy_trend_data(item, company_a_name="Company A", company_b_name=None):
    base = {"ROE":[11.2,12.8,13.5,14.1,15.4],"ROA":[7.5,8.0,8.6,8.7,8.3],"Gross Margin":[31,31.6,32.1,33,34.2],"Operating Margin":[14.8,15.5,16.2,17.4,18.6],"Net Profit Margin":[8.4,8.9,9.3,9.8,10.5],"EPS":[1.22,1.34,1.48,1.62,1.84],"Current Ratio":[1.42,1.55,1.64,1.74,1.92],"Quick Ratio":[.82,.88,.91,.93,.95],"Cash Ratio":[.22,.25,.28,.26,.30],"Working Capital":[120,140,155,175,210],"Operating Cash Flow Ratio":[.38,.41,.46,.43,.49],"Asset Turnover":[1.18,1.21,1.24,1.24,1.24],"Inventory Turnover":[5.4,5.1,4.9,4.7,4.5],"Receivable Turnover":[6.2,6.5,6.8,7,7.2],"PPE Turnover":[2.1,2.2,2.2,2.3,2.4],"Days Inventory":[67,72,75,78,81],"Collection Period":[59,56,54,52,51],"Debt-to-Equity":[.62,.58,.53,.51,.45],"Debt Ratio":[.42,.40,.38,.36,.34],"Equity Ratio":[.58,.60,.62,.64,.66],"Interest Coverage":[4.9,5.4,5.8,5.9,6.8],"Financial Leverage":[1.72,1.67,1.61,1.56,1.52],"Dividend Yield":[3.2,3.1,3,3.3,3.4],"Dividend Payout":[48,49,50,50,51],"PE Ratio":[18.5,19.1,20.3,21.2,22.1],"Book Value per Share":[6.1,6.4,6.8,7.1,7.5],"Altman Z Score":[2.85,3.02,3.18,3.41,3.62],"Working Capital / Total Assets":[.12,.13,.14,.15,.16],"Retained Earnings / Total Assets":[.22,.23,.24,.25,.26],"EBIT / Total Assets":[.09,.10,.10,.11,.12],"Equity / Liabilities":[1.38,1.46,1.55,1.65,1.78],"Sales / Total Assets":[1.18,1.21,1.24,1.24,1.24],"Piotroski F Score":[6,7,7,8,8],"Profitability Signals":[3,3,3,4,4],"Leverage / Liquidity Signals":[2,2,2,2,2],"Operating Efficiency Signals":[1,2,2,2,2],"Beneish M Score":[-2.20,-2.35,-2.51,-2.68,-2.85],"DSRI":[1.02,1.01,.98,.96,.95],"GMI":[1.01,.99,.98,.97,.96],"AQI":[1.05,1.03,1.02,1.01,1.00],"SGI":[1.08,1.07,1.05,1.04,1.03],"DEPI":[.99,.98,.97,.97,.96],"SGAI":[1.02,1.01,1,.99,.99],"TATA":[.04,.03,.03,.02,.02],"LVGI":[1.01,1,.98,.97,.96]}
    years = ["2022","2023","2024","2025","2026"]
    vals = base.get(item, [10,11,12,13,14])
    data = pd.DataFrame({"Year": years, company_a_name: vals})
    if company_b_name:
        data[company_b_name] = [round(v * .90 + (i * .25), 2) for i, v in enumerate(vals)]
    return data


def analysis_workspace(ws):
    query_value = st.query_params.get("analysis")
    if isinstance(query_value, list):
        query_value = query_value[0] if query_value else None
    if query_value:
        st.session_state.selected_dashboard_category = query_value
    selected = st.session_state.get("selected_dashboard_category")

    st.markdown("## 📊 Detail Analysis Workspace")
    if not selected:
        st.info("Select any Financial Ratio Category or Financial Diagnostic Model above to begin. FinSight will display trend analysis, C3-level guidance, and financial flags.")
        st.markdown("""
        <div style="padding:30px;border-radius:18px;border:1px dashed #cbd5e1;background:#f8fafc;text-align:center;">
            <div style="font-size:46px;">☝️</div><h3>Select a category or diagnostic model above</h3>
            <p style="color:#6b7280;">FinSight organises the evidence. Students remain responsible for interpretation, evaluation and conclusion.</p>
        </div>""", unsafe_allow_html=True)
        return

    data = CATEGORY_DATA.get(selected) or DIAGNOSTIC_DATA.get(selected)
    if not data:
        st.warning("Please select a valid category or diagnostic model.")
        return

    company_a = ws["company_a"]["company_name"]
    company_b = ws["company_b"]["company_name"] if ws["analysis_type"] == "Compare Two Companies" else None

    profitability_a = profitability_output(ws, "company_a") if selected == "Profitability" else None
    profitability_b = profitability_output(ws, "company_b") if selected == "Profitability" and ws["analysis_type"] == "Compare Two Companies" else None
    if selected == "Profitability":
        live_trend, live_flag = profitability_signal(profitability_a)
        live_guidance = profitability_guidance(profitability_a)
        data = dict(data)
        data["trend"] = live_trend
        data["flag"] = live_flag
        data.update(live_guidance)

    st.markdown(f"### {data['emoji']} {selected} Analysis")
    st.caption(f"Overall Trend: {data['trend']}")

    if ws["analysis_type"] == "Compare Two Companies":
        chart_mode = st.radio("Display Mode", ["Combined Comparison", "Separate Company View"], horizontal=True, key=f"chart_mode_{selected}")
    else:
        chart_mode = "Single Company"

    component_key = f"component_select_{selected}"
    if component_key not in st.session_state:
        st.session_state[component_key] = data["ratios"][0]
    selected_ratio = st.selectbox("Component", data["ratios"], key=component_key)

    st.markdown("### Trend Analysis")

    if selected == "Profitability":
        chart_a = profitability_a[["Year", selected_ratio]].rename(columns={selected_ratio: company_a})
        chart_b = None
        if profitability_b is not None:
            chart_b = profitability_b[["Year", selected_ratio]].rename(columns={selected_ratio: company_b})

        if ws["analysis_type"] == "Compare Two Companies" and chart_mode == "Combined Comparison":
            combined = chart_a.merge(chart_b, on="Year", how="outer")
            st.caption(f"{company_a} and {company_b}")
            st.line_chart(combined.set_index("Year"))
        elif ws["analysis_type"] == "Compare Two Companies" and chart_mode == "Separate Company View":
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown(f"#### {company_a}")
                st.line_chart(chart_a.set_index("Year"))
            with col_b:
                st.markdown(f"#### {company_b}")
                st.line_chart(chart_b.set_index("Year"))
        else:
            st.caption(company_a)
            st.line_chart(chart_a.set_index("Year"))

        with st.expander("View Profitability Formula & Methodology"):
            st.write(f"**{selected_ratio}:** {PROFITABILITY_FORMULAS[selected_ratio]}")
            st.caption(
                "Average balances are used where available. For the first financial year, closing balances are used because prior-year balances are unavailable. "
                "ROCE uses Capital Employed = Total Assets − Current Liabilities. EPS uses Net Income ÷ average Shares Outstanding in Version 1.0."
            )

        with st.expander("View Calculated Profitability Table"):
            display_a = profitability_a.copy()
            for col in display_a.columns:
                if col != "Year":
                    display_a[col] = display_a[col].map(lambda x, c=col: format_ratio_value(c, x))
            st.markdown(f"#### {company_a}")
            st.dataframe(display_a, hide_index=True, width="stretch")
            if profitability_b is not None:
                display_b = profitability_b.copy()
                for col in display_b.columns:
                    if col != "Year":
                        display_b[col] = display_b[col].map(lambda x, c=col: format_ratio_value(c, x))
                st.markdown(f"#### {company_b}")
                st.dataframe(display_b, hide_index=True, width="stretch")
    else:
        if ws["analysis_type"] == "Compare Two Companies" and chart_mode == "Combined Comparison":
            trend_df = dummy_trend_data(selected_ratio, company_a, company_b)
            st.caption(f"{company_a} and {company_b}")
            st.line_chart(trend_df.set_index("Year"))
        elif ws["analysis_type"] == "Compare Two Companies" and chart_mode == "Separate Company View":
            trend_df_a = dummy_trend_data(selected_ratio, company_a, None)
            trend_df_b = dummy_trend_data(selected_ratio, company_b, None)
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown(f"#### {company_a}")
                st.line_chart(trend_df_a.set_index("Year"))
            with col_b:
                st.markdown(f"#### {company_b}")
                st.line_chart(trend_df_b.set_index("Year"))
        else:
            trend_df = dummy_trend_data(selected_ratio, company_a, None)
            st.caption(company_a)
            st.line_chart(trend_df.set_index("Year"))

    st.markdown("### Supporting Indicators")
    st.caption("Use the Component dropdown above to switch between supporting ratios or model components.")
    cols = st.columns(3)
    for i, ratio in enumerate(data["ratios"]):
        with cols[i % 3]:
            st.markdown(f"<div style='padding:10px 12px;margin-bottom:8px;border-radius:10px;border:1px solid #e5e7eb;background:#f8fafc;font-size:14px;text-align:center;'>{ratio}</div>", unsafe_allow_html=True)

    c1, c2 = st.columns([1.15, 1])
    with c1:
        st.markdown("### 💡 FinSight Assistant")
        st.markdown(f"""
        <div style="padding:18px;border-radius:16px;border:1px solid #bfdbfe;background:#eff6ff;">
            <div style="font-weight:800;color:#1e3a8a;">Observation</div><p>{data['observation']}</p>
            <div style="font-weight:800;color:#1e3a8a;">Evidence</div><p>{data['evidence']}</p>
            <div style="font-weight:800;color:#1e3a8a;">What to Investigate</div><p>{data['investigate']}</p>
            <div style="font-weight:800;color:#1e3a8a;">Annual Report Reminder</div><p>{data['annual_report']}</p>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("### 🚩 Financial Flags")
        st.markdown("#### 🟢 Positive Signals")
        for item in data["positive"]: st.success(item)
        st.markdown("#### 🟡 Review Required")
        for item in data["review"]: st.warning(item)
        st.markdown("#### 🔴 Potential Concerns")
        for item in data["concern"]: st.error(item)

def dashboard_page():
    idx = st.session_state.active_workspace
    ws = st.session_state.workspaces[idx]
    company_a = ws["company_a"]
    company_b = ws.get("company_b")
    if ws["analysis_type"] == "Compare Two Companies":
        back_button("← Back to Analysis Ready", "company_b_complete")
    else:
        back_button("← Back to Analysis Ready", "company_a_complete")
    st.title("📊 FinSight Dashboard")
    st.caption("Sprint 4.1: Profitability ratios are calculated from entered financial data. Other categories remain placeholders.")
    st.markdown("## 1. Company Overview")
    with st.container(border=True):
        if ws["analysis_type"] == "Compare Two Companies":
            st.markdown(f"### {company_a['company_name']} vs {company_b['company_name']}")
            st.write(f"**Company A:** {company_a['business_activity']} | {company_a['currency']} | {company_a['financial_year_end']}")
            st.write(f"**Company B:** {company_b['business_activity']} | {company_b['currency']} | {company_b['financial_year_end']}")
            st.write(f"**Company A Years:** {' • '.join(map(str, company_a['financial_years']))}")
            st.write(f"**Company B Years:** {' • '.join(map(str, company_b['financial_years']))}")
            if company_a["currency"] != company_b["currency"]:
                st.info("📘 Academic Note: The selected companies use different reporting currencies. Financial ratio comparisons remain valid, but absolute values such as revenue, assets, liabilities and profit should not be compared directly unless converted into a common currency.")
        else:
            st.markdown(f"### {company_a['company_name']}")
            st.write(f"**Primary Business Activity:** {company_a['business_activity']}")
            st.write(f"**Currency:** {company_a['currency']}")
            st.write(f"**Financial Year-End:** {company_a['financial_year_end']}")
            st.write(f"**Financial Years:** {' • '.join(map(str, company_a['financial_years']))}")
    st.divider()
    st.markdown("## 2. FinSight Score")
    left, right = st.columns([1.2, 1])
    with left:
        finsight_score_bar(82)
    with right:
        st.markdown("### Score Breakdown")
        st.write("📈 Profitability: 13 / 15")
        st.write("💧 Liquidity: 8 / 10")
        st.write("⚙️ Efficiency: 7 / 10")
        st.write("🏦 Solvency: 9 / 10")
        st.write("📊 Market Performance: 4 / 5")
        st.write("🟢 Altman Z: 13 / 15")
        st.write("🔵 Piotroski F: 12 / 15")
        st.write("🟣 Beneish M: 17 / 20")
    st.divider()
    st.markdown("## 3. Financial Ratio Categories")
    st.caption("Review the Overall Trend and Key Flag, then select View Analysis for details.")
    selected = st.session_state.get("selected_dashboard_category")
    live_profitability = dict(CATEGORY_DATA["Profitability"])
    profit_result = profitability_output(ws, "company_a")
    live_profitability["trend"], live_profitability["flag"] = profitability_signal(profit_result)
    live_profitability["status"] = live_profitability["trend"].replace("🟢 ", "").replace("🟡 ", "").replace("🔴 ", "").replace("⚪ ", "")

    c1, c2, c3 = st.columns(3)
    with c1:
        category_card("Profitability", live_profitability, "cat_profitability", selected == "Profitability")
    with c2:
        category_card("Liquidity", CATEGORY_DATA["Liquidity"], "cat_liquidity", selected == "Liquidity")
    with c3:
        category_card("Efficiency", CATEGORY_DATA["Efficiency"], "cat_efficiency", selected == "Efficiency")
    c4, c5 = st.columns(2)
    with c4:
        category_card("Solvency", CATEGORY_DATA["Solvency"], "cat_solvency", selected == "Solvency")
    with c5:
        category_card("Market Performance", CATEGORY_DATA["Market Performance"], "cat_market", selected == "Market Performance")
    st.divider()
    st.markdown("## 4. Financial Diagnostic Models")
    st.caption("Review the Overall Trend and Key Flag, then select View Analysis for details.")
    d1, d2, d3 = st.columns(3)
    with d1:
        diagnostic_model_card("Altman Z Score", DIAGNOSTIC_DATA["Altman Z Score"], "diag_altman", selected == "Altman Z Score")
    with d2:
        diagnostic_model_card("Piotroski F Score", DIAGNOSTIC_DATA["Piotroski F Score"], "diag_piotroski", selected == "Piotroski F Score")
    with d3:
        diagnostic_model_card("Beneish M Score", DIAGNOSTIC_DATA["Beneish M Score"], "diag_beneish", selected == "Beneish M Score")
    st.divider()
    analysis_workspace(ws)



def about_finsight_page():
    st.title("ℹ️ About FinSight")
    st.caption("Platform information, methodology, diagnostic models, limitations and version information.")

    with st.expander("About FinSight", expanded=True):
        st.write(
            "FinSight is an integrated financial statement analysis platform designed to assist users "
            "in evaluating corporate financial performance through accounting ratios, financial diagnostic models, "
            "trend analysis and interactive analytical visualisations."
        )
        st.write(
            "The platform provides a structured analytical environment to support informed financial analysis "
            "and decision-making."
        )

    with st.expander("Methodology"):
        st.markdown(
            """
            FinSight calculates financial ratios using a consistent methodology based on the financial information entered by the user.

            - Ratios requiring average balances use the average of opening and closing balances where both balances are available.
            - For the first financial year entered into the workspace, closing balances are used because prior-year balances are unavailable.
            - All calculations follow the methodology adopted by the active FinSight calculation engine.
            - Piotroski F-Score and Beneish M-Score require prior-year data; therefore, the first financial year is displayed as N/A.
            - Altman Z-Score Version 1.0 applies the original public-manufacturing-company model.
            - EPS uses Net Income as a proxy for profit attributable to ordinary shareholders unless a more specific figure is entered in a future version.
            """
        )

    with st.expander("Diagnostic Models"):
        st.markdown(
            """
            FinSight incorporates recognised financial diagnostic models to complement traditional ratio analysis:

            - **Altman Z-Score**
            - **Piotroski F-Score**
            - **Beneish M-Score**

            These models provide additional analytical indicators and should be interpreted together with financial ratios and qualitative information.
            """
        )

    st.markdown(
        """
        <div style="
            height:3px;
            border-radius:999px;
            margin:22px 0;
            background:linear-gradient(90deg,#2563eb 0%,#60a5fa 55%,rgba(96,165,250,0.08) 100%);
        "></div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("## 🚀 Explore FinSight")
    st.write("Experience the complete FinSight workflow using sample financial data.")

    if st.button("Open Interactive Demo →", key="open_interactive_demo", width="stretch"):
        go("interactive_demo")

    st.markdown(
        """
        <div style="
            height:3px;
            border-radius:999px;
            margin:22px 0;
            background:linear-gradient(90deg,#2563eb 0%,#60a5fa 55%,rgba(96,165,250,0.08) 100%);
        "></div>
        """,
        unsafe_allow_html=True,
    )

    with st.expander("Limitations"):
        st.markdown(
            """
            Users should consider the following when interpreting FinSight outputs:

            - Results depend on the accuracy and completeness of the financial information entered.
            - Financial ratios should be interpreted together with qualitative disclosures in the annual report.
            - Industry characteristics and accounting-policy differences may affect comparability.
            - One-off transactions and unusual events may influence reported financial performance.
            - Financial diagnostic models provide analytical indicators and should not be interpreted as definitive conclusions.
            - FinSight supports financial analysis and informed decision-making but does not replace professional judgement.
            """
        )

    with st.expander("Version Information"):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Application Version", "1.0")
        with c2:
            st.metric("Release Date", "July 2026")
        with c3:
            st.metric("Last Updated", "11 July 2026")

    st.divider()
    st.caption("FinSight")



# -----------------------------
# Official FinSight sample dataset
# -----------------------------
DEMO_YEARS = [2022, 2023, 2024, 2025, 2026]

DEMO_COMPANY = {
    "company_name": "FinSight Sample Manufacturing Berhad",
    "business_activity": "Manufacturing",
    "currency": "RM (MYR)",
    "financial_year_end": "31 December",
    "financial_years": DEMO_YEARS,
}

DEMO_VALUES = {
    "Sales (Revenue)": [520.00, 565.00, 610.00, 670.00, 735.00],
    "COGS": [353.60, 378.55, 402.60, 435.50, 470.40],
    "EBIT": [62.40, 73.45, 85.40, 103.85, 124.95],
    "Net Income": [41.60, 49.16, 57.95, 70.35, 88.20],
    "Interest Expense": [12.00, 11.50, 11.00, 10.50, 10.00],
    "Current Assets": [250.00, 280.00, 320.00, 375.00, 440.00],
    "Inventory": [62.00, 78.00, 100.00, 130.00, 170.00],
    "Accounts Receivable": [58.00, 70.00, 86.00, 108.00, 135.00],
    "Current Liabilities": [155.00, 170.00, 190.00, 215.00, 245.00],
    "Total Assets": [580.00, 625.00, 680.00, 745.00, 820.00],
    "Total Liabilities": [220.00, 235.00, 255.00, 280.00, 310.00],
    "Long-term Liabilities": [65.00, 65.00, 65.00, 65.00, 65.00],
    "Total Equity": [360.00, 390.00, 425.00, 465.00, 510.00],
    "Retained Earnings": [120.00, 145.00, 175.00, 215.00, 270.00],
    "Cash Equivalents": [82.00, 80.00, 74.00, 68.00, 61.00],
    "Marketable Securities": [12.00, 13.00, 14.00, 15.00, 16.00],
    "PPE (Net)": [330.00, 345.00, 360.00, 370.00, 380.00],
    "Shares Outstanding": [72.00, 73.00, 74.00, 75.00, 76.00],
    "Dividends per Share": [0.25, 0.28, 0.32, 0.36, 0.42],
    "Share Price": [8.50, 9.10, 10.20, 11.40, 13.20],
    "Operating Cash Flow": [49.00, 56.00, 65.00, 77.00, 95.00],
    "Depreciation Expense": [26.00, 27.00, 28.00, 29.00, 30.00],
    "SG&A Expense": [48.00, 50.00, 52.00, 55.00, 58.00],
}


def demo_financial_table():
    rows = []
    for item in blank_financial_table(DEMO_YEARS)["Financial Statement Item"].tolist():
        row = {"Financial Statement Item": item}
        values = DEMO_VALUES.get(item, [None] * len(DEMO_YEARS))
        for year, value in zip(DEMO_YEARS, values):
            row[str(year)] = value
        rows.append(row)
    return pd.DataFrame(rows)


def demo_profitability_result():
    return calculate_profitability_ratios(demo_financial_table(), DEMO_YEARS)


def demo_category_link(title, data, selected=False):
    border = "2px solid #2563eb" if selected else "1px solid #e5e7eb"
    bg = "#eff6ff" if selected else "#ffffff"
    encoded = title.replace(" ", "%20")
    st.markdown(
        f"""
        <div style="
            padding:20px;
            border-radius:18px;
            border:{border};
            background:{bg};
            box-shadow:0 5px 18px rgba(0,0,0,0.05);
            min-height:305px;
        ">
            <div style="font-size:24px;font-weight:800;">{data['emoji']} {title}</div>
            {_divider_html()}
            <div style="font-size:14px;color:#6b7280;font-weight:700;">Overall Trend</div>
            <div style="font-size:18px;font-weight:800;margin-top:6px;">{data['trend']}</div>
            {_divider_html()}
            <div style="font-size:14px;color:#6b7280;font-weight:700;">🚩 Key Flag</div>
            <div style="font-size:15px;line-height:1.55;margin-top:7px;min-height:72px;">{data['flag']}</div>
            {_divider_html()}
            <a href="?page=interactive_demo&demo_analysis={encoded}" style="
                color:#1d4ed8;
                font-weight:800;
                font-size:16px;
                text-decoration:underline;
                text-decoration-thickness:2px;
                text-underline-offset:4px;
            ">View Analysis →</a>
        </div>
        """,
        unsafe_allow_html=True,
    )


def demo_diagnostic_link(title, data, selected=False):
    border = "2px solid #2563eb" if selected else "1px solid #e5e7eb"
    bg = "#eff6ff" if selected else "#ffffff"
    encoded = title.replace(" ", "%20")
    st.markdown(
        f"""
        <div style="
            padding:20px;
            border-radius:18px;
            border:{border};
            background:{bg};
            box-shadow:0 5px 18px rgba(0,0,0,0.05);
            min-height:340px;
        ">
            <div style="font-size:24px;font-weight:800;">{data['emoji']} {title}</div>
            <div style="font-size:34px;font-weight:850;margin-top:8px;">{data['score']}</div>
            {_divider_html()}
            <div style="font-size:14px;color:#6b7280;font-weight:700;">Overall Trend</div>
            <div style="font-size:18px;font-weight:800;margin-top:6px;">{data['trend']}</div>
            {_divider_html()}
            <div style="font-size:14px;color:#6b7280;font-weight:700;">🚩 Key Flag</div>
            <div style="font-size:15px;line-height:1.55;margin-top:7px;min-height:72px;">{data['flag']}</div>
            {_divider_html()}
            <a href="?page=interactive_demo&demo_analysis={encoded}" style="
                color:#1d4ed8;
                font-weight:800;
                font-size:16px;
                text-decoration:underline;
                text-decoration-thickness:2px;
                text-underline-offset:4px;
            ">View Analysis →</a>
        </div>
        """,
        unsafe_allow_html=True,
    )


def demo_analysis_workspace(selected):
    st.markdown(
        '<div id="demo-analysis-anchor"></div>',
        unsafe_allow_html=True,
    )
    st.markdown("## 📊 Detail Analysis Workspace")

    if selected and st.session_state.pop("demo_scroll_pending", False):
        components.html(
            """
            <script>
            setTimeout(function() {
                const anchor = window.parent.document.getElementById("demo-analysis-anchor");
                if (anchor) {
                    anchor.scrollIntoView({behavior: "smooth", block: "start"});
                }
            }, 350);
            </script>
            """,
            height=0,
        )

    if not selected:
        st.info(
            "Select any Financial Ratio Category or Financial Diagnostic Model above to begin. "
            "FinSight will display trend analysis, guidance and financial flags using the official sample dataset."
        )
        st.markdown(
            """
            <div style="padding:30px;border-radius:18px;border:1px dashed #cbd5e1;background:#f8fafc;text-align:center;">
                <div style="font-size:46px;">☝️</div>
                <h3>Select a category or diagnostic model above</h3>
                <p style="color:#6b7280;">
                    This demonstration shows how FinSight organises evidence before users form their own conclusions.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    data = CATEGORY_DATA.get(selected) or DIAGNOSTIC_DATA.get(selected)
    if not data:
        st.warning("Please select a valid category or diagnostic model.")
        return

    profitability = demo_profitability_result() if selected == "Profitability" else None

    if selected == "Profitability":
        data = dict(data)
        live_trend, live_flag = profitability_signal(profitability)
        data["trend"] = live_trend
        data["flag"] = live_flag
        data.update(profitability_guidance(profitability))

    st.markdown(f"### {data['emoji']} {selected} Analysis")
    st.caption(f"Overall Trend: {data['trend']}")

    component_key = f"demo_component_{selected}"
    if component_key not in st.session_state:
        st.session_state[component_key] = data["ratios"][0]

    selected_ratio = st.selectbox(
        "Component",
        data["ratios"],
        key=component_key,
    )

    st.markdown("### Trend Analysis")

    if selected == "Profitability":
        chart = profitability[["Year", selected_ratio]].rename(
            columns={selected_ratio: DEMO_COMPANY["company_name"]}
        )
        st.caption(DEMO_COMPANY["company_name"])
        st.line_chart(chart.set_index("Year"))

        with st.expander("View Profitability Formula & Methodology"):
            st.write(f"**{selected_ratio}:** {PROFITABILITY_FORMULAS[selected_ratio]}")
            st.caption(
                "Average balances are used where available. For the first financial year, closing balances are used "
                "because prior-year balances are unavailable. ROCE uses Capital Employed = Total Assets − Current Liabilities. "
                "EPS uses Net Income ÷ average Shares Outstanding in Version 1.0."
            )

        with st.expander("View Calculated Profitability Table"):
            display = profitability.copy()
            for col in display.columns:
                if col != "Year":
                    display[col] = display[col].map(
                        lambda x, c=col: format_ratio_value(c, x)
                    )
            st.dataframe(display, hide_index=True, width="stretch")
    else:
        trend_df = dummy_trend_data(
            selected_ratio,
            DEMO_COMPANY["company_name"],
            None,
        )
        st.caption(DEMO_COMPANY["company_name"])
        st.line_chart(trend_df.set_index("Year"))

    st.markdown("### Supporting Indicators")
    st.caption("Use the Component dropdown above to switch between supporting ratios or model components.")
    cols = st.columns(3)
    for i, ratio in enumerate(data["ratios"]):
        with cols[i % 3]:
            st.markdown(
                f"""
                <div style="
                    padding:10px 12px;
                    margin-bottom:8px;
                    border-radius:10px;
                    border:1px solid #e5e7eb;
                    background:#f8fafc;
                    font-size:14px;
                    text-align:center;
                ">{ratio}</div>
                """,
                unsafe_allow_html=True,
            )

    c1, c2 = st.columns([1.15, 1])
    with c1:
        st.markdown("### 💡 FinSight Assistant")
        st.markdown(
            f"""
            <div style="padding:18px;border-radius:16px;border:1px solid #bfdbfe;background:#eff6ff;">
                <div style="font-weight:800;color:#1e3a8a;">Observation</div>
                <p>{data['observation']}</p>
                <div style="font-weight:800;color:#1e3a8a;">Evidence</div>
                <p>{data['evidence']}</p>
                <div style="font-weight:800;color:#1e3a8a;">What to Investigate</div>
                <p>{data['investigate']}</p>
                <div style="font-weight:800;color:#1e3a8a;">Annual Report Reminder</div>
                <p>{data['annual_report']}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown("### 🚩 Financial Flags")
        st.markdown("#### 🟢 Positive Signals")
        for item in data["positive"]:
            st.success(item)
        st.markdown("#### 🟡 Review Required")
        for item in data["review"]:
            st.warning(item)
        st.markdown("#### 🔴 Potential Concerns")
        for item in data["concern"]:
            st.error(item)



def interactive_demo_page():
    if st.button("← Back to About FinSight", width="content"):
        st.query_params.clear()
        st.session_state.pop("demo_selected_category", None)
        st.session_state.pop("demo_scroll_pending", None)
        go("about_finsight")

    st.title("🚀 Explore FinSight")

    st.markdown(
        """
        <div style="
            padding:20px 22px;
            border-radius:16px;
            border:1px solid #bfdbfe;
            background:#eff6ff;
            margin-bottom:18px;
        ">
            <div style="font-size:19px;font-weight:800;color:#1e3a8a;">
                🎓 Official FinSight Demonstration Dataset
            </div>
            <div style="margin-top:7px;line-height:1.55;">
                You are viewing a fictional company created solely to demonstrate the complete FinSight workflow.
                The figures are sample financial data and do not represent a real organisation.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("## 1. Company Overview")
    with st.container(border=True):
        st.markdown(f"### {DEMO_COMPANY['company_name']}")
        st.write(f"**Primary Business Activity:** {DEMO_COMPANY['business_activity']}")
        st.write(f"**Currency:** {DEMO_COMPANY['currency']}")
        st.write(f"**Financial Year-End:** {DEMO_COMPANY['financial_year_end']}")
        st.write(f"**Financial Years:** {' • '.join(map(str, DEMO_COMPANY['financial_years']))}")

    with st.expander("View Sample Financial Data"):
        st.caption("All monetary figures are fictional and presented in RM million, except per-share information.")
        st.dataframe(demo_financial_table(), hide_index=True, width="stretch")

    with st.expander("View Layer 1 Calculation Results"):
        st.caption(
            "These tables are produced directly by the FinSight Financial Calculator. "
            "First-year average-balance ratios use closing balances because prior-year balances are unavailable."
        )
        demo_results = calculate_all_layer1(demo_financial_table(), DEMO_YEARS)
        tabs = st.tabs(list(demo_results.keys()))
        for tab, (group_name, result_df) in zip(tabs, demo_results.items()):
            with tab:
                st.markdown(f"### {group_name}")
                st.dataframe(
                    format_layer1_table(result_df),
                    hide_index=True,
                    width="stretch",
                )
                with st.expander("Formula Reference"):
                    for ratio_name, formula in LAYER1_FORMULAS[group_name].items():
                        st.write(f"**{ratio_name}:** {formula}")

    st.divider()

    st.markdown("## 2. FinSight Score")
    score_col, breakdown_col = st.columns([1.2, 1])
    with score_col:
        finsight_score_bar(89)
    with breakdown_col:
        st.markdown("### Score Breakdown")
        st.write("📈 Profitability: 14 / 15")
        st.write("💧 Liquidity: 7 / 10")
        st.write("⚙️ Efficiency: 6 / 10")
        st.write("🏦 Solvency: 10 / 10")
        st.write("📊 Market Performance: 5 / 5")
        st.write("🟢 Altman Z: 14 / 15")
        st.write("🔵 Piotroski F: 14 / 15")
        st.write("🟣 Beneish M: 19 / 20")

    st.divider()

    query_value = st.query_params.get("demo_analysis")
    if isinstance(query_value, list):
        query_value = query_value[0] if query_value else None

    previous_selected = st.session_state.get("demo_selected_category")

    if query_value:
        selected = query_value
        if selected != previous_selected:
            st.session_state.demo_scroll_pending = True
        st.session_state.demo_selected_category = selected
    else:
        selected = previous_selected

    demo_profit_data = dict(CATEGORY_DATA["Profitability"])
    demo_profit_result = demo_profitability_result()
    demo_profit_data["trend"], demo_profit_data["flag"] = profitability_signal(
        demo_profit_result
    )

    st.markdown("## 3. Financial Ratio Categories")
    st.caption("Review the Overall Trend and Key Flag, then select View Analysis for details.")

    c1, c2, c3 = st.columns(3)
    with c1:
        demo_category_link(
            "Profitability",
            demo_profit_data,
            selected == "Profitability",
        )
    with c2:
        demo_category_link(
            "Liquidity",
            CATEGORY_DATA["Liquidity"],
            selected == "Liquidity",
        )
    with c3:
        demo_category_link(
            "Efficiency",
            CATEGORY_DATA["Efficiency"],
            selected == "Efficiency",
        )

    c4, c5 = st.columns(2)
    with c4:
        demo_category_link(
            "Solvency",
            CATEGORY_DATA["Solvency"],
            selected == "Solvency",
        )
    with c5:
        demo_category_link(
            "Market Performance",
            CATEGORY_DATA["Market Performance"],
            selected == "Market Performance",
        )

    st.divider()

    st.markdown("## 4. Financial Diagnostic Models")
    st.caption("Select a model to explore its trend, components, guidance and flags.")

    d1, d2, d3 = st.columns(3)
    with d1:
        demo_diagnostic_link(
            "Altman Z Score",
            DIAGNOSTIC_DATA["Altman Z Score"],
            selected == "Altman Z Score",
        )
    with d2:
        demo_diagnostic_link(
            "Piotroski F Score",
            DIAGNOSTIC_DATA["Piotroski F Score"],
            selected == "Piotroski F Score",
        )
    with d3:
        demo_diagnostic_link(
            "Beneish M Score",
            DIAGNOSTIC_DATA["Beneish M Score"],
            selected == "Beneish M Score",
        )

    st.divider()
    demo_analysis_workspace(selected)

    st.divider()
    if st.button("📂 Open My Workspace", width="stretch"):
        st.query_params.clear()
        st.session_state.pop("demo_selected_category", None)
        st.session_state.pop("demo_scroll_pending", None)
        go("home")


REPORT_SECTION_CONFIG = OrderedDict({
    'Profitability Discussion': {'icon': '📈', 'aliases': ['profitability analysis', 'profitability', 'gross profit margin', 'net profit margin', 'return on assets', 'return on equity']},
    'Liquidity Discussion': {'icon': '💧', 'aliases': ['liquidity analysis', 'liquidity', 'current ratio', 'quick ratio']},
    'Efficiency Discussion': {'icon': '⚙️', 'aliases': ['efficiency analysis', 'efficiency', 'inventory turnover', 'receivables turnover', 'receivable turnover', 'asset turnover', 'collection period', 'inventory days', 'working capital turnover', 'cash turnover', 'ppe turnover']},
    'Solvency Discussion': {'icon': '🏦', 'aliases': ['solvency analysis', 'solvency', 'debt-to-equity', 'debt to equity', 'debt ratio', 'interest coverage', 'times interest earned', 'long-term debt']},
    'Market Performance Discussion': {'icon': '📊', 'aliases': ['market performance analysis', 'market ratio analysis', 'market performance', 'price earnings', 'price-to-earnings', 'p/e ratio', 'earnings yield', 'dividend yield', 'dividend payout', 'price-to-book']},
    'Financial Diagnostic Discussion': {'icon': '🧠', 'aliases': ['financial diagnostic', 'diagnostic models', 'altman z-score', 'altman z score', 'piotroski f-score', 'piotroski f score', 'beneish m-score', 'beneish m score']},
})

REPORT_COMPONENTS = ['Trend Identification','Supporting Evidence','Evaluation','Annual Report Integration','Professional Conclusion']
TREND_TERMS = re.compile(r'\b(increas(?:e|ed|ing)|decreas(?:e|ed|ing)|declin(?:e|ed|ing)|improv(?:e|ed|ing)|deteriorat(?:e|ed|ing)|recover(?:ed|y|ing)|fluctuat(?:e|ed|ing)|rose|risen|fell|dropped|grew|growth|upward|downward|stable|remained|highest|lowest|trend|movement|turnaround)\b', re.I)
EVALUATION_TERMS = re.compile(r'\b(because|due to|driven by|caused by|resulted from|attribut(?:ed|able) to|supported by|reflects?|indicates?|suggests?|consequently|therefore|as a result|owing to|mainly from|contributed to|linked to)\b', re.I)
ANNUAL_REPORT_TERMS = re.compile(r'\b(annual report|according to the .*report|management (?:stated|reported|highlighted)|chairman(?:\'s)? statement|management discussion|financial review|segment report|notes? to the financial statements|bursa malaysia|company announced|the group reported|the report states)\b', re.I)
CONCLUSION_TERMS = re.compile(r'\b(overall|in conclusion|to conclude|therefore|hence|consequently|this demonstrates|this indicates|the company should|management should|it is recommended|recommend(?:ed|ation)?|going forward|future outlook|financially stable|financial health|sustainable growth|suitable investment)\b', re.I)
EVIDENCE_TERMS = re.compile(r'(?:RM\s?\d|\d[\d,]*(?:\.\d+)?\s?(?:%|times?|days?|sen|million|billion)|from\s+[-+]?\d[\d,.%]*\s+to\s+[-+]?\d[\d,.%]*)', re.I)


def report_coverage_label(score):
    if score >= 90: return 'Comprehensive Coverage'
    if score >= 80: return 'Strong Coverage'
    if score >= 70: return 'Adequate Coverage'
    if score >= 60: return 'Basic Coverage'
    return 'Incomplete Coverage'


def report_coverage_icon(score):
    if score >= 80: return '🟢'
    if score >= 70: return '🟡'
    if score >= 60: return '🟠'
    return '🔴'


def extract_report_text(uploaded_file):
    raw = uploaded_file.getvalue()
    filename = uploaded_file.name.lower()
    if filename.endswith('.pdf'):
        reader = PdfReader(io.BytesIO(raw))
        return '\n<<<FINSIGHT_PAGE_BREAK>>>\n'.join((page.extract_text() or '') for page in reader.pages)
    if filename.endswith('.docx'):
        document = Document(io.BytesIO(raw))
        blocks = [p.text for p in document.paragraphs]
        for table in document.tables:
            for row in table.rows:
                blocks.append(' | '.join(cell.text for cell in row.cells))
        return '\n'.join(blocks)
    raise ValueError('Unsupported report format. Please upload PDF or DOCX.')


def normalise_report_text(text):
    text = text.replace('\u00ad', '').replace('\xa0', ' ')
    text = re.sub(r'[ \t]+', ' ', text)
    return re.sub(r'\n{3,}', '\n\n', text).strip()


def find_section_positions(text):
    lower = text.lower()
    positions = []
    threshold = max(1000, int(len(text) * 0.05))
    ceiling = int(len(text) * 0.90)
    for title, config in REPORT_SECTION_CONFIG.items():
        aliases = config['aliases']
        primary_hits = [m.start() for m in re.finditer(re.escape(aliases[0]), lower)]
        body_hits = [h for h in primary_hits if threshold < h < ceiling]
        if body_hits:
            candidate = min(body_hits)
        else:
            fallback_hits = []
            for alias in aliases[1:]:
                fallback_hits.extend(m.start() for m in re.finditer(re.escape(alias), lower))
            body_fallback = [h for h in fallback_hits if threshold < h < ceiling]
            if not body_fallback:
                continue
            candidate = min(body_fallback)
        positions.append((candidate, title))
    return sorted(positions)


def classify_report_page(page_text):
    head = page_text[:1800].lower()
    if 'table of contents' in head:
        return None
    heading_rules = [
        ('Profitability Discussion', ['profitability analysis', 'gross profit margin', 'net profit margin', 'return on assets', 'return on equity']),
        ('Liquidity Discussion', ['liquidity analysis', 'current ratio', 'quick ratio']),
        ('Efficiency Discussion', ['efficiency analysis', 'inventory turnover', 'receivables turnover', 'receivable turnover', 'collection period', 'inventory days', 'total asset turnover', 'working capital turnover', 'cash turnover', 'ppe turnover']),
        ('Solvency Discussion', ['solvency analysis', 'debt to equity', 'debt-to-equity', 'debt ratio', 'long term debt', 'long-term debt', 'times interest earned', 'interest coverage']),
        ('Market Performance Discussion', ['market ratio analysis', 'market performance analysis', 'price-to-book', 'price to book', 'price-earnings', 'price earnings', 'p/e ratio', 'earnings yield', 'dividend yield', 'dividend payout']),
        ('Financial Diagnostic Discussion', ['financial diagnostic', 'diagnostic models', 'altman z-score', 'altman z score', 'piotroski f-score', 'piotroski f score', 'beneish m-score', 'beneish m score']),
    ]
    for title, terms in heading_rules:
        if any(term in head for term in terms):
            return title
    return None


def split_report_sections(text):
    sections = {title: '' for title in REPORT_SECTION_CONFIG}
    if '<<<FINSIGHT_PAGE_BREAK>>>' in text:
        current = None
        for page in text.split('<<<FINSIGHT_PAGE_BREAK>>>'):
            page = page.strip()
            detected = classify_report_page(page)
            if detected:
                current = detected
            if current and page:
                sections[current] += '\n' + page
        return {k: v.strip() for k, v in sections.items()}

    positions = find_section_positions(text)
    for i, (start, title) in enumerate(positions):
        end = positions[i + 1][0] if i + 1 < len(positions) else len(text)
        chunk = text[start:end].strip()
        if len(chunk) > len(sections[title]):
            sections[title] = chunk
    return sections


def component_checks(section_text, title):
    text = section_text or ''
    words = len(re.findall(r'\b\w+\b', text))
    checks = {
        'Trend Identification': bool(TREND_TERMS.search(text)) and words >= 35,
        'Supporting Evidence': bool(EVIDENCE_TERMS.search(text)) and words >= 35,
        'Evaluation': len(EVALUATION_TERMS.findall(text)) >= 2 and words >= 80,
        'Annual Report Integration': bool(ANNUAL_REPORT_TERMS.search(text)) and words >= 80,
        'Professional Conclusion': bool(CONCLUSION_TERMS.search(text[-2200:])) and words >= 80,
    }
    diagnostics = None
    if title == 'Financial Diagnostic Discussion':
        lower = text.lower()
        diagnostics = {
            'Altman Z-Score': 'altman' in lower and ('z-score' in lower or 'z score' in lower),
            'Piotroski F-Score': 'piotroski' in lower and ('f-score' in lower or 'f score' in lower),
            'Beneish M-Score': 'beneish' in lower and ('m-score' in lower or 'm score' in lower),
        }
        if not all(diagnostics.values()):
            checks['Professional Conclusion'] = False
    return checks, diagnostics


def next_improvement_from_checks(checks, diagnostics=None):
    if diagnostics:
        missing_models = [name for name, present in diagnostics.items() if not present]
        if missing_models:
            return 'Include the missing diagnostic model discussion: ' + ', '.join(missing_models) + '.'
    messages = {
        'Trend Identification': 'Identify the direction and pattern of the financial trend.',
        'Supporting Evidence': 'Support the discussion using relevant ratios and financial data.',
        'Evaluation': 'Explain the factors contributing to the movement.',
        'Annual Report Integration': 'Integrate relevant company-specific evidence from the annual report.',
        'Professional Conclusion': 'Conclude the discussion with a clear financial implication or recommendation.',
    }
    missing = [item for item in REPORT_COMPONENTS if not checks.get(item)]
    return messages[missing[0]] if missing else 'Maintain the integration of evidence and professional judgement.'


def writing_feedback(checks, diagnostics=None):
    present = [c for c in REPORT_COMPONENTS if checks.get(c)]
    missing = [c for c in REPORT_COMPONENTS if not checks.get(c)]
    parts = []
    if present:
        parts.append('The discussion demonstrates ' + ', '.join(p.lower() for p in present) + '.')
    if diagnostics:
        missing_models = [name for name, present in diagnostics.items() if not present]
        if missing_models:
            parts.append('The report does not include ' + ' and '.join(missing_models) + ', which are required under the current FinSight diagnostic framework.')
    action_map = {
        'Trend Identification': 'identify the direction and pattern of the ratio movement',
        'Supporting Evidence': 'support the discussion with ratio values or financial statement data',
        'Evaluation': 'explain the company-specific factors contributing to the movement',
        'Annual Report Integration': 'link the explanation to relevant annual report disclosures',
        'Professional Conclusion': 'end with a clear financial implication or recommendation',
    }
    if missing:
        parts.append('Before submission, ' + '; '.join(action_map[m] for m in missing) + '.')
    else:
        parts.append('The five expected professional discussion components were detected.')
    parts.append('FinSight evaluates the presence of essential analytical components; final academic judgement remains with the lecturer.')
    return ' '.join(parts)


def analyse_report(uploaded_file):
    text = normalise_report_text(extract_report_text(uploaded_file))
    if len(text) < 500:
        raise ValueError('Very little readable text was extracted. The report may be image-based or scanned.')
    sections = split_report_sections(text)
    results = OrderedDict()
    for title, config in REPORT_SECTION_CONFIG.items():
        section_text = sections.get(title, '')
        checks, diagnostics = component_checks(section_text, title)
        if not section_text:
            checks = {component: False for component in REPORT_COMPONENTS}
        score = round(sum(bool(v) for v in checks.values()) / 5 * 100)
        results[title] = {
            'icon': config['icon'], 'score': score, 'label': report_coverage_label(score),
            'checks': checks, 'diagnostics': diagnostics,
            'next_improvement': next_improvement_from_checks(checks, diagnostics),
            'feedback': writing_feedback(checks, diagnostics),
            'section_detected': bool(section_text),
            'word_count': len(re.findall(r'\b\w+\b', section_text)),
        }
    return text, sections, results


def discussion_card(title, result, selected=False):
    icon, score, label = result['icon'], result['score'], result['label']
    improvement = result['next_improvement']
    border = '2px solid #2563eb' if selected else '1px solid #e5e7eb'
    bg = '#eff6ff' if selected else '#ffffff'
    html = f'''<div style="padding:20px 20px 8px;border-radius:18px 18px 0 0;border:{border};border-bottom:none;background:{bg};box-shadow:0 5px 18px rgba(0,0,0,.05);min-height:270px;">
    <div style="font-size:22px;font-weight:800;">{icon} {title}</div>
    <div style="height:3px;border-radius:999px;margin:14px 0;background:linear-gradient(90deg,#2563eb,#60a5fa 55%,rgba(96,165,250,.08));"></div>
    <div style="font-size:14px;color:#6b7280;font-weight:700;">Coverage</div>
    <div style="font-size:18px;font-weight:800;margin-top:6px;">{report_coverage_icon(score)} {label}</div>
    <div style="font-size:13px;color:#6b7280;margin-top:4px;">{score}%</div>
    <div style="height:3px;border-radius:999px;margin:14px 0;background:linear-gradient(90deg,#2563eb,#60a5fa 55%,rgba(96,165,250,.08));"></div>
    <div style="font-size:14px;color:#6b7280;font-weight:700;">Next Improvement</div>
    <div style="font-size:15px;line-height:1.55;margin-top:7px;min-height:72px;">{improvement}</div>
    <div style="height:3px;border-radius:999px;margin:14px 0 0;background:linear-gradient(90deg,#2563eb,#60a5fa 55%,rgba(96,165,250,.08));"></div>
    </div>'''
    st.markdown(html, unsafe_allow_html=True)
    if st.button('Review Discussion →', key=f'review_{title}', width='stretch'):
        st.session_state.selected_report_discussion = title
        st.session_state.report_scroll_pending = True
        st.query_params['page'] = 'report'
        st.rerun()


def discussion_review(selected, results):
    st.markdown('<div id="report-review-anchor"></div>', unsafe_allow_html=True)
    st.markdown('## 📝 Discussion Writing Review')
    if not selected or selected not in results:
        st.info('Select any discussion card above to review its coverage and receive focused writing guidance.')
        return
    result = results[selected]
    st.markdown(f"### {result['icon']} {selected}")
    st.progress(result['score'] / 100)
    st.markdown(f"### {result['score']}% — {result['label']}")
    if st.session_state.pop('report_scroll_pending', False):
        components.html('<script>setTimeout(function(){const a=window.parent.document.getElementById("report-review-anchor");if(a){a.scrollIntoView({behavior:"smooth",block:"start"});}},350);</script>', height=0)
    left, right = st.columns([1, 1.15])
    with left:
        st.markdown('### 📋 Coverage Checklist')
        for item in REPORT_COMPONENTS:
            (st.success if result['checks'].get(item) else st.warning)(('✓ ' if result['checks'].get(item) else '○ ') + item)
        if result.get('diagnostics'):
            st.markdown('#### Diagnostic Models')
            for model, present in result['diagnostics'].items():
                (st.success if present else st.warning)(('✓ ' if present else '○ ') + model)
    with right:
        st.markdown('### 💡 FinSight Writing Assistant')
        st.info(result['feedback'])
        if not result['section_detected']:
            st.error('This discussion section was not detected in the uploaded report.')
    with st.expander('🎯 What Makes a Good Financial Analysis?'):
        st.markdown('''| Component | Description | Bloom |
|---|---|:---:|
| **Trend Identification** | Has the student identified the trend? | C3 |
| **Supporting Evidence** | Has the student supported the discussion using ratios or financial data? | C3 |
| **Evaluation** | Has the student explained the factors contributing to the movement? | **C5** |
| **Annual Report Integration** | Has the student linked the discussion to relevant information from the annual report? | C5 |
| **Professional Conclusion** | Has the student concluded the discussion appropriately? | C5 |''')
        st.caption('Analysis Report Progress measures content coverage only. It is not a mark or automated academic judgement.')


def report_page():
    st.title('📝 My Report')
    st.caption('Upload, monitor and strengthen your financial analysis report.')
    st.markdown('## Upload Current Report')
    uploaded = st.file_uploader('Upload your current report', type=['docx', 'pdf'], key='report_upload')
    results = None
    if uploaded is None:
        st.info('Upload a DOCX or PDF report to activate the Report Coverage Engine.')
        st.session_state.pop('report_analysis_results', None)
        st.session_state.pop('report_analysis_hash', None)
    else:
        file_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()
        try:
            if st.session_state.get('report_analysis_hash') != file_hash:
                with st.spinner('FinSight is reviewing the report coverage...'):
                    _, _, results = analyse_report(uploaded)
                st.session_state.report_analysis_results = results
                st.session_state.report_analysis_hash = file_hash
                st.session_state.pop('selected_report_discussion', None)
            else:
                results = st.session_state.get('report_analysis_results')
            st.success(f'Uploaded and analysed: {uploaded.name}')
        except Exception as exc:
            st.error(f'FinSight could not analyse this report: {exc}')
    if not results:
        st.stop()
    overall = round(sum(v['score'] for v in results.values()) / len(results))
    st.divider()
    st.markdown('## 📊 Analysis Report Progress')
    st.progress(overall / 100)
    st.markdown(f'### {overall}% — {report_coverage_label(overall)}')
    missing_sections = [title for title, item in results.items() if not item['section_detected']]
    if missing_sections:
        st.warning('Sections not detected: ' + ', '.join(missing_sections) + '.')
    st.caption('Progress reflects coverage of expected analytical components. It is not a mark.')
    q = st.query_params.get('report_discussion')
    if isinstance(q, list):
        q = q[0] if q else None
    previous = st.session_state.get('selected_report_discussion')
    selected = previous or q
    if q and q != previous:
        st.session_state.report_scroll_pending = True
        st.session_state.selected_report_discussion = q
        selected = q
    st.divider()
    st.markdown('## Discussion Coverage')
    st.caption('Select Review Discussion for focused writing guidance.')
    titles = list(results)
    for row in (titles[:3], titles[3:]):
        cols = st.columns(3)
        for col, title in zip(cols, row):
            with col: discussion_card(title, results[title], selected == title)
    st.divider()
    discussion_review(selected, results)


# Preserve the public interactive demo route when a demo-analysis link is selected.
requested_page = st.query_params.get("page")
if isinstance(requested_page, list):
    requested_page = requested_page[0] if requested_page else None

if requested_page == "interactive_demo" or st.query_params.get("demo_analysis"):
    st.session_state.page = "interactive_demo"
elif requested_page == "report" or st.query_params.get("report_discussion"):
    st.session_state.page = "report"

sidebar_navigation()
if st.session_state.page == "welcome":
    welcome_page()
elif st.session_state.page == "login":
    login_page()
elif st.session_state.page == "register":
    register_page()
elif st.session_state.page == "home":
    home_page()
elif st.session_state.page == "create_workspace":
    create_workspace_page()
elif st.session_state.page == "workspace_summary":
    workspace_summary_page()
elif st.session_state.page == "data_entry_a":
    data_entry_page("company_a", "Company A", "company_a_complete")
elif st.session_state.page == "company_a_complete":
    company_a_complete_page()
elif st.session_state.page == "data_entry_b":
    data_entry_page("company_b", "Company B", "company_b_complete")
elif st.session_state.page == "company_b_complete":
    company_b_complete_page()
elif st.session_state.page == "dashboard":
    dashboard_page()
elif st.session_state.page == "about_finsight":
    about_finsight_page()
elif st.session_state.page == "interactive_demo":
    interactive_demo_page()
elif st.session_state.page == "report":
    report_page()
else:
    welcome_page()
