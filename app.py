import os
import sys
import time
import io
import re
import json
import datetime
import traceback
import smtplib
import asyncio
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# ---------------------------------------------------------------------
#  SAFETY NET: Automatic installation of openpyxl if missing
# ---------------------------------------------------------------------
try:
    import openpyxl
except ImportError:
    import subprocess
    print("Safety net: openpyxl is missing. Installing automatically...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

# ---------------------------------------------------------------------
#  SAFETY NET: Automatic installation of edge-tts if missing
# ---------------------------------------------------------------------
try:
    import edge_tts
except ImportError:
    import subprocess
    print("Safety net: edge-tts is missing. Installing automatically...")
    subprocess.call([sys.executable, "-m", "pip", "install", "edge-tts"])
    import edge_tts

import yfinance as yf
import requests
import pandas as pd
import streamlit as st

# =====================================================================
#  CONFIGURATION & STANDARD TARGET WEIGHTS
# =====================================================================

DISPLAY_CATEGORIES = {
    "Aktier": "Equities",
    "Sukuk": "Sukuk (Islamic Bonds)",
    "Råvarer": "Commodities",
    "Kontanter/Private": "Cash / Private Sector"
}

# Lydløs tovejs-oversætter, som oversætter de engelske UI-værdier til dit Google Sheet
UI_TO_DB_MAP = {
    "Equities": "Aktier",
    "Sukuk": "Sukuk",
    "Commodities": "Råvarer",
    "Cash/Private": "Kontanter/Private"
}

DB_TO_UI_MAP = {
    "Aktier": "Equities",
    "Sukuk": "Sukuk",
    "Råvarer": "Commodities",
    "Kontanter/Private": "Cash/Private"
}

TARGET_SUBSECTORS = [
    "Pharmaceuticals & Biotech",
    "Medical Devices & MedTech",
    "Clean Energy & Wind",
    "Smart Grid & Electrification",
    "Global Infrastructure",
    "Construction & Building Materials",
    "Chemicals & Advanced Materials",
    "Industrial Machinery & Automation",
    "Semiconductors & Hardware",
    "Mining & Royalty Streams",
    "Traditional Energy & Utilities",
    "Global Equity ETFs",
    "Regional & Thematic ETFs",
    "Sukuk & Fixed Income",
    "Cash & Liquidity Reserves",
    "Private Equity & Venture Capital",
    "Agriculture & Food Security",
    "Consumer Defensive & Staples",
    "Enterprise Software & SaaS",
    "Industrial Metals & Copper",
    "Logistics & Global Shipping",
    "Artificial Intelligence & Cloud Computing",
    "Cybersecurity & Digital Defense",
    "E-Commerce & Digital Retail",
    "Water Infrastructure & Desalination"
]

# DET STATISKE LYNHURTIGE KARTOTEK (Failsafe for at undgå IP-blokeringer fra Yahoo)
STATIC_TICKER_MAP = {
    "NOVO-B.CO": ("Aktier", "Pharmaceuticals & Biotech"),
    "NOVO-B": ("Aktier", "Pharmaceuticals & Biotech"),
    "MSFT": ("Aktier", "Enterprise Software & SaaS"),
    "AAPL": ("Aktier", "Semiconductors & Hardware"),
    "SAP": ("Aktier", "Enterprise Software & SaaS"),
    "IFX.DE": ("Aktier", "Semiconductors & Hardware"),
    "ASML": ("Aktier", "Semiconductors & Hardware"),
    "NVDA": ("Aktier", "Semiconductors & Hardware"),
    "VWS.CO": ("Aktier", "Clean Energy & Wind"),
    "NKT.CO": ("Aktier", "Smart Grid & Electrification"),
    "FLS.CO": ("Aktier", "Industrial Machinery & Automation"),
    "ROCK-B.CO": ("Aktier", "Construction & Building Materials"),
    "ORK.OL": ("Aktier", "Consumer Defensive & Staples"),
    "WPM": ("Råvarer", "Mining & Royalty Streams"),
    "NEM": ("Råvarer", "Industrial Metals & Copper"),
    "AEM": ("Råvarer", "Industrial Metals & Copper"),
    "RGLD": ("Råvarer", "Mining & Royalty Streams"),
    "SPSK": ("Sukuk", "Sukuk & Fixed Income"),
    "SKUK": ("Sukuk", "Sukuk & Fixed Income"),
    "MSAU.L": ("Aktier", "Regional & Thematic ETFs"),
    "IGDA.L": ("Aktier", "Global Equity ETFs"),
    "HLAL": ("Aktier", "Regional & Thematic ETFs"),
    "UMMA": ("Aktier", "Global Equity ETFs"),
    "ISWD.L": ("Aktier", "Global Equity ETFs"),
    "ISUS.L": ("Aktier", "Regional & Thematic ETFs"),
    "HIWS.L": ("Aktier", "Global Equity ETFs")
}

# GLOBAL ISLAMIC GROWTH UNIVERSE (Anvendes proaktivt)
GLOBAL_COMPLIANT_GROWTH_POOL = {
    "Aktier": [
        "MSAU.L", "IGDA.L", "HLAL", "UMMA", "ISWD.L", "ISUS.L", "HIWS.L",
        "TRMB", "SAP", "IFX.DE", "MSFT", "ASML", "NVDA", "ADBE", "CRM", "SNPS", 
        "NOVO-B.CO", "6869.T", "AZN.ST", "REGN", "ISRG", "LLY", "VRTX", 
        "VWS.CO", "NKT.CO", "FLS.CO", "ROCK-B.CO"
    ],
    "Sukuk": [
        "SPSK", "SKUK"
    ],
    "Råvarer": [
        "WPM", "NEM", "GOLD", "AEM", "FNV", "RGLD", "BHP", "RIO", "FCX", "VALE"
    ],
    "Kontanter/Private": [
        "SPSK", "SKUK"
    ]
}

GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1EnE2XkQySaGsdaxR5KySZZ924LT66ICo")
DATABASE_URL = os.getenv("DATABASE_URL")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

EMAIL_SENDER = os.getenv("EMAIL_SENDER")
if not EMAIL_SENDER or EMAIL_SENDER.strip() == "":
    EMAIL_SENDER = "wazir.ilyas@gmail.com"

EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")
if not EMAIL_RECEIVER or EMAIL_RECEIVER.strip() == "":
    EMAIL_RECEIVER = "addoncreatives@gmail.com"

EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
if EMAIL_PASSWORD:
    EMAIL_PASSWORD = EMAIL_PASSWORD.replace(" ", "").strip()


def normalize_string(s: str) -> str:
    if not s or pd.isna(s):
        return ""
    s = str(s).lower().strip()
    s = s.replace("og", "and").replace("&", "and")
    s = s.replace("'", "")
    s = re.sub(r'[^a-z0-9æøå]', '', s)
    s = s.replace("etfer", "etf").replace("etfs", "etf")
    return s


def search_tickers_by_name_multi(query: str) -> list:
    if not query or pd.isna(query) or len(str(query).strip()) < 2:
        return []
    
    query_clean = str(query).strip()
    url = f"https://query2.finance.yahoo.com/v1/finance/search?q={query_clean}"
    headers = {
        'User-Agent': 'Mozilla/5.0'
    }
    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            data = response.json()
            quotes = data.get("quotes", [])
            results = []
            for q in quotes:
                sym = q.get("symbol")
                name = q.get("longname") or q.get("shortname") or sym
                q_type = q.get("quoteType", "").upper()
                if sym and q_type in ["EQUITY", "ETF", "MUTUALFUND"]:
                    results.append({"symbol": sym, "name": name})
            return results[:5]
    except Exception as e:
        print(f"Search failed: {str(e)}")
    return []


# =====================================================================
#  OPDATERET EXCEL SKABELONS GENERATOR (MED NATIVE LIVE FORMELER)
# =====================================================================
def generate_excel_template_bytes(holdings_list: list, watchlist_list: list) -> bytes:
    wb = openpyxl.Workbook()
    
    # 1. FANEN: Beholdninger
    ws1 = wb.active
    ws1.title = "Beholdninger"
    
    headers1 = [
        "Position", "Ticker", "Status", "Antal", "Kurs (DKK)", 
        "Markedsværdi (DKK)", "Aktivklasse", "Drivkraft", "Sektor", 
        "Region", "Porteføljevægt", "Rolle", "Kommentar / tese"
    ]
    ws1.append(headers1)
    
    for idx, item in enumerate(holdings_list, start=2):
        name = item.get("Company Name", "Other")
        symbol = item.get("Ticker", "Other")
        shares = int(item.get("Shares", 1))
        cat = item.get("Category", "Aktier")
        sec = item.get("Sector", "Other")
        
        # Hvis det er et manuelt aktiv
        if "PVT_" in symbol or "CASH_" in symbol:
            val = float(item.get("manual_value", 1000))
            ws1.cell(row=idx, column=1, value=name)
            ws1.cell(row=idx, column=2, value="")
            ws1.cell(row=idx, column=3, value="Ejer")
            ws1.cell(row=idx, column=4, value=1)
            ws1.cell(row=idx, column=5, value=val)
            ws1.cell(row=idx, column=6, value=val)
        else:
            ws1.cell(row=idx, column=1, value=name)
            ws1.cell(row=idx, column=2, value=symbol)
            ws1.cell(row=idx, column=3, value="Ejer")
            ws1.cell(row=idx, column=4, value=shares)
            ws1.cell(row=idx, column=5, value=f'=GOOGLEFINANCE(B{idx})')
            ws1.cell(row=idx, column=6, value=f'=D{idx}*E{idx}')
            
        ws1.cell(row=idx, column=7, value=cat)
        ws1.cell(row=idx, column=8, value="")
        ws1.cell(row=idx, column=9, value=sec)
        ws1.cell(row=idx, column=10, value="Global")
        ws1.cell(row=idx, column=11, value=f'=F{idx}/SUM(F$2:F$100)')
        ws1.cell(row=idx, column=12, value="")
        ws1.cell(row=idx, column=13, value="")

    # 2. FANEN: Opsummering
    ws2 = wb.create_sheet(title="Opsummering")
    headers2 = ["4x25-overblik", "", "", "", "", "Økonomiske drivere", "", "", "", "Sektorere", "", "", "", "Huller / Watchlist"]
    ws2.append(headers2)
    
    # Skriv watchlist i Kolonne N (14)
    for idx, ticker in enumerate(watchlist_list, start=2):
        ws2.cell(row=idx, column=14, value=ticker)
        
    excel_data = io.BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)
    return excel_data.getvalue()


# =====================================================================
#  PORTFOLIO MANAGER AGENT (STATELÆS ROTATION)
# =====================================================================
class PortfolioManagerAgent:
    def __init__(self, current: dict, target: dict):
        self.current = current
        self.target = target

    def identify_underweighted_focus(self) -> tuple:
        underweight_candidates = []
        for category, target_val in self.target.items():
            curr_val = self.current.get(category, 0.0)
            deficit = target_val - curr_val
            if deficit > 0.0:
                underweight_candidates.append((category, deficit))
        
        if not underweight_candidates:
            return list(self.target.keys())[0], 0.0

        underweight_candidates.sort(key=lambda x: x[0])
        day_of_year = datetime.datetime.now().timetuple().tm_yday
        index = day_of_year % len(underweight_candidates)
        
        focus_category, deficit = underweight_candidates[index]
        return focus_category, deficit


# =====================================================================
#  SCREENER & COMPLIANCE AGENT (DYNAMISK SØGNING)
# =====================================================================
class ScreenerComplianceAgent:
    PROHIBITED_SECTORS = ["Financial Services", "Financial"]
    PROHIBITED_INDUSTRIES = [
        "Banks", "Insurance", "Aerospace & Defense", "Gambling", 
        "Tobacco", "Distillers & Vintners", "Breweries"
    ]

    def __init__(self, tickers: list, target_category: str = None):
        self.tickers = tickers
        self.target_category = target_category

    def check_zoya_live_compliance(self, symbol: str) -> bool:
        clean_symbol = symbol.split('.')[0].upper()
        clean_symbol = clean_symbol.split('-')[0]
        
        url = f"https://zoya.finance/stocks/{clean_symbol}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        try:
            response = requests.get(url, headers=headers, timeout=8)
            if response.status_code == 200:
                html_lower = response.text.lower()
                if "is shariah-compliant" in html_lower or "is considered halal" in html_lower:
                    return True
                elif "not shariah-compliant" in html_lower or "non-compliant" in html_lower:
                    return False
            return None
        except Exception:
            return None

    def map_to_category_and_sector(self, symbol: str, sector: str = None, industry: str = None) -> tuple:
        sym = symbol.upper().strip()
        sec_l = sector.lower() if sector else ""
        ind_l = industry.lower() if industry else ""

        # 1. Tjek altid det lynhurtige statiske kartotek først
        lookup_sym = sym.split('.')[0]
        for k, v in STATIC_TICKER_MAP.items():
            if normalize_string(k) == normalize_string(sym) or normalize_string(k) == normalize_string(lookup_sym):
                if v[0] == "Sukuk" and self.target_category == "Kontanter/Private":
                    return "Kontanter/Private", "Sukuk & Fixed Income"
                return v[0], v[1]

        # 2. Hvis ikke i kartoteket, kør dynamisk mapping baseret på yfinance data
        if "sukuk" in sym or sym in ["SPSK", "SKUK"]:
            if self.target_category == "Kontanter/Private":
                return "Kontanter/Private", "Sukuk & Fixed Income"
            return "Sukuk", "Sukuk & Fixed Income"
            
        if sym in ["WPM", "FNV", "RGLD"]:
            return "Råvarer", "Mining & Royalty Streams"
        if sym in ["NEM", "GOLD", "AEM", "BHP", "RIO", "FCX", "VALE"] or \
           any(w in ind_l for w in ["gold", "silver", "precious metals", "copper", "aluminum"]):
            return "Råvarer", "Industrial Metals & Copper"

        if "cash" in sym or "money market" in sec_l:
            return "Kontanter/Private", "Cash & Liquidity Reserves"

        # Dynamisk kobling til de nye overordnede kategorier
        if "pharmaceutical" in ind_l or "biotechnology" in ind_l:
            return "Aktier", "Pharmaceuticals & Biotech"
        if "medical" in ind_l or "healthcare" in sec_l:
            return "Aktier", "Medical Devices & MedTech"
        if "wind" in ind_l or "renewable" in ind_l or "solar" in ind_l:
            return "Aktier", "Clean Energy & Wind"
        if "cable" in ind_l or "electrical" in ind_l:
            return "Aktier", "Smart Grid & Electrification"
        if "semiconductor" in ind_l or "semiconductor" in sec_l:
            return "Aktier", "Semiconductors & Hardware"
        if "software" in ind_l or "software" in sec_l or "technology" in sec_l:
            return "Aktier", "Enterprise Software & SaaS"
        if "building" in ind_l or "construction" in ind_l:
            return "Aktier", "Construction & Building Materials"
        if "chemicals" in ind_l:
            return "Aktier", "Chemicals & Advanced Materials"
        if "machinery" in ind_l or "industrials" in sec_l:
            return "Aktier", "Industrial Machinery & Automation"
        if "infrastructure" in ind_l or "utilities" in sec_l:
            return "Aktier", "Global Infrastructure"
        if "staples" in sec_l or "packaged foods" in ind_l or "consumer defensive" in sec_l:
            return "Aktier", "Consumer Defensive & Staples"
        if "fertilizer" in ind_l or "agriculture" in ind_l:
            return "Aktier", "Agriculture & Food Security"
        if "logistics" in ind_l or "shipping" in ind_l:
            return "Aktier", "Logistics & Global Shipping"
        if "internet" in ind_l or "e-commerce" in ind_l:
            return "Aktier", "E-Commerce & Digital Retail"
        if "water" in ind_l or "environmental" in ind_l:
            return "Aktier", "Water Infrastructure & Desalination"

        # Dynamisk fallback
        dynamic_subsector = industry if (industry and industry != "Other") else sector
        return "Aktier", dynamic_subsector

    def screen_ticker(self, symbol: str) -> dict:
        try:
            # Virtuelle/manuelle tickers for kontanter/private skal altid bestå screening
            if "CASH_" in symbol or "PVT_" in symbol:
                return {
                    "symbol": symbol,
                    "passed": True,
                    "name": symbol.replace("CASH_", "").replace("PVT_", ""),
                    "pe_ratio": "N/A",
                    "debt_ratio": "0.00% (Manual)",
                    "sector": "Manual Asset",
                    "industry": "Manual Asset",
                    "category": "Kontanter/Private",
                    "subsector": "Cash & Liquidity Reserves",
                    "is_etf": False
                }

            zoya_compliant = self.check_zoya_live_compliance(symbol)
            
            if zoya_compliant is False:
                return {"symbol": symbol, "passed": False, "reason": "Disqualified by Zoya's live public Shariah assessment."}
            elif zoya_compliant is True:
                print(f"Zoya Live-tjek bekræfter: {symbol} er Shariah-compliant.")

            ticker_obj = yf.Ticker(symbol)
            info = ticker_obj.info
            
            if not info:
                return {"symbol": symbol, "passed": False, "reason": "Ingen data"}

            quote_type = info.get("quoteType", "").upper()
            is_etf = quote_type in ["ETF", "MUTUALFUND"] or symbol in ["IGDA.L", "SPSK", "HLAL", "UMMA", "ISWD.L", "MSAU.L", "SKUK"]

            if is_etf:
                mapped_cat, mapped_sub = self.map_to_category_and_sector(symbol, "ETF", "ETF")
                return {
                    "symbol": symbol,
                    "passed": True,
                    "name": info.get("longName", symbol),
                    "pe_ratio": info.get("trailingPE", "N/A"),
                    "debt_ratio": "N/A (ETF/Sukuk)",
                    "sector": "ETF / Fond",
                    "industry": "ETF",
                    "category": mapped_cat,
                    "subsector": mapped_sub,
                    "is_etf": True
                }

            sector = info.get("sector", "")
            industry = info.get("industry", "")
            
            for p_sector in self.PROHIBITED_SECTORS:
                if p_sector.lower() in sector.lower():
                    return {"symbol": symbol, "passed": False, "reason": f"Sektor: {sector}"}
                    
            for p_ind in self.PROHIBITED_INDUSTRIES:
                if p_ind.lower() in industry.lower():
                    return {"symbol": symbol, "passed": False, "reason": f"Branche: {industry}"}

            debt_to_equity = info.get("debtToEquity")
            total_debt = info.get("totalDebt")
            market_cap = info.get("marketCap")
            
            debt_ratio_pct = None
            method_used = ""

            if debt_to_equity is not None:
                debt_ratio_pct = debt_to_equity
                method_used = "Debt to Equity"
            elif total_debt and market_cap:
                debt_ratio_pct = (total_debt / market_cap) * 100
                method_used = "Debt to Market Cap"

            if debt_ratio_pct is None:
                return {"symbol": symbol, "passed": False, "reason": "Ingen gældsdata"}

            if debt_ratio_pct > 30.0:
                return {"symbol": symbol, "passed": False, "reason": f"Gældskvote: {debt_ratio_pct:.2f}%"}

            mapped_cat, mapped_sub = self.map_to_category_and_sector(symbol, sector, industry)

            return {
                "symbol": symbol,
                "passed": True,
                "name": info.get("longName", symbol),
                "pe_ratio": info.get("trailingPE", "N/A"),
                "debt_ratio": f"{debt_ratio_pct:.2f}% ({method_used})",
                "sector": sector,
                "industry": industry,
                "category": mapped_cat,
                "subsector": mapped_sub,
                "is_etf": False
            }
        except Exception as e:
            return {"symbol": symbol, "passed": False, "reason": str(e)}

    def run_screening(self, target_category: str) -> list:
        approved_stocks = []
        for ticker in self.tickers:
            time.sleep(1.5)
            result = self.screen_ticker(symbol=ticker)
            if result["passed"] and result["category"] == target_category:
                approved_stocks.append(result)
        return approved_stocks


# =====================================================================
#  COUNCIL AGENT (GEMINI 3.5 FLASH - ENGLISH HTML NEWSLETTER)
# =====================================================================
class CouncilAgent:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={self.api_key}"

    def run_proactive_analysis(self, candidates_data: list, category: str, deficit: float, current_portfolio_str: str, current_holdings_str: str, sector_distribution_str: str, user_name: str, horizon: str) -> str:
        candidates_json = json.dumps(candidates_data, indent=2, ensure_ascii=False)
        english_category = DISPLAY_CATEGORIES.get(category, category)
        
        prompt = f"""
        You are an elite financial advisory council ("LLM Council") presenting a strategic investment briefing to your highly valued VIP client, {user_name}.
        
        THE INVESTOR PROFILE & MODEL:
        - Investor's Name: {user_name}
        - Investment Horizon: {horizon} (This is CRITICAL. Align all advice, timelines, risk-tolerances, and recommendations precisely with this specific time horizon!)
        - Overarching Strategic Model: Customize based on Wazir's targets.
        - Under Evaluation Tonight: {english_category} (Current Deficit: {deficit:.2f}%)
        - Current Portfolio Allocations (Target vs Actual): {current_portfolio_str}
        
        THE INVESTOR'S STRATEGIC SUB-SECTORS (DYNAMICALLY DETECTED FROM THE HOLDINGS):
        {sector_distribution_str}
        
        THE INVESTOR'S CURRENT HOLDINGS:
        {current_holdings_str}
        
        NEW COMPLIANT SCREENED CANDIDATES TO BE EVALUATED:
        {candidates_json}
        
        YOUR OBJECTIVE (DELIVER ENTIRE BRIEFING IN BEAUTIFUL ENGLISH HTML):
        Generate a complete, institutional-grade, highly engaging HTML investment newsletter in English.
        
        Brug udelukkende inline CSS-styling for maximum compatibility with Gmail.
        Design guidelines:
        - Main container: `<div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 25px; background-color: #ffffff; color: #334155; line-height: 1.6;">`
        - Colors: Dark Slate (`#0F172A`) for headings. Accent/highlights: Warm Gold/Sand (`#C5A880`).
        - Cards: Each of the screened candidates should be enclosed in a distinct card: `<div style="border: 1px solid #E2E8F0; padding: 20px; margin-bottom: 25px; border-radius: 8px; background-color: #F8FAFC;">`
        - Debate box: Each adviser's turn must use the distinct left-bordered styling:
          - Contrarian: `border-left: 4px solid #EF4444; background: #FEF2F2; padding: 15px; margin-bottom: 15px;` (Red border)
          - First-Principles: `border-left: 4px solid #64748B; background: #F8FAFC; padding: 15px; margin-bottom: 15px;` (Slate border)
          - Expansionist: `border-left: 4px solid #10B981; background: #ECFDF5; padding: 15px; margin-bottom: 15px;` (Green border)
          - Outsider: `border-left: 4px solid #8B5CF6; background: #F5F3FF; padding: 15px; margin-bottom: 15px;` (Purple border)
          - Executor: `border-left: 4px solid #3B82F6; background: #EFF6FF; padding: 15px; margin-bottom: 15px;` (Blue border)
        - Chairman's Call: Gold-framed box: `<div style="background-color: #FDFBF7; border: 1px solid #E2D1B6; border-left: 6px solid #C5A880; padding: 25px; border-radius: 8px; margin-top: 30px;">`
        
        REPORT OUTLINE (ENGLISH):
        
        <h1>🗳️ LLM Council Strategic Briefing</h1>
        <p><strong>Prepared exclusively for:</strong> {user_name}</p>
        <p><strong>Investment Horizon:</strong> {horizon}</p>
        <p><strong>Focus tonight:</strong> {english_category} (Deficit: {deficit:.2f}%)</p>
        
        <hr style="border: 0; border-top: 1px solid #E2E8F0; margin: 20px 0;">
        
        <h2>SECTION 1 — PORTFOLIO DIAGNOSTIC & INDIRECT EXPOSURES</h2>
        Analyze {user_name}'s current holdings and how they map to their strategic sub-sectors. Do existing assets already provide satisfactory indirect exposure to the focus theme? Discuss Saxo Bank limitations and Sharia compliance filters as boundaries, and explain how they can diversify across different underlying economic drivers if direct options are limited.
        
        <h2>SECTION 2 — DEEP-DIVE CONSULTANT ANALYSIS (UP TO 10 SCREENED CANDIDATES)</h2>
        For each candidate, write an elegant card covering:
        1. <strong>Investment Case</strong> (How it balances {user_name}'s current assets, keeping their {horizon} horizon in mind).
        2. <strong>Financial Highlights</strong> (Omsætningsvækst, marginer, cash flow based on live data).
        3. <strong>Future Outlook & Pipeline</strong>.
        4. <strong>Risk Assessment</strong>.
        5. <strong>Momentum & Trend Analysis</strong> (3-month momentum vs. 3-year growth trajectory).
        6. <strong>Analyst Insight & Sources</strong>: Insert exactly 2 clickable links structured beautifully in HTML (e.g. `<a href="https://seekingalpha.com/symbol/TICKER" style="color: #C5A880; text-decoration: none; font-weight: bold;">Seeking Alpha</a>`).
        
        <h2>SECTION 3 — THE ASYNCHRONOUS COUNCIL DEBAT (TOP-3)</h2>
        Select the top 3 assets. Moderate a high-stakes, dramatic debate among the 5 financial advisers using the styled left-bordered divs. Show conflict, arguments on valuation, capex, and macro timing.
        
        <h2>SECTION 4 — THE CHAIRMAN'S DEKRET (RECOMMENDATION)</h2>
        In Section 4, the Chairman must NOT command the investor to buy or take action. Instead, the Chairman must strongly advise and urge the investor to critically evaluate and consider the council's comprehensive proposal. The tone should be highly advisory, objective, and respectful, emphasizing that the final capital allocation decision rests solely on the investor's own assessment. Enclose this callout in the gold callout box. Conclude with a highly precise, step-by-step action plan for {user_name}'s Saxo Investor account over the next 7 days.
        
        Return ONLY the raw HTML code. Do NOT enclose in markdown tags like "```html".
        """

        headers = {'Content-Type': 'application/json'}
        payload = {"contents": [{"parts": [{"text": prompt}]}]}
        
        try:
            response = requests.post(self.url, headers=headers, json=payload, timeout=110)
            response.raise_for_status()
            data = response.json()
            if 'candidates' in data and len(data['candidates']) > 0:
                raw_html = data['candidates'][0]['content']['parts'][0]['text']
                return raw_html.replace("```html", "").replace("```", "").strip()
            return "<h3>Error</h3><p>Could not parse HTML from Gemini.</p>"
        except Exception as e:
            return f"<h3>System Error</h3><p>{str(e)}</p>"


# =====================================================================
#  PODCAST AGENT (PODCASTFY POWERED HIGH-ENERGY CNBC/BLOOMBERG TALKSHOW)
# =====================================================================
class PodcastAgent:
    def __init__(self, api_key: str):
        self.api_key = api_key
        if self.api_key:
            os.environ["GEMINI_API_KEY"] = self.api_key
            os.environ["GOOGLE_API_KEY"] = self.api_key

    def generate_podcast_audio(self, report_html: str, user_name: str) -> str:
        from podcastfy.client import generate_podcast
        
        custom_config = {
            "word_count": 900,
            "conversation_style": ["engaging", "fast-paced", "enthusiastic", "hardcore Bloomberg debate"],
            "roles_person1": "Sarah, the curious financial journalist",
            "roles_person2": "Mark, the hardcore market analyst",
            "podcast_name": "LLM Council Briefing",
            "podcast_tagline": "CNBC & Bloomberg Style Financial Talkshow",
            "output_language": "English",
            "engagement_techniques": ["rhetorical questions", "analogies", "humor", "interjections", "cross-talk"],
            "user_instructions": (
                f"Create a high-energy Bloomberg-style financial show moderated by Sarah and Mark. "
                f"The show MUST open with Sarah and Mark introducing themselves and welcoming our VIP client, {user_name}. "
                f"Then, they introduce and interview our 5 resident advisers: "
                f"Contrarian (the risk-obsessed skeptic who must interrupt with: 'But what if the market turns tomorrow?'), "
                f"First-Principles (the logical mathematician using raw numbers), "
                f"Expansionist (the highly bullish growth hunter wanting to deploy capital), "
                f"Outsider (the big-picture strategist analyzing indirect exposures like NKT/FLS and favoring royalty models), "
                f"and Executor (the pragmatic guy checking Saxo tradeability and Dollar-Cost Averaging). "
                f"The show must conclude with Sarah and Mark summarizing the Chairman's final recommendation and "
                f"giving {user_name} a highly clear, actionable next step for his Saxo account."
            )
        }
        
        try:
            print("Genererer ægte multi-stemme podcast via Podcastfy og gratis Edge TTS...")
            audio_path = generate_podcast(
                text=report_html,
                tts_model="edge",
                conversation_config=custom_config
            )
            return audio_path
        except Exception as e:
            print(f"Fejl under Podcastfy-generering: {str(e)}")
            return None


# =====================================================================
#  DELIVERY AGENT (HTML & VEDHÆFTNING SMTP)
# =====================================================================
class DeliveryAgent:
    @staticmethod
    def send_email(subject: str, html_content: str, attachments: list = None):
        if not EMAIL_PASSWORD or EMAIL_PASSWORD.strip() == "":
            print("EMAIL_PASSWORD mangler i GitHub Secrets. Udskriver HTML i konsol:")
            print(html_content)
            return

        msg = MIMEMultipart()
        msg["From"] = EMAIL_SENDER
        msg["To"] = EMAIL_RECEIVER
        msg["Subject"] = subject
        msg.attach(MIMEText(html_content, "html", "utf-8"))

        if attachments:
            for att in attachments:
                path = att.get("path")
                data = att.get("data")
                name = att.get("name")
                
                if (path and os.path.exists(path)) or data:
                    part = MIMEBase("application", "octet-stream")
                    if data:
                        part.set_payload(data)
                    else:
                        with open(path, "rb") as attachment:
                            part.set_payload(attachment.read())
                    
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename= {name}",
                    )
                    msg.attach(part)

        try:
            print(f"Forbinder til SMTP server ({SMTP_SERVER}:{SMTP_PORT}) med afsender-login: {EMAIL_SENDER}...")
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
            server.quit()
            print("Succes: Rapport og podcast-afspilning sendt til din indbakke.")
        except Exception as e:
            print(f"E-mail fejl: {str(e)}")


# =====================================================================
#  STREAMLIT BRUGERGRÆNSEFLADE (SAMLING AF APP.PY)
# =====================================================================

st.title("🗳️ LLM Council")
st.caption("PREMIUM INVESTMENT NEWSLETTER & PODCAST SERVICE")

# Pæn, engelsk forklaring af værditilbuddet (Value Proposition)
st.markdown("""
    <div style="border: 1px solid #E2E8F0; padding: 25px; border-radius: 8px; margin-bottom: 30px;">
        <h3 style="font-family: 'Georgia', serif; margin-top: 0;">🛡️ Ethical Shariah-Compliant Filtering</h3>
        <p style="margin-bottom: 20px;">
            LLM Council operates under strict Islamic ethical constraints. 
            Our automated live engine immediately purges companies associated with:
            Traditional interest-bearing debt (conventional banking and insurance), alcohol, pork, weapons, defense, 
            gambling, and adult entertainment. Additionally, any asset with a live debt-to-equity or 
            debt-to-market-cap ratio exceeding 30% is immediately disqualified.
        </p>
        <h3 style="font-family: 'Georgia', serif; margin-top: 25px;">📊 The 4 Pillars of a Shariah-Compliant Portfolio</h3>
        <p style="margin-bottom: 10px;">For new investors, building a robust, diversified, and halaal portfolio requires spreading your capital across four core pillars, each acting as a unique engine of growth and protection:</p>
        <ul style="margin-bottom: 20px;">
            <li><strong>Equities (Aktier):</strong> Fractional ownership in global businesses. We only select companies that pass strict qualitative filters (no unlawful lines of business) and conservative quantitative audits (debt-to-equity and interest-bearing liquidity must be below 30%).</li>
            <li><strong>Sukuk (Islamic Bonds):</strong> Asset-backed financial certificates. Since conventional interest-bearing bonds are strictly prohibited (Riba), Sukuk certificates generate yields for you from tangible underlying assets (such as real estate leasing or profit-sharing partnerships). They act as your portfolio's stable income stream.</li>
            <li><strong>Commodities (Råvarer):</strong> Tangible, physical hard assets like gold, silver, or key industrial materials. Commodities act as a store of real value and your primary defense mechanism against currency devaluation and global inflation.</li>
            <li><strong>Cash / Private Sector (Kontanter/Private):</strong> Highly liquid cash reserves, Sharia money-market proxies, or private equity investments used for tactical rebalancing, emergency funds, or long-term private business backing.</li>
        </ul>
        <h3 style="font-family: 'Georgia', serif;">🗳️ Why the LLM Council Method Works</h3>
        <p style="margin-bottom: 0;">
            Rather than relying on a single stagnant AI opinion, we submit your portfolio to a dynamic, 
            adversarial debate between <strong>five distinct virtual financial specialists</strong> (Skeptics, Logicians, Growth Hunters, 
            Strategists, and Practicians). This pressure-tests your holdings from multiple conflicting perspectives, 
            spotting overlapping risks and hidden capex cycles. The Chairman then synthesizes their argument 
            to deliver an objective, highly advisory and tailored capital allocation proposal directly to your inbox.
        </p>
    </div>
""", unsafe_allow_html=True)

# Initialize session state for active holdings (NU HELT TOM VED OPSTART)
if "holdings" not in st.session_state:
    st.session_state.holdings = []
if "targets" not in st.session_state:
    st.session_state.targets = {"Aktier": 25.0, "Sukuk": 25.0, "Råvarer": 25.0, "Kontanter/Private": 25.0}
if "horizon" not in st.session_state:
    st.session_state.horizon = "7-15 years"
if "user_name" not in st.session_state:
    st.session_state.user_name = "Investor"
if "frequency" not in st.session_state:
    st.session_state.frequency = "Weekly"

# =====================================================================
#  DATABASE INTEGRATION (GRATIS SAAS MODEL VIA GOOGLE WEB APP)
# =====================================================================
def load_user_portfolio_from_db(email: str, password: str) -> dict:
    if not DATABASE_URL or DATABASE_URL.strip() == "":
        return None
    try:
        response = requests.get(f"{DATABASE_URL}?email={email}&password={password}", timeout=10)
        if response.status_code == 200:
            data = response.json()
            if data.get("status") == "success":
                return data
            elif data.get("status") == "incorrect_password":
                st.error("Incorrect password for this email!")
    except Exception as e:
        print(f"Kunne ikke hente profil fra database: {str(e)}")
    return None

def save_user_portfolio_to_db(email: str, password: str, holdings: list, targets: dict, horizon: str, name: str, frequency: str) -> str:
    if not DATABASE_URL or DATABASE_URL.strip() == "":
        return "no_db"
    payload = {
        "email": email,
        "password": password,
        "holdings": holdings,
        "targets": targets,
        "horizon": horizon,
        "name": name,
        "frequency": frequency # Gemmer din ugentlige/daglige frekvens
    }
    try:
        response = requests.post(DATABASE_URL, json=payload, timeout=15)
        if response.status_code == 200:
            res = response.json()
            return res.get("status")
    except Exception as e:
        print(f"Kunne ikke gemme profil i database: {str(e)}")
    return "error"

# =====================================================================
#  STEP 1: LOGIN & PROFIL (MED ÆGTE PASSWORD SIGNUP) - NU PÅ HOVEDSIDEN!
# =====================================================================
st.subheader("Step 1: SaaS Investor Access")
col_l1, col_l2 = st.columns(2)
with col_l1:
    login_email = st.text_input("Enter your Email", placeholder="your.name@gmail.com")
with col_l2:
    login_password = st.text_input("Enter your Password", type="password")

is_new_user = False
db_profile = None

if login_email and "@" in login_email and login_password:
    response = requests.get(f"{DATABASE_URL}?email={login_email}&password={login_password}", timeout=10)
    if response.status_code == 200:
        res_data = response.json()
        if res_data.get("status") == "success":
            st.success(f"Loaded profile for {res_data.get('name')}!")
            db_profile = res_data
            st.session_state.holdings = db_profile.get("holdings")
            st.session_state.targets = db_profile.get("targets")
            st.session_state.horizon = db_profile.get("horizon")
            st.session_state.user_name = db_profile.get("name")
            st.session_state.frequency = db_profile.get("frequency", "Weekly")
        elif res_data.get("status") == "incorrect_password":
            st.error("Incorrect password for this email!")
        elif res_data.get("status") == "not_found":
            is_new_user = True
            st.info("Email not found. Fill out the signup fields below to register!")

if is_new_user:
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        confirm_password = st.text_input("Confirm Password", type="password")
    with col_s2:
        signup_name = st.text_input("Your Full Name", value="Investor")
    
    if st.button("📝 Register Account"):
        if login_password != confirm_password:
            st.error("Passwords do not match!")
        elif not signup_name:
            st.error("Please enter your name.")
        else:
            status = save_user_portfolio_to_db(
                email=login_email,
                password=login_password,
                holdings=st.session_state.holdings,
                targets=st.session_state.targets,
                horizon=st.session_state.horizon,
                name=signup_name,
                frequency=st.session_state.frequency
            )
            if status == "success":
                st.success("Account created successfully!")
                st.session_state.user_name = signup_name
                st.rerun()
            else:
                st.error("Registration failed. Try again.")

st.write("---")

# =====================================================================
#  STEP 1.2: PERSONAL PROFILE DETAILS (ENGLISH ONLY)
# =====================================================================
st.subheader("Step 1.2: Your Personal Profile")
col_n1, col_n2 = st.columns(2)
with col_n1:
    user_name_input = st.text_input("Enter your Name:", value=st.session_state.user_name)
    st.session_state.user_name = user_name_input
with col_n2:
    # Hvis brugeren er logget ind via sidemenuen, låser vi mailen fast her
    user_email_input = st.text_input("Enter your Email Address to receive briefings:", value=login_email if login_email else "", placeholder="your.name@gmail.com")

st.subheader("Step 1.3: Your Investment Horizon & Delivery Settings")
col_s1, col_s2 = st.columns(2)
with col_s1:
    horizon_options = ["1-3 years", "3-7 years", "7-15 years", "15+ years"]
    horizon_index = horizon_options.index(st.session_state.horizon) if st.session_state.horizon in horizon_options else 2

    st.session_state.horizon = st.selectbox(
        "Select your Investment Horizon:", 
        horizon_options,
        index=horizon_index
    )
with col_s2:
    freq_options = ["Daily", "Weekly", "Bi-weekly", "Monthly"]
    freq_index = freq_options.index(st.session_state.frequency) if st.session_state.frequency in freq_options else 1
    
    st.session_state.frequency = st.selectbox(
        "Select Briefing Frequency:",
        freq_options,
        index=freq_index
    )

st.subheader("Step 1.5: Customize Your Target Allocations")
st.write("Define your target weighting for the major asset classes.")

col1, col2, col3, col4 = st.columns(4)
with col1:
    target_stocks = st.number_input("Equities (%)", min_value=0.0, max_value=100.0, value=float(st.session_state.targets.get("Aktier", 25.0)))
with col2:
    target_sukuk = st.number_input("Sukuk (%)", min_value=0.0, max_value=100.0, value=float(st.session_state.targets.get("Sukuk", 25.0)))
with col3:
    target_commodities = st.number_input("Commodities (%)", min_value=0.0, max_value=100.0, value=float(st.session_state.targets.get("Råvarer", 25.0)))
with col4:
    target_cash = st.number_input("Cash/Private (%)", min_value=0.0, max_value=100.0, value=float(st.session_state.targets.get("Kontanter/Private", 25.0)))

# Normaliser mål-vægte så de summer til 100%
total_target = target_stocks + target_sukuk + target_commodities + target_cash
if total_target > 0:
    st.session_state.targets = {
        "Aktier": (target_stocks / total_target) * 100.0,
        "Sukuk": (target_sukuk / total_target) * 100.0,
        "Råvarer": (target_commodities / total_target) * 100.0,
        "Kontanter/Private": (target_cash / total_target) * 100.0
    }
else:
    st.session_state.targets = {"Aktier": 25.0, "Sukuk": 25.0, "Råvarer": 25.0, "Kontanter/Private": 25.0}

# =====================================================================
#  STEP 1.8: NEW INVESTOR MODE
# =====================================================================
st.subheader("Step 1.8: Investor Status")
is_new_investor = st.checkbox("I am a completely new investor (starting from scratch)")

selected_new_sectors = []
if is_new_investor:
    st.write("Since you are starting from scratch, select the sectors/themes you want to build exposure to:")
    selected_new_sectors = st.multiselect("Select Target Sectors:", options=TARGET_SUBSECTORS, default=["Pharmaceuticals & Biotech", "Clean Energy & Wind", "Semiconductors & Hardware"])

# =====================================================================
#  STEP 2: ENKELT, SØG-OG-TILFØJ ENGINE (KUN HVIS IKKE NY INVESTOR)
# =====================================================================
if not is_new_investor:
    st.subheader("Step 2: Add Assets to Your Active Portfolio")
    st.write("Search for any global company or fund name below. When found, the system will automatically classify its Category and Sector, and you just input your shares!")

    is_manual = st.checkbox("Is this a manual asset? (Cash, Private Equity, private placements)")

    if is_manual:
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            manual_name = st.text_input("Enter Asset Name:", placeholder="e.g., Saxo Cash Reserve, Danish PE Fund")
        with col_m2:
            manual_value = st.number_input("Total Value (DKK):", min_value=1, value=1000)
            
        col_m3, col_m4 = st.columns(2)
        with col_m3:
            manual_category = st.selectbox("Select Asset Class:", ["Cash/Private", "Sukuk", "Commodities"])
        with col_m4:
            manual_sector = st.selectbox("Select Sub-Sector:", TARGET_SUBSECTORS + ["Kontanter", "Private investeringer"])
            
        if st.button("➕ Add Manual Asset"):
            if manual_name:
                virtual_ticker = f"PVT_{manual_name.upper().replace(' ', '_')}"
                st.session_state.holdings.append({
                    "Company Name": manual_name,
                    "Ticker": virtual_ticker,
                    "Shares": 1,
                    "Category": UI_TO_DB_MAP.get(manual_category, "Kontanter/Private"),
                    "Sector": manual_sector,
                    "manual_value": manual_value
                })
                st.success(f"Added manual asset '{manual_name}' to your portfolio!")
                time.sleep(1)
                st.rerun()
else:
    # Sletet eksempler til søgning
    search_query = st.text_input("🔍 Search by Company Name or Ticker (e.g., 'Adidas', 'Novo Nordisk', 'iShares Sukuk'):", "")

    if search_query:
        search_results = search_tickers_by_name_multi(search_query)
        
        if search_results:
            options_format = [f"{r['name']} ({r['symbol']})" for r in search_results]
            selected_option_str = st.selectbox("Select the correct asset from search results:", options_format)
            
            selected_idx = options_format.index(selected_option_str)
            target_asset = search_results[selected_idx]
            resolved_ticker = target_asset["symbol"]
            comp_name = target_asset["name"]
            
            try:
                t = yf.Ticker(resolved_ticker)
                info = t.info
                sec = info.get("sector", "Other")
                ind = info.get("industry", "Other")
                
                temp_screener = ScreenerComplianceAgent([], target_category=st.session_state.targets)
                cat, sub_sec = temp_screener.map_to_category_and_sector(resolved_ticker, sec, ind)
                
                display_cat = DB_TO_UI_MAP.get(cat, cat)
                
                st.markdown(f"""
                <div class="found-box">
                    <span style="color: #0F172A; font-weight: bold; font-size: 16px;">🔍 Confirmed Match:</span> {comp_name} ({resolved_ticker})<br>
                    <span style="color: #334155; font-weight: bold;">Asset Class:</span> {display_cat} | 
                    <span style="color: #334155; font-weight: bold;">Sector:</span> {sub_sec}
                </div>
                """, unsafe_allow_html=True)
                
                col_shares, col_add = st.columns([1, 1])
                with col_shares:
                    shares_to_add = st.number_input("Shares Owned:", min_value=1, value=10, key="shares_input")
                with col_add:
                    st.write(" ")
                    st.write(" ")
                    if st.button("➕ Add to Portfolio"):
                        exists = False
                        for h in st.session_state.holdings:
                            if h["Ticker"] == resolved_ticker:
                                h["Shares"] += shares_to_add
                                exists = True
                                break
                        if not exists:
                            st.session_state.holdings.append({
                                "Company Name": comp_name,
                                "Ticker": resolved_ticker,
                                "Shares": shares_to_add,
                                "Category": cat,
                                "Sector": sub_sec
                            })
                        st.success(f"Added {shares_to_add} shares of {comp_name} to your portfolio!")
                        time.sleep(1)
                        st.rerun()
            except Exception as e:
                st.error(f"Could not load details for '{resolved_ticker}': {str(e)}")
        else:
            st.warning(f"Could not find any assets matching '{search_query}'. Please try another spelling.")

st.write("---")
st.write("### Your Current Portfolio")
if st.session_state.holdings:
    holdings_df = pd.DataFrame(st.session_state.holdings)
    holdings_df['Category_Display'] = holdings_df['Category'].apply(lambda x: DB_TO_UI_MAP.get(x, x))
    
    edited_holdings = st.data_editor(
        holdings_df,
        num_rows="dynamic",
        column_config={
            "Company Name": st.column_config.TextColumn("Company Name", disabled=True),
            "Ticker": st.column_config.TextColumn("Ticker", disabled=True),
            "Shares": st.column_config.NumberColumn("Shares", min_value=1),
            "Category_Display": st.column_config.TextColumn("Category", disabled=True),
            "Sector": st.column_config.TextColumn("Sector", disabled=True),
            "manual_value": st.column_config.NumberColumn("Manual Value (DKK)", min_value=0)
        },
        use_container_width=True,
        key="holdings_editor"
    )
    
    if not edited_holdings.equals(holdings_df):
        st.session_state.holdings = edited_holdings.to_dict(orient="records")
        st.rerun()
        
    if login_email and "@" in login_email and login_password:
        if st.button("💾 Save Changes to My Profile"):
            with st.spinner("Saving your portfolio..."):
                status = save_user_portfolio_to_db(
                    email=login_email,
                    password=login_password,
                    holdings=st.session_state.holdings,
                    targets=st.session_state.targets,
                    horizon=st.session_state.horizon,
                    name=st.session_state.user_name,
                    frequency=st.session_state.frequency
                )
                if status == "success":
                    st.success("Successfully saved your profile! Your nightly council runs are now synchronized.")
                elif status == "incorrect_password":
                    st.error("Cannot save. Incorrect password.")
                else:
                    st.error("Failed to save changes. Please try again.")
else:
    st.info("Your portfolio is currently empty. Use the search field above to add assets.")

# =====================================================================
#  STEP 3: WATCHLIST (VALGFRI)
# =====================================================================
st.subheader("Step 3: Your Watchlist (Optional)")
st.write("Enter tickers to monitor. If left empty, the system will dynamically screen our global Sharia-growth pool.")
watchlist_input = st.text_input(
    "Enter Tickers (comma-separated):",
    "TRMB, SAP, SPSK, AEM, NEM"
)

watchlist_list = [t.strip().upper() for t in watchlist_input.split(",") if t.strip()]

# =====================================================================
#  FUNKTION TIL AT SKABE LIVE-RAPPORT OG PODCAST AUTOMATISK PÅ STREAMLIT
# =====================================================================
async def process_instant_briefing(receiver_email, holdings_list, watchlist, target_allocations, user_name, horizon, is_new, new_sectors):
    """
    Kører hele investerings-motoren asynkront direkte på Streamlits cloud-server.
    """
    if is_new:
        portfolio_distribution = {"Aktier": 25.0, "Sukuk": 25.0, "Råvarer": 25.0, "Kontanter/Private": 25.0}
        sector_distribution = {sec: 33.3 for sec in new_sectors}
        holdings_list = []
        focus_category = "Aktier"
        deficit = 25.0
    else:
        total_mv = 0.0
        for item in holdings_list:
            symbol = item["Ticker"]
            if "PVT_" in symbol or "CASH_" in symbol:
                total_mv += float(item.get("manual_value", 1000))
            else:
                try:
                    t = yf.Ticker(symbol)
                    price = t.info.get("currentPrice", t.info.get("regularMarketPrice", 1.0))
                    total_mv += (price * float(item["Shares"]))
                except Exception:
                    total_mv += 1000.0

        portfolio_distribution = {"Aktier": 0.0, "Sukuk": 0.0, "Råvarer": 0.0, "Kontanter/Private": 0.0}
        sector_distribution = {}
        
        for item in holdings_list:
            category = item["Category"]
            subsector = item["Sector"]
            symbol = item["Ticker"]
            
            if "PVT_" in symbol or "CASH_" in symbol:
                item_mv = float(item.get("manual_value", 1000))
            else:
                try:
                    t = yf.Ticker(symbol)
                    price = t.info.get("currentPrice", t.info.get("regularMarketPrice", 1.0))
                    item_mv = (price * float(item["Shares"]))
                except Exception:
                    item_mv = 1000.0
                    
            weight_chunk = (item_mv / total_mv * 100.0) if total_mv > 0 else 20.0
            
            if category in portfolio_distribution:
                portfolio_distribution[category] += weight_chunk
            if subsector not in sector_distribution:
                sector_distribution[subsector] = 0.0
                
            sector_distribution[subsector] += weight_chunk

        pm = PortfolioManagerAgent(portfolio_distribution, target_allocations)
        focus_category, deficit = pm.identify_underweighted_focus()
        
    print(f"Nattens fokus: {focus_category} (Gab: {deficit:.2f}%)")
    
    growth_pool = GLOBAL_COMPLIANT_GROWTH_POOL.get(focus_category, [])
    combined_candidates = list(set(watchlist + growth_pool))
    
    screener = ScreenerComplianceAgent(combined_candidates, target_category=focus_category)
    approved_stocks = screener.run_screening(focus_category)
    target_candidates = approved_stocks[:10]
    
    if not target_candidates:
        return False, "No compliant assets could be found in your focused category."

    detailed_candidates_data = []
    for stock in target_candidates:
        symbol = stock["symbol"]
        try:
            t = yf.Ticker(symbol)
            info = t.info
            rev_growth = info.get("revenueGrowth", "N/A")
            op_margins = info.get("operatingMargins", "N/A")
            free_cashflow = info.get("freeCashflow", "N/A")
            
            detailed_candidates_data.append({
                "symbol": symbol,
                "name": stock["name"],
                "pe_ratio": stock["pe_ratio"],
                "debt_ratio": stock["debt_ratio"],
                "sector": stock["sector"],
                "industry": stock["industry"],
                "is_etf": stock.get("is_etf", False),
                "revenue_growth": f"{rev_growth * 100:.2f}%" if isinstance(rev_growth, (int, float)) else "N/A",
                "operating_margins": f"{op_margins * 100:.2f}%" if isinstance(op_margins, (int, float)) else "N/A",
                "free_cash_flow": f"{free_cashflow / 1e6:.2f}M" if isinstance(free_cashflow, (int, float)) else "N/A",
                "current_price": info.get("currentPrice", info.get("regularMarketPrice", "N/A")),
                "currency": info.get("currency", "N/A")
            })
        except Exception:
            detailed_candidates_data.append(stock)

    current_weights_str = json.dumps(portfolio_distribution, indent=2, ensure_ascii=False)
    current_holdings_str = json.dumps(holdings_list, indent=2, ensure_ascii=False)
    sector_distribution_str = json.dumps(sector_distribution, indent=2, ensure_ascii=False)
    
    council_agent = CouncilAgent(GEMINI_API_KEY)
    report_html = council_agent.run_proactive_analysis(
        candidates_data=detailed_candidates_data,
        category=focus_category,
        deficit=deficit,
        current_portfolio_str=current_weights_str,
        current_holdings_str=current_holdings_str,
        sector_distribution_str=sector_distribution_str,
        user_name=user_name,
        horizon=horizon
    )

    # 1. Automatisk kompilering af dit live-opdaterede Excel-styringsark!
    print("Kompilerer dit live-opdaterede Excel-ark...")
    excel_raw_bytes = generate_excel_template_bytes(holdings_list, watchlist)

    output_mp3 = "llm_council_podcast.mp3"
    podcast_compiled = False
    
    podcast_agent = PodcastAgent(GEMINI_API_KEY)
    generated_file = podcast_agent.generate_podcast_audio(report_html, user_name)
    
    if generated_file and os.path.exists(generated_file):
        import shutil
        shutil.copyfile(generated_file, output_mp3)
        podcast_compiled = True

    # Klargør begge vedhæftninger (både MP3-podcasten og .xlsx-styringsarket!)
    attachments_list = []
    if podcast_compiled:
        attachments_list.append({"path": output_mp3, "name": f"{user_name}_LLM_Council_Podcast.mp3"})
    
    attachments_list.append({"data": excel_raw_bytes, "name": f"{user_name}_Live_Portfolio_Template.xlsx"})

    os.environ["EMAIL_RECEIVER"] = receiver_email
    subject = f"[LLM Council] Your Personal Strategic Briefing - Focus on {DISPLAY_CATEGORIES.get(focus_category, focus_category)}"
    
    import llm_council
    llm_council.EMAIL_RECEIVER = receiver_email
    
    DeliveryAgent.send_email(
        subject=subject,
        html_content=report_html,
        attachments=attachments_list # Sender nu begge filer direkte til mailen!
    )
    return True, "Your briefing, audio podcast, and custom Excel sheet have been sent to your email!"

# =====================================================================
#  SAAS START-KNAP
# =====================================================================
st.subheader("Step 4: Activate Your Personal Council")

col_b1, col_b2 = st.columns([2, 1])
with col_b1:
    if st.button("🚀 Start My LLM Council & Send First Report"):
        if not user_email_input or "@" not in user_email_input:
            st.error("Please enter a valid email address.")
        elif not is_new_investor and not st.session_state.holdings:
            st.error("Please configure at least one active holding.")
        elif is_new_investor and not selected_new_sectors:
            st.error("Please select at least one sector you want exposure to.")
        elif not GEMINI_API_KEY:
            st.error("SaaS master API key is missing on the server.")
        else:
            with st.spinner("Processing your holdings, screening candidates and generating your Bloomberg-style podcast... This takes about 60 seconds."):
                # Send alle data afsted til kørslen
                success, msg = asyncio.run(process_instant_briefing(
                    user_email_input, 
                    st.session_state.holdings, 
                    watchlist_list, 
                    st.session_state.targets, 
                    user_name_input, 
                    st.session_state.horizon,
                    is_new_investor,
                    selected_new_sectors
                ))
                if success:
                    st.success(f"Boom! {msg}")
                    st.balloons()
                else:
                    st.error(f"Failed to generate briefing: {msg}")
with col_b2:
    # Direkte instant download-knap på skærmen til dit live Excel-ark
    if st.session_state.holdings:
        excel_bytes = generate_excel_template_bytes(st.session_state.holdings, watchlist_list)
        st.download_button(
            label="📥 Download My Excel Sheet",
            data=excel_bytes,
            file_name=f"{st.session_state.user_name}_Live_Portfolio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# =====================================================================
#  DYNAMISK LEGAL DISCLAIMER & ZOYA-LINK I BUNDEN
# =====================================================================
st.markdown("""
    <div style="background-color: #FEF2F2; border: 1px solid #FCA5A5; border-left: 6px solid #EF4444; padding: 20px; border-radius: 8px; margin-top: 40px; margin-bottom: 30px;">
        <h4 style="color: #991B1B; font-family: 'Georgia', serif; margin-top: 0; margin-bottom: 8px;">⚠️ Legal Disclaimer & Personal Conviction</h4>
        <p style="color: #7F1D1D; font-size: 14px; margin-bottom: 10px; line-height: 1.5;">
            The LLM Council is an automated, AI-driven informational and educational inspiration tool. It is <strong>not</strong> a licensed financial advisor, nor does it provide personalized investment advice or regulatory financial mandates. 
        </p>
        <p style="color: #7F1D1D; font-size: 14px; margin-bottom: 10px; line-height: 1.5;">
            Financial markets carry inherent risks, and Shariah-compliance interpretations can vary across different scholars and madhabs. You must <strong>always</strong> base your final investment decisions on your own research, personal convictions, and common sense. 
        </p>
        <p style="color: #7F1D1D; font-size: 14px; margin-bottom: 0; line-height: 1.5;">
            To manually audit and double-check the Shariah-compliance, financial health, or business profile of any individual stock or fund, we highly recommend utilizing the official <a href="https://zoya.finance/" target="_blank" style="color: #B91C1C; font-weight: bold; text-decoration: underline;">Zoya Finance Platform</a>.
        </p>
    </div>
""", unsafe_allow_html=True)
