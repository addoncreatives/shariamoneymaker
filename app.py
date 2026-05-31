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

# Sæt sidens titel og layout for et moderne look
st.set_page_config(page_title="LLM Council - Conscious Wealth", page_icon="🗳️", layout="centered")


# =====================================================================
#  HJÆLPEFUNKTIONER TIL HTML & TOSPROGET SUPPORT (Live-oversættelse)
# =====================================================================
def render_html(html_str: str):
    """
    Renser HTML-strengen for linjeskift og indrykninger, så Streamlits
    Markdown-parser aldrig forveksler HTML med en rå kodeblok.
    """
    clean_html = "".join([line.strip() for line in html_str.splitlines()])
    st.markdown(clean_html, unsafe_allow_html=True)


# Sprogvælger placeret øverst til højre
col_title, col_lang = st.columns([3, 1])
with col_title:
    st.title("🗳️ LLM Council")
with col_lang:
    if "lang" not in st.session_state:
        st.session_state.lang = "Dansk"
    st.session_state.lang = st.selectbox(
        "Sprog / Language", 
        ["Dansk", "English"], 
        index=0 if st.session_state.lang == "Dansk" else 1,
        label_visibility="collapsed"
    )

def _t(da_text: str, en_text: str) -> str:
    """Hjælpefunktion til live-oversættelse af UI-tekster."""
    return da_text if st.session_state.lang == "Dansk" else en_text


# =====================================================================
#  CONFIGURATION & STANDARD TARGET WEIGHTS
# =====================================================================

DISPLAY_CATEGORIES = {
    "Aktier": "Equities",
    "Sukuk": "Sukuk (Islamic Bonds)",
    "Råvarer": "Commodities",
    "Kontanter/Private": "Cash / Private Sector"
}

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

# DET STATISKE LYNHURTIGE KARTOTEK
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

@st.cache_data
def load_global_db_from_github():
    url = "https://raw.githubusercontent.com/addoncreatives/shariamoneymaker/main/failsafe_db.json"
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            loaded_db = response.json()
            merged_db = STATIC_TICKER_MAP.copy()
            merged_db.update(loaded_db)
            return merged_db
    except Exception:
        pass
    return STATIC_TICKER_MAP

failsafe_db = load_global_db_from_github()

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

EMAIL_SENDER = os.getenv("EMAIL_SENDER", "wazir.ilyas@gmail.com")
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER", "addoncreatives@gmail.com")
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
    headers = {'User-Agent': 'Mozilla/5.0'}
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
    except Exception:
        pass
    return []


# =====================================================================
#  EXCEL GENERATOR (OPDATERET MED RIGTIGE VÆRDIER & DETALJER)
# =====================================================================
def generate_excel_template_bytes(holdings_list: list, watchlist_list: list, portfolio_weights: dict = None, sector_distribution: dict = None) -> bytes:
    wb = openpyxl.Workbook()
    
    # 1. Beholdninger
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
        shares = item.get("Shares", 1)
        kurs = item.get("Kurs", 0.0)
        cat = item.get("Category", "Aktier")
        sec = item.get("Sector", "Other")
        
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
            
            if kurs > 0.0:
                ws1.cell(row=idx, column=5, value=kurs)
            else:
                ws1.cell(row=idx, column=5, value=f'=GOOGLEFINANCE(B{idx})')
                
            ws1.cell(row=idx, column=6, value=f'=D{idx}*E{idx}')
            
        ws1.cell(row=idx, column=7, value=cat)
        ws1.cell(row=idx, column=8, value="")
        ws1.cell(row=idx, column=9, value=sec)
        ws1.cell(row=idx, column=10, value="Global")
        ws1.cell(row=idx, column=11, value=f'=F{idx}/SUM(F$2:F$100)')
        ws1.cell(row=idx, column=12, value="")
        ws1.cell(row=idx, column=13, value="")

    # 2. Opsummering
    ws2 = wb.create_sheet(title="Opsummering")
    ws2.cell(row=1, column=1, value="4x25-overblik")
    ws2.cell(row=1, column=6, value="Økonomiske drivere")
    ws2.cell(row=1, column=10, value="Sektorer")
    ws2.cell(row=1, column=14, value="Huller / Watchlist")
    
    ws2.cell(row=2, column=1, value="Kategori")
    ws2.cell(row=2, column=2, value="Aktuel Vægt (%)")
    ws2.cell(row=2, column=3, value="Mål Vægt (%)")
    
    ws2.cell(row=2, column=10, value="Delsektor")
    ws2.cell(row=2, column=11, value="Vægt (%)")
    ws2.cell(row=2, column=14, value="Ticker")

    if portfolio_weights:
        categories = ["Aktier", "Sukuk", "Råvarer", "Kontanter/Private"]
        for i_idx, cat in enumerate(categories, start=3):
            ws2.cell(row=i_idx, column=1, value=cat)
            ws2.cell(row=i_idx, column=2, value=portfolio_weights.get(cat, 0.0))
            ws2.cell(row=i_idx, column=3, value=25.0)

    if sector_distribution:
        for s_idx, (sec, weight) in enumerate(sector_distribution.items(), start=3):
            ws2.cell(row=s_idx, column=10, value=sec)
            ws2.cell(row=s_idx, column=11, value=weight)

    for w_idx, ticker in enumerate(watchlist_list, start=3):
        ws2.cell(row=w_idx, column=14, value=ticker)
        
    excel_data = io.BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)
    return excel_data.getvalue()


def get_category_and_sector_failsafe(ticker: str, target_category: str = None) -> tuple:
    sym = str(ticker).upper().strip()
    lookup_sym = sym.split('.')[0]
    
    for k, v in failsafe_db.items():
        if normalize_string(k) == normalize_string(sym) or normalize_string(k) == normalize_string(lookup_sym):
            if v[0] == "Sukuk" and target_category == "Kontanter/Private":
                return "Kontanter/Private", "Sukuk & Fixed Income"
            return v[0], v[1]
            
    try:
        t = yf.Ticker(sym)
        info = t.info
        sec = info.get("sector", "Other")
        ind = info.get("industry", "Other")
        
        temp_screener = ScreenerComplianceAgent([], target_category=target_category)
        cat, sub_sec = temp_screener.map_to_category_and_sector(sym, sec, ind)
        return cat, sub_sec
    except Exception:
        return "Aktier", "Other"


# =====================================================================
#  SÆRLIG SCREENING & AGENTER
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


class ScreenerComplianceAgent:
    PROHIBITED_SECTORS = ["Financial Services", "Financial"]
    PROHIBITED_INDUSTRIES = ["Banks", "Insurance", "Aerospace & Defense", "Gambling", "Tobacco", "Distillers & Vintners", "Breweries"]

    def __init__(self, tickers: list, target_category: str = None):
        self.tickers = tickers
        self.target_category = target_category

    def check_zoya_live_compliance(self, symbol: str) -> bool:
        clean_symbol = symbol.split('.')[0].upper().split('-')[0]
        url = f"https://zoya.finance/stocks/{clean_symbol}"
        headers = {'User-Agent': 'Mozilla/5.0'}
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

        lookup_sym = sym.split('.')[0]
        for k, v in failsafe_db.items():
            if normalize_string(k) == normalize_string(sym) or normalize_string(k) == normalize_string(lookup_sym):
                if v[0] == "Sukuk" and self.target_category == "Kontanter/Private":
                    return "Kontanter/Private", "Sukuk & Fixed Income"
                return v[0], v[1]

        if "sukuk" in sym or sym in ["SPSK", "SKUK"]:
            if self.target_category == "Kontanter/Private":
                return "Kontanter/Private", "Sukuk & Fixed Income"
            return "Sukuk", "Sukuk & Fixed Income"
            
        if sym in ["WPM", "FNV", "RGLD"]:
            return "Råvarer", "Mining & Royalty Streams"
        if sym in ["NEM", "GOLD", "AEM", "BHP", "RIO", "FCX", "VALE"] or any(w in ind_l for w in ["gold", "silver", "precious metals", "copper", "aluminum"]):
            return "Råvarer", "Industrial Metals & Copper"

        if "cash" in sym or "money market" in sec_l:
            return "Kontanter/Private", "Cash & Liquidity Reserves"

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

        dynamic_subsector = industry if (industry and industry != "Other") else sector
        return "Aktier", dynamic_subsector

    def screen_ticker(self, symbol: str) -> dict:
        try:
            if "CASH_" in symbol or "PVT_" in symbol:
                return {
                    "symbol": symbol, "passed": True, "name": symbol.replace("CASH_", "").replace("PVT_", ""),
                    "pe_ratio": "N/A", "debt_ratio": "0.00% (Manual)", "sector": "Manual Asset",
                    "industry": "Manual Asset", "category": "Kontanter/Private", "subsector": "Cash & Liquidity Reserves", "is_etf": False
                }

            zoya_compliant = self.check_zoya_live_compliance(symbol)
            if zoya_compliant is False:
                return {"symbol": symbol, "passed": False, "reason": "Disqualified by Zoya assessment."}

            ticker_obj = yf.Ticker(symbol)
            info = ticker_obj.info
            if not info:
                return {"symbol": symbol, "passed": False, "reason": "Ingen data"}

            quote_type = info.get("quoteType", "").upper()
            is_etf = quote_type in ["ETF", "MUTUALFUND"] or symbol in ["IGDA.L", "SPSK", "HLAL", "UMMA", "ISWD.L", "MSAU.L", "SKUK"]

            if is_etf:
                mapped_cat, mapped_sub = self.map_to_category_and_sector(symbol, "ETF", "ETF")
                return {
                    "symbol": symbol, "passed": True, "name": info.get("longName", symbol),
                    "pe_ratio": info.get("trailingPE", "N/A"), "debt_ratio": "N/A (ETF/Sukuk)",
                    "sector": "ETF / Fond", "industry": "ETF", "category": mapped_cat, "subsector": mapped_sub, "is_etf": True
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
                "symbol": symbol, "passed": True, "name": info.get("longName", symbol),
                "pe_ratio": info.get("trailingPE", "N/A"), "debt_ratio": f"{debt_ratio_pct:.2f}% ({method_used})",
                "sector": sector, "industry": industry, "category": mapped_cat, "subsector": mapped_sub, "is_etf": False
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
#  COUNCIL AGENT
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
        - Investment Horizon: {horizon}
        - Focus tonight: {english_category} (Current Deficit: {deficit:.2f}%)
        - Current Portfolio Allocations: {current_portfolio_str}
        - Sub-sectors: {sector_distribution_str}
        - Current Holdings: {current_holdings_str}
        - Screened candidates: {candidates_json}
        
        YOUR OBJECTIVE:
        Generate a complete, institutional-grade, highly engaging HTML investment newsletter in English with inline CSS styling.
        Refer to AAOIFI ethical standards (under 30% debt rule). Avoid extreme financial jargon. Frame recommendations respectfully.
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
        except Exception as e:
            return f"<h3>System Error</h3><p>{str(e)}</p>"


# =====================================================================
#  PODCAST AGENT (KORRIGERET TIL MULTI-ADVISOR SYNERGY OG RETTEDE XML-TAGS)
# =====================================================================
class PodcastAgent:
    def __init__(self, api_key: str):
        self.api_key = api_key
        if self.api_key:
            os.environ["GEMINI_API_KEY"] = self.api_key
            os.environ["GOOGLE_API_KEY"] = self.api_key

    def generate_podcast_audio(self, report_html: str, user_name: str) -> str:
        from podcastfy.client import generate_podcast
        
        user_instructions_content = (
            f"Set up a high-energy educational narrative podcast. "
            f"Explain to the listener our VIP client {user_name}'s current portfolio status and macro investment options. "
            f"Sara (Person 1) acts as the primary host. She must present the global macro events (such as supply chains or central banks) "
            f"and then discuss 2-3 Shariah-compliant candidates. "
            f"To present a structured critique, Sara must summarize the views of three virtual advisors using their names and titles:\n"
            f"- David, Chief Strategist at a global short-fund (skeptical and cautious)\n"
            f"- Michael, CIO of a Silicon Valley growth fund (enthusiastic about pipelines and growth)\n"
            f"- Elena, Head of Macro Research at a Swiss investment bank (focused on Shariah compliance limits and asset balance)\n\n"
            f"Marcus (Person 2), the Investment Committee Chairman, must join Sara in the second half of the episode. "
            f"Marcus must act as the Executor, explaining the practical tradeability on Saxo Bank and recommending a calm, step-by-step strategy. "
            f"Both hosts must summarize the final Masterclass verdict, keep 'number noise' to an absolute minimum, and make it educational and story-driven."
        )

        custom_config = {
            "word_count": 1100,
            "conversation_style": ["educational", "professional", "highly engaging", "storytelling", "structured dialogue"],
            "roles_person1": "Sara, the sharp financial journalist and solo host",
            "roles_person2": "Marcus, the wise Investment Committee Chairman",
            "podcast_name": "The Investor's Journey",
            "podcast_tagline": "Your money, your journey, your Shariah-compliant future",
            "output_language": "English",
            "engagement_techniques": ["rhetorical questions", "analogies", "humor", "interjections"],
            "user_instructions": user_instructions_content
        }
        
        try:
            print("Genererer struktureret podcast (Sara & Marcus) på engelsk...")
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
#  DELIVERY AGENT
# =====================================================================
class DeliveryAgent:
    @staticmethod
    def send_email(subject: str, html_content: str, attachments: list = None):
        if not EMAIL_PASSWORD or EMAIL_PASSWORD.strip() == "":
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
                    part.add_header("Content-Disposition", f"attachment; filename= {name}")
                    msg.attach(part)

        try:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
            server.quit()
        except Exception:
            pass


# =====================================================================
#  FUNKTION TIL AT SKABE LIVE-RAPPORT OG PODCAST AUTOMATISK PÅ STREAMLIT
# =====================================================================
async def process_instant_briefing(receiver_email, holdings_list, watchlist, target_allocations, user_name, horizon, is_new, new_sectors):
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
                    item["Kurs"] = price
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
        
    print(f"SaaS Fokus: {focus_category} (Gab: {deficit:.2f}%)")
    
    growth_pool = GLOBAL_COMPLIANT_GROWTH_POOL.get(focus_category, [])
    combined_candidates = list(set(watchlist + growth_pool))
    
    screener = ScreenerComplianceAgent(combined_candidates, target_category=focus_category)
    approved_stocks = screener.run_screening(focus_category)
    target_candidates = approved_stocks[:10]
    
    if not target_candidates:
        return False, "Ingen selskaber bestod screening i denne kategori i dag."

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

    excel_raw_bytes = generate_excel_template_bytes(holdings_list, watchlist, portfolio_distribution, sector_distribution)

    output_mp3 = "llm_council_podcast.mp3"
    podcast_compiled = False
    
    # Opret en ren tekstbeskrivelse til podcast-værterne i stedet for tung e-mail HTML
    clean_data_text = (
        f"VIP Client: {user_name}\n"
        f"Portfolio allocations (Actual vs Target):\n{current_weights_str}\n\n"
        f"Dynamic sectors:\n{sector_distribution_str}\n\n"
        f"Current holdings:\n{current_holdings_str}\n\n"
        f"Target focus category tonight: {focus_category} (deficit: {deficit:.2f}%)\n\n"
        f"Screened compliant stocks:\n"
    )
    for s in detailed_candidates_data:
        clean_data_text += f"- {s.get('name')} ({s.get('symbol')}): Sector: {s.get('sector')}, Industry: {s.get('industry')}. Price: {s.get('current_price')} {s.get('currency')}. Revenue Growth: {s.get('revenue_growth')}, Operating Margin: {s.get('operating_margins')}, FCF: {s.get('free_cash_flow')}. Debt Ratio: {s.get('debt_ratio')}.\n"

    podcast_agent = PodcastAgent(GEMINI_API_KEY)
    generated_file = podcast_agent.generate_podcast_audio(clean_data_text, user_name)
    
    if generated_file and os.path.exists(generated_file):
        import shutil
        shutil.copyfile(generated_file, output_mp3)
        podcast_compiled = True

    attachments_list = []
    if podcast_compiled:
        attachments_list.append({"path": output_mp3, "name": f"{user_name}_LLM_Council_Podcast.mp3"})
    attachments_list.append({"data": excel_raw_bytes, "name": f"{user_name}_Live_Portfolio_Template.xlsx"})

    os.environ["EMAIL_RECEIVER"] = receiver_email
    subject = f"[LLM Council] Your Personal Strategic Briefing - Focus on {DISPLAY_CATEGORIES.get(focus_category, focus_category)}"
    
    DeliveryAgent.send_email(
        subject=subject,
        html_content=report_html,
        attachments=attachments_list
    )
    return True, "Briefing, lyd-podcast og dit Excel-ark er nu sendt til din e-mailadresse."


# =====================================================================
#  DATABASE INTEGRATION (GOOGLE WEB APP)
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
    except Exception:
        pass
    return None

def save_user_portfolio_to_db(email: str, password: str, holdings: list, targets: dict, horizon: str, name: str, frequency: str) -> str:
    if not DATABASE_URL or DATABASE_URL.strip() == "":
        return "no_db"
    payload = {
        "email": email, "password": password, "holdings": holdings,
        "targets": targets, "horizon": horizon, "name": name, "frequency": frequency
    }
    try:
        response = requests.post(DATABASE_URL, json=payload, timeout=15)
        if response.status_code == 200:
            return response.json().get("status")
    except Exception:
        pass
    return "error"


# =====================================================================
#  SESSION STATE INITIALIZATION (MED LÅSTE SLIDER-VARIABLER)
# =====================================================================
if "step" not in st.session_state:
    st.session_state.step = 1
if "investor_holdings" not in st.session_state:
    st.session_state.investor_holdings = []
if "targets" not in st.session_state:
    st.session_state.targets = {"Aktier": 25.0, "Sukuk": 25.0, "Råvarer": 25.0, "Kontanter/Private": 25.0}
if "horizon" not in st.session_state:
    st.session_state.horizon = "7-15 years"
if "user_name" not in st.session_state:
    st.session_state.user_name = "Investor"
if "frequency" not in st.session_state:
    st.session_state.frequency = "Weekly"
if "user_email" not in st.session_state:
    st.session_state.user_email = ""
if "is_logged_in" not in st.session_state:
    st.session_state.is_logged_in = False
if "is_new_investor" not in st.session_state:
    st.session_state.is_new_investor = False

# Låste variabler til slider-hukommelsen
if "slider_stocks" not in st.session_state:
    st.session_state.slider_stocks = int(st.session_state.targets.get("Aktier", 25.0))
if "slider_sukuk" not in st.session_state:
    st.session_state.slider_sukuk = int(st.session_state.targets.get("Sukuk", 25.0))
if "slider_commodities" not in st.session_state:
    st.session_state.slider_commodities = int(st.session_state.targets.get("Råvarer", 25.0))
if "slider_cash" not in st.session_state:
    st.session_state.slider_cash = int(st.session_state.targets.get("Kontanter/Private", 25.0))


# =====================================================================
#  STREAMLIT UI STEPS
# =====================================================================

# Diskret status-indikator i toppen (erstatter den gamle stepper-menu)
st.caption(_t(f"Trin {st.session_state.step} af 5", f"Step {st.session_state.step} of 5"))


# --- TRIN 1: VELKOMST & KORT KONTEKST (BIBEHOLDER DIN ORIGINALE DARK MODE STYLING UDEN FASTE BAGGRUNDSFARVER) ---
if st.session_state.step == 1:
    # Komprimeret og lynhurtigt læst onboardingtekst, tilpasset Dark Mode
    render_html(_t("""
<div style="border: 1px solid #E2E8F0; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
    <h3 style="font-family: 'Georgia', serif; margin-top: 0;">🛡️ En gennemsigtig guide for den bevidste investor</h3>
    <p style="margin-bottom: 12px;">
        Som muslimsk investor i Norden (Saxo Bank, Nordnet, Avanza osv.) har du ikke adgang til automatiserede integrationsløsninger. 
        Det tvinger dig ofte ud i tidskrævende, manuel screening af Shariah-gældsregler (AAOIFI-standarder) og selskabshistorier.
    </p>
    <p style="margin-bottom: 15px;">
        <strong>LLM Council er din selvhjulpne makker.</strong> Vi tilbyder ikke finansiel rådgivning eller låste AI-beslutninger. Vi hjælper dig med at opbygge og rebalancere en robust, diversificeret portefølje fordelt på fire søjler: <strong>Equities</strong> (Aktier), <strong>Sukuk</strong> (Islamiske Certifikater), <strong>Commodities</strong> (Råvarer) og <strong>Cash/Private</strong>.
    </p>
</div>
""", """
<div style="border: 1px solid #E2E8F0; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
    <h3 style="font-family: 'Georgia', serif; margin-top: 0;">🛡️ A Transparent Guide for the Conscious Investor</h3>
    <p style="margin-bottom: 12px;">
        As a Muslim investor in the Nordics (Saxo Bank, Nordnet, Avanza, etc.), you lack automated integration solutions. 
        This often forces you into time-consuming, manual screening of Shariah debt rules (AAOIFI standards) and business models.
    </p>
    <p style="margin-bottom: 15px;">
        <strong>LLM Council is your self-reliant companion.</strong> We do not offer financial advice or locked AI decisions. We help you build and rebalance a robust, diversified portfolio across four pillars: <strong>Equities</strong>, <strong>Sukuk</strong> (Islamic Bonds), <strong>Commodities</strong>, and <strong>Cash/Private</strong>.
    </p>
</div>
"""))
    
    # Stor, klar knap rykket højt op for nem betjening på telefonen
    start_btn_text = _t("Opsæt dit porteføljestyringsværktøj ➔", "Setup your portfolio management tool ➔")
    if st.button(start_btn_text, use_container_width=True):
        st.session_state.step = 2
        st.rerun()


# --- TRIN 2: LOGIN & PROFILOPRETTELSE (RETTET: LOG-IND KNAP TILFØJET!) ---
elif st.session_state.step == 2:
    st.subheader(_t("Opret din profil eller log ind", "Create your profile or sign in"))
    st.write(_t("Dine oplysninger gemmes sikkert, så din personlige portefølje og dine ugentlige briefinger synkroniseres automatisk.", "Your details are stored securely so your personal portfolio and weekly briefings sync automatically."))
    
    col_l1, col_l2 = st.columns(2)
    with col_l1:
        login_email = st.text_input("E-mail", placeholder="navn@gmail.com", value=st.session_state.user_email)
    with col_l2:
        login_password = st.text_input(_t("Adgangskode", "Password"), type="password")

    is_new_user = False
    db_profile = None

    # Tydelig log-ind knap, så brugeren slipper for at skulle trykke "Enter" i passwordfeltet
    if st.button(_t("Log ind ➔", "Log In ➔"), use_container_width=True):
        if login_email and "@" in login_email and login_password:
            with st.spinner(_t("Forbinder til profil...", "Connecting to profile...")):
                response = requests.get(f"{DATABASE_URL}?email={login_email}&password={login_password}", timeout=10)
                if response.status_code == 200:
                    res_data = response.json()
                    if res_data.get("status") == "success":
                        db_profile = res_data
                        st.session_state.investor_holdings = db_profile.get("holdings", [])
                        st.session_state.targets = db_profile.get("targets", {"Aktier": 25.0, "Sukuk": 25.0, "Råvarer": 25.0, "Kontanter/Private": 25.0})
                        
                        # Synkroniserer øjeblikkeligt sliderne til de indlæste værdier fra databasen
                        st.session_state.slider_stocks = int(st.session_state.targets.get("Aktier", 25.0))
                        st.session_state.slider_sukuk = int(st.session_state.targets.get("Sukuk", 25.0))
                        st.session_state.slider_commodities = int(st.session_state.targets.get("Råvarer", 25.0))
                        st.session_state.slider_cash = int(st.session_state.targets.get("Kontanter/Private", 25.0))
                        
                        st.session_state.horizon = db_profile.get("horizon", "7-15 years")
                        st.session_state.user_name = db_profile.get("name", "Investor")
                        st.session_state.frequency = db_profile.get("frequency", "Weekly")
                        st.session_state.user_email = login_email
                        st.session_state.is_logged_in = True
                        st.success(_t(f"Velkommen tilbage, {st.session_state.user_name}!", f"Welcome back, {st.session_state.user_name}!"))
                        time.sleep(1)
                        st.rerun()
                    elif res_data.get("status") == "incorrect_password":
                        st.error(_t("Forkert adgangskode for denne e-mailadresse.", "Incorrect password for this email address."))
                    elif res_data.get("status") == "not_found":
                        is_new_user = True
        else:
            st.error(_t("Udfyld venligst både e-mail og adgangskode.", "Please enter both email and password."))

    if is_new_user:
        st.info(_t("E-mailen blev ikke fundet. Udfyld felterne nedenfor for at oprette en ny profil:", "Email not found. Fill in the fields below to create a new profile:"))
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            confirm_password = st.text_input(_t("Bekræft adgangskode", "Confirm password"), type="password")
        with col_s2:
            signup_name = st.text_input(_t("Dit fulde navn", "Your full name"), value="Investor")
        
        if st.button(_t("📝 Registrer profil", "📝 Register profile"), use_container_width=True):
            if login_password != confirm_password:
                st.error(_t("Adgangskoderne matcher ikke!", "Passwords do not match!"))
            elif not signup_name:
                st.error(_t("Udfyld venligst dit navn.", "Please enter your name."))
            else:
                status = save_user_portfolio_to_db(
                    email=login_email,
                    password=login_password,
                    holdings=st.session_state.investor_holdings,
                    targets=st.session_state.targets,
                    horizon=st.session_state.horizon,
                    name=signup_name,
                    frequency=st.session_state.frequency
                )
                if status == "success":
                    st.session_state.user_name = signup_name
                    st.session_state.user_email = login_email
                    st.session_state.is_logged_in = True
                    st.success(_t("Profilen blev oprettet!", "Profile created!"))
                    st.session_state.step = 3
                    st.rerun()
                else:
                    st.error(_t("Fejl under oprettelsen. Prøv igen.", "Registration failed. Try again."))

    # Navigationsknapper
    st.write(" ")
    col_prev, col_next = st.columns(2)
    with col_prev:
        if st.button("⬅ " + _t("Tilbage", "Back"), use_container_width=True):
            st.session_state.step = 1
            st.rerun()
    with col_next:
        if st.session_state.is_logged_in:
            if st.button(_t("Næste trin", "Next step") + " ➔", use_container_width=True):
                st.session_state.step = 3
                st.rerun()


# --- TRIN 3: INVESTERINGSPROFIL & ALLOKERING (RETTET: INTERAKTIVE OG SIKRE SLIDERE!) ---
elif st.session_state.step == 3:
    st.subheader(_t("Definer din investeringsprofil", "Define your investment profile"))
    
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        st.session_state.user_name = st.text_input(_t("Dit navn i rapporten:", "Your name in the report:"), value=st.session_state.user_name)
    with col_n2:
        st.session_state.user_email = st.text_input(_t("E-mailadresse til briefinger:", "Email address for briefings:"), value=st.session_state.user_email)

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        horizon_options = ["1-3 years", "3-7 years", "7-15 years", "15+ years"]
        horizon_index = horizon_options.index(st.session_state.horizon) if st.session_state.horizon in horizon_options else 2
        st.session_state.horizon = st.selectbox(_t("Investeringshorisont:", "Investment horizon:"), horizon_options, index=horizon_index)
    with col_s2:
        freq_options = ["Daily", "Weekly", "Bi-weekly", "Monthly"]
        freq_index = freq_options.index(st.session_state.frequency) if st.session_state.frequency in freq_options else 1
        st.session_state.frequency = st.selectbox(_t("Hvor ofte ønsker du briefing?", "How often do you want briefings?"), freq_options, index=freq_index)

    st.write("---")
    st.subheader(_t("Angiv din ønskede mål-allokering", "Specify your target asset allocation"))
    st.write(_t("Træk i sliderne nedenfor. Din samlede vægtning låses automatisk, så den ALDRIG kan skyde over 100% samlet.", "Adjust the sliders below. Your total allocation is automatically capped, so it can NEVER exceed 100% in total."))
    
    # 1. Dynamisk loft-beregning (Gør det umuligt at skyde over 100% samlet!)
    st.session_state.slider_stocks = min(st.session_state.slider_stocks, 100)
    
    max_sukuk = 100 - st.session_state.slider_stocks
    st.session_state.slider_sukuk = min(st.session_state.slider_sukuk, max_sukuk)
    
    max_commodities = 100 - st.session_state.slider_stocks - st.session_state.slider_sukuk
    st.session_state.slider_commodities = min(st.session_state.slider_commodities, max_commodities)
    
    max_cash = 100 - st.session_state.slider_stocks - st.session_state.slider_sukuk - st.session_state.slider_commodities
    st.session_state.slider_cash = min(st.session_state.slider_cash, max_cash)

    # 2. Tegn sliderne med dynamiske max-værdier (Ingen resætninger!)
    target_stocks = st.slider(_t("Equities (Aktier) %", "Equities %"), 0, 100, key="slider_stocks")
    target_sukuk = st.slider(_t("Sukuk %", "Sukuk %"), 0, max_sukuk, key="slider_sukuk")
    target_commodities = st.slider(_t("Commodities (Råvarer) %", "Commodities %"), 0, max_commodities, key="slider_commodities")
    target_cash = st.slider(_t("Cash/Private %", "Cash/Private %"), 0, max_cash, key="slider_cash")

    # Beregn summen og giv intuitiv visuel feedback
    total_target = target_stocks + target_sukuk + target_commodities + target_cash
    
    if total_target != 100:
        difference = 100 - total_target
        st.warning(_t(
            f"⚠️ Allokeringen skal give 100% tilsammen. Nuværende sum: {total_target}%. Du mangler at fordele {difference}%.",
            f"⚠️ Allocation must equal 100% in total. Current sum: {total_target}%. You need to allocate {difference}%."
        ))
    else:
        st.success(_t("✅ Allokeringen er præcis 100%! Du kan nu fortsætte.", "✅ Allocation is exactly 100%! You can now proceed."))
        st.session_state.targets = {
            "Aktier": float(target_stocks),
            "Sukuk": float(target_sukuk),
            "Råvarer": float(target_commodities),
            "Kontanter/Private": float(target_cash)
        }

    # Navigationsknapper - Deaktiver næste trin, hvis mål-vægtene ikke giver 100%
    st.write(" ")
    col_prev, col_next = st.columns(2)
    with col_prev:
        if st.button("⬅ " + _t("Tilbage", "Back"), use_container_width=True):
            st.session_state.step = 2
            st.rerun()
    with col_next:
        st.button(
            _t("Næste trin", "Next step") + " ➔", 
            use_container_width=True, 
            disabled=(total_target != 100),
            key="next_to_step4"
        )
        if st.session_state.get("next_to_step4"):
            st.session_state.step = 4
            st.rerun()


# --- TRIN 4: PORTEFØLJEOPBYGNING (RETTET: INGEN FORUDVALGTE SEKTORE!) ---
elif st.session_state.step == 4:
    st.subheader(_t("Indtast dine nuværende aktiver", "Input your current holdings"))
    
    # Gemmes nu stabilt i session-state så det huskes mellem trin
    st.session_state.is_new_investor = st.checkbox(
        _t("Jeg er helt ny investor (starter fra bunden med tom portefølje)", "I am a completely new investor (starting from scratch with an empty portfolio)"),
        value=st.session_state.is_new_investor
    )

    selected_new_sectors = []
    if st.session_state.is_new_investor:
        # Expander-menu med simple afkrydsningsfelter løser problemet med at tastaturet popper op og menuen hopper
        with st.expander(_t("Vælg de sektorer du vil opbygge eksponering mod ▾", "Select the sectors you want to build exposure to ▾")):
            for sector in TARGET_SUBSECTORS:
                # RETTET: Sætter standardværdien til False, så ingen sektorer er forudvalgte!
                if st.checkbox(sector, value=False):
                    selected_new_sectors.append(sector)
    else:
        is_manual = st.checkbox(_t("Er dette et manuelt aktiv? (F.eks. kontantbeholdning, unoterede selskaber)", "Is this a manual asset? (e.g., cash, private equity)"))

        if is_manual:
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                manual_name = st.text_input(_t("Navn på aktiv:", "Asset name:"), placeholder="F.eks. Saxo Kontant DKK")
            with col_m2:
                manual_value = st.number_input(_t("Samlet værdi i DKK:", "Total value in DKK:"), min_value=1, value=1000)
                
            col_m3, col_m4 = st.columns(2)
            with col_m3:
                manual_category = st.selectbox(_t("Aktivklasse:", "Asset class:"), ["Cash/Private", "Sukuk", "Commodities"])
            with col_m4:
                manual_sector = st.selectbox(_t("Delsektor:", "Sub-sector:"), TARGET_SUBSECTORS + ["Kontanter", "Private investeringer"])
                
            if st.button(_t("➕ Tilføj manuelt aktiv", "➕ Add manual asset"), use_container_width=True):
                if manual_name:
                    virtual_ticker = f"PVT_{manual_name.upper().replace(' ', '_')}"
                    st.session_state.investor_holdings.append({
                        "Company Name": manual_name,
                        "Ticker": virtual_ticker,
                        "Shares": 1,
                        "Category": UI_TO_DB_MAP.get(manual_category, "Kontanter/Private"),
                        "Sector": manual_sector,
                        "manual_value": manual_value,
                        "Kurs": manual_value
                    })
                    st.success(_t(f"Tilføjede {manual_name}.", f"Added {manual_name}."))
                    time.sleep(1)
                    st.rerun()
        else:
            search_query = st.text_input(_t("🔍 Søg efter selskab eller ticker:", "🔍 Search by company name or ticker:"))

            if search_query:
                search_results = search_tickers_by_name_multi(search_query)
                if search_results:
                    options_format = [f"{r['name']} ({r['symbol']})" for r in search_results]
                    selected_option_str = st.selectbox(_t("Vælg det rigtige aktiv fra listen:", "Select the correct asset:"), options_format)
                    
                    selected_idx = options_format.index(selected_option_str)
                    target_asset = search_results[selected_idx]
                    resolved_ticker = target_asset["symbol"]
                    comp_name = target_asset["name"]
                    
                    try:
                        cat, sub_sec = get_category_and_sector_failsafe(resolved_ticker, target_category=st.session_state.targets)
                        display_cat = DB_TO_UI_MAP.get(cat, cat)
                        
                        render_html(f"""
<div style="background-color: #F8FAFC; border: 1px solid #C5A880; padding: 15px; border-radius: 6px; margin-top: 15px; margin-bottom: 15px;">
    <strong>🔍 Bekræftet match:</strong> {comp_name} ({resolved_ticker})<br>
    <strong>Aktivklasse:</strong> {display_cat} | <strong>Sektor:</strong> {sub_sec}
</div>
""")
                        
                        col_shares, col_add = st.columns([1, 1])
                        with col_shares:
                            shares_to_add = st.number_input(_t("Antal aktier ejet:", "Shares owned:"), min_value=1, value=10)
                        with col_add:
                            st.write(" ")
                            st.write(" ")
                            if st.button(_t("➕ Tilføj til min portefølje", "➕ Add to portfolio"), use_container_width=True):
                                exists = False
                                for h in st.session_state.investor_holdings:
                                    if h["Ticker"] == resolved_ticker:
                                        h["Shares"] += shares_to_add
                                        exists = True
                                        break
                                if not exists:
                                    st.session_state.investor_holdings.append({
                                        "Company Name": comp_name,
                                        "Ticker": resolved_ticker,
                                        "Shares": shares_to_add,
                                        "Category": cat,
                                        "Sector": sub_sec,
                                        "Kurs": 0.0
                                    })
                                st.success(_t(f"Tilføjede {shares_to_add} stk. {comp_name}.", f"Added {shares_to_add} shares of {comp_name}."))
                                time.sleep(1)
                                st.rerun()
                    except Exception as e:
                        st.error(f"Kunne ikke hente data for {resolved_ticker}: {str(e)}")

    st.write("---")
    st.write(_t("### Dine aktive positioner:", "### Your active holdings:"))
    if st.session_state.investor_holdings and not st.session_state.is_new_investor:
        holdings_df = pd.DataFrame(st.session_state.investor_holdings)
        holdings_df['Category_Display'] = holdings_df['Category'].apply(lambda x: DB_TO_UI_MAP.get(x, x))
        
        edited_holdings = st.data_editor(
            holdings_df,
            num_rows="dynamic",
            column_config={
                "Company Name": st.column_config.TextColumn(_t("Navn", "Name"), disabled=True),
                "Ticker": st.column_config.TextColumn("Ticker", disabled=True),
                "Shares": st.column_config.NumberColumn(_t("Antal", "Shares"), min_value=1),
                "Category_Display": st.column_config.TextColumn(_t("Aktivklasse", "Asset class"), disabled=True),
                "Sector": st.column_config.TextColumn(_t("Delsektor", "Sub-sector"), disabled=True),
                "manual_value": st.column_config.NumberColumn(_t("Manuel værdi (DKK)", "Manual value (DKK)"), min_value=0)
            },
            use_container_width=True,
            key="holdings_editor"
        )
        
        if not edited_holdings.equals(holdings_df):
            st.session_state.investor_holdings = edited_holdings.to_dict(orient="records")
            st.rerun()
            
        if st.button(_t("💾 Gem ændringer i min profil", "💾 Save changes to my profile"), use_container_width=True):
            status = save_user_portfolio_to_db(
                email=st.session_state.user_email,
                password=login_password,
                holdings=st.session_state.investor_holdings,
                targets=st.session_state.targets,
                horizon=st.session_state.horizon,
                name=st.session_state.user_name,
                frequency=st.session_state.frequency
            )
            if status == "success":
                st.success(_t("Ændringerne blev gemt på din profil!", "Changes saved successfully!"))
    elif st.session_state.is_new_investor:
        st.info(_t("Nybegynder-status aktiveret. Du behøver ikke indtaste beholdninger.", "New investor status activated. No holdings required."))
    else:
        st.info(_t("Porteføljen er tom lige nu. Tilføj aktiver herover for at komme videre.", "Your portfolio is empty. Add assets above to proceed."))

    # Navigationsknapper
    st.write(" ")
    col_prev, col_next = st.columns(2)
    with col_prev:
        if st.button("⬅ " + _t("Tilbage", "Back"), use_container_width=True):
            st.session_state.step = 3
            st.rerun()
    with col_next:
        if st.button(_t("Næste trin", "Next step") + " ➔", use_container_width=True):
            st.session_state.step = 5
            st.rerun()


# --- TRIN 5: WATCHLIST & AKTIVERING (MED DIREKTE NAVIGATION VED FEJL) ---
elif st.session_state.step == 5:
    st.subheader(_t("Udsendelse & Aktiver dit LLM Council", "Delivery & Activate your LLM Council"))
    
    watchlist_input = st.text_input(_t("Monitorer selskaber i Watchlist (kommasepareret):", "Monitor tickers in your Watchlist (comma-separated):"), "TRMB, SAP, SPSK, AEM, NEM")
    watchlist_list = [t.strip().upper() for t in watchlist_input.split(",") if t.strip()]

    st.write("---")
    
    # 1. VALIDERING OG SMART NAVIGERINGS-GENVEJ
    validation_passed = True
    if not st.session_state.is_new_investor and not st.session_state.investor_holdings:
        validation_passed = False
        st.error(_t(
            "⚠️ Fejl: Du skal tilføje mindst én aktiv position under Trin 4 for at kunne generere rapporten.",
            "⚠️ Error: You must add at least one active position under Step 4 to run your briefing."
        ))
        # UX Rettelse: En direkte knap, der tager dig direkte tilbage til Trin 4 uden nulstilling
        if st.button(_t("Gå direkte til Trin 4 for at tilføje aktiver ➔", "Go directly to Step 4 to add assets ➔"), use_container_width=True):
            st.session_state.step = 4
            st.rerun()
            
    elif not st.session_state.user_email or "@" not in st.session_state.user_email:
        validation_passed = False
        st.error(_t("⚠️ Fejl: Du mangler at angive en gyldig e-mailadresse i din profil under Trin 2.", "⚠️ Error: Please enter a valid email address in your profile under Step 2."))

    # 2. AKTIVERINGSKNAPPER
    col_b1, col_b2 = st.columns([2, 1])
    with col_b1:
        if st.button("🚀 " + _t("Kør LLM Council & Send min første rapport", "Run LLM Council & Send my first report"), use_container_width=True, disabled=(not validation_passed)):
            with st.spinner(_t("Screening mod Sharia- og gældskrav samt dannelse af podcast... Det tager ca. 60 sekunder.", "Screening against Shariah debt-limits and compiling your podcast... This takes about 60 seconds.")):
                success, msg = asyncio.run(process_instant_briefing(
                    st.session_state.user_email,
                    st.session_state.investor_holdings,
                    watchlist_list,
                    st.session_state.targets,
                    st.session_state.user_name,
                    st.session_state.horizon,
                    st.session_state.is_new_investor,
                    []
                ))
                if success:
                    st.success(f"Udført! {msg}")
                    st.balloons()
                else:
                    st.error(f"Fejl: {msg}")
    with col_b2:
        if st.session_state.investor_holdings and validation_passed:
            excel_bytes = generate_excel_template_bytes(
                st.session_state.investor_holdings, 
                watchlist_list,
                st.session_state.targets,
                {}
            )
            st.download_button(
                label="📥 " + _t("Hent mit Excel-ark", "Download Excel Sheet"),
                data=excel_bytes,
                file_name=f"{st.session_state.user_name}_Live_Portfolio.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # Navigationsknapper
    st.write(" ")
    col_prev, col_reset = st.columns(2)
    with col_prev:
        if st.button("⬅ " + _t("Tilbage", "Back"), use_container_width=True):
            st.session_state.step = 4
            st.rerun()
    with col_reset:
        if st.button("🔄 " + _t("Start forfra (Trin 1)", "Start over (Step 1)"), use_container_width=True):
            st.session_state.step = 1
            st.rerun()


# =====================================================================
#  NATIVE, RESPONSIV DISCLAIMER (LØSER JURIDISK ANVARS-PROBLEMET)
# =====================================================================
st.write(" ")
st.warning(_t(
    "Legal Disclaimer:\n\n"
    "LLM Council er et automatiseret, AI-baseret informations- og inspirationsværktøj til personligt brug. "
    "Vi tilbyder IKKE autoriseret eller licenseret finansiel rådgivning, og vi foretager ikke formelle investeringsbeslutninger på dine vegne.\n\n"
    "Finansielle markeder indebærer altid en risiko for tab, og Shariah-fortolkninger kan variere på tværs av forskellige retslærde og madhabs. "
    "Du bør altid basere dine endelige investeringsvalg på dine egne vurderinger, personlige overbevisninger og sund fornuft.\n\n"
    "For en uafhængig og manuel revision af gældsforhold, regnskabstal og compliance anbefaler vi at anvende det anerkendte værktøj Zoya Finance Platform.",
    
    "Legal Disclaimer:\n\n"
    "The LLM Council is an automated, AI-driven informational and educational inspiration tool. It is NOT a licensed financial advisor, nor does it provide personalized investment advice or regulatory financial mandates.\n\n"
    "Financial markets carry inherent risks, and Shariah-compliance interpretations can vary across different scholars and madhabs. You must always base your final investment decisions on your own research, personal convictions, and common sense.\n\n"
    "To manually audit and double-check Shariah-compliance, financial health, or business profiles, we highly recommend utilizing the official Zoya Finance Platform."
))
