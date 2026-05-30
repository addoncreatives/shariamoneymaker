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
#  SIKKERHEDSNET: Automatisk installation af openpyxl, hvis det mangler
# ---------------------------------------------------------------------
try:
    import openpyxl
except ImportError:
    import subprocess
    print("Sikkerhedsnet: openpyxl mangler. Installerer automatisk...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

# ---------------------------------------------------------------------
#  SIKKERHEDSNET: Automatisk installation af edge-tts, hvis det mangler
# ---------------------------------------------------------------------
try:
    import edge_tts
except ImportError:
    import subprocess
    print("Sikkerhedsnet: edge-tts mangler. Installerer automatisk...")
    subprocess.call([sys.executable, "-m", "pip", "install", "edge-tts"])
    import edge_tts

import yfinance as yf
import requests
import pandas as pd
import streamlit as st

# =====================================================================
#  KONFIGURATION OG STANDARD-MÅLVÆGTE (DYNAMISK JUSTERBARE)
# =====================================================================

DISPLAY_CATEGORIES = {
    "Aktier": "Equities",
    "Sukuk": "Sukuk (Islamic Bonds)",
    "Råvarer": "Commodities",
    "Kontanter/Private": "Cash / Private Sector"
}

TARGET_SUBSECTORS = [
    "Pharma", "Medico", "Vind", "Elektrificering", "Infrastruktur", 
    "Byggeri", "Materialer", "Industri", "Halvleder", "Mining royalty", 
    "Energi", "ETF - global", "ETF - regional", "Sukuk", "Kontanter", 
    "Private investeringer", "Landbrug / gødning", "Consumer staples / dagligvarer", 
    "Software / platform", "Industrielle metaller / kobber", "Logistik"
]

# GLOBAL ISLAMIC GROWTH UNIVERSE (Tilpasset dine delsektorer)
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

# Hent dine master-nøgler fra Streamlit Secrets
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1EnE2XkQySaGsdaxR5KySZZ924LT66ICo")
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


# =====================================================================
#  STÆRK TEKST-NORMALISERING (FJERNER STAVE- OG APOSTROF-FEJL)
# =====================================================================
def normalize_string(s: str) -> str:
    if not s or pd.isna(s):
        return ""
    s = str(s).lower().strip()
    s = s.replace("og", "and").replace("&", "and")
    s = s.replace("'", "")
    s = re.sub(r'[^a-z0-9æøå]', '', s)
    s = s.replace("etfer", "etf").replace("etfs", "etf")
    return s


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
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
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

    def map_to_category_and_sector(self, symbol: str, sector: str, industry: str) -> tuple:
        sym = symbol.upper()
        sec_l = sector.lower() if sector else ""
        ind_l = industry.lower() if industry else ""

        # Sukuk
        if "sukuk" in sym or sym in ["SPSK", "SKUK"]:
            if self.target_category == "Kontanter/Private":
                return "Kontanter/Private", "Sukuk"
            return "Sukuk", "Sukuk"
            
        # Råvarer
        if sym in ["WPM", "FNV", "RGLD"]:
            return "Råvarer", "Mining royalty"
        if sym in ["NEM", "GOLD", "AEM", "BHP", "RIO", "FCX", "VALE"] or \
           any(w in ind_l for w in ["gold", "silver", "precious metals", "copper", "aluminum"]):
            return "Råvarer", "Industrielle metaller / kobber"

        # Kontanter
        if "cash" in sym or "money market" in sec_l:
            return "Kontanter/Private", "Cash Equivalents"

        # Dynamisk fallback: Selskabets reelle industri fra yfinance bruges direkte som delsektor
        dynamic_subsector = industry if (industry and industry != "Other") else sector
        return "Aktier", dynamic_subsector

    def screen_ticker(self, symbol: str) -> dict:
        try:
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

    def run_proactive_analysis(self, candidates_data: list, category: str, deficit: float, current_portfolio_str: str, current_holdings_str: str, sector_distribution_str: str) -> str:
        candidates_json = json.dumps(candidates_data, indent=2, ensure_ascii=False)
        english_category = DISPLAY_CATEGORIES.get(category, category)
        
        prompt = f"""
        You are an elite financial advisory council ("LLM Council") presenting a strategic investment briefing to your highly valued VIP client, Wazir [3].
        
        THE INVESTOR PROFILE & MODEL:
        - Overarching Strategic Model: Customize based on Wazir's targets [3].
        - Under Evaluation Tonight: {english_category} (Current Deficit: {deficit:.2f}%) [3].
        - Current Portfolio Allocations (Target vs Actual): {current_portfolio_str} [3].
        
        THE INVESTOR'S STRATEGIC SUB-SECTORS (DYNAMICALLY DETECTED FROM THE HOLDINGS):
        {sector_distribution_str}
        
        THE INVESTOR'S CURRENT HOLDINGS:
        {current_holdings_str}
        
        NEW COMPLIANT SCREENED CANDIDATES TO BE EVALUATED:
        {candidates_json}
        
        YOUR OBJECTIVE (DELIVER ENTIRE BRIEFING IN BEAUTIFUL ENGLISH HTML):
        Generate a complete, institutional-grade, highly engaging HTML investment newsletter in English [3].
        
        Brug udelukkende inline CSS-styling for maximum compatibility with Gmail [3].
        Design guidelines:
        - Main container: `<div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 25px; background-color: #ffffff; color: #334155; line-height: 1.6;">`
        - Colors: Dark Slate (`#0F172A`) for headings. Accent/highlights: Warm Gold/Sand (`#C5A880`) [3].
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
        <p><strong>Prepared exclusively for:</strong> Wazir</p>
        <p><strong>Focus tonight:</strong> {english_category} (Deficit: {deficit:.2f}%)</p>
        
        <hr style="border: 0; border-top: 1px solid #E2E8F0; margin: 20px 0;">
        
        <h2>SECTION 1 — PORTFOLIO DIAGNOSTIC & INDIRECT EXPOSURES</h2>
        Analyze Wazir's current holdings and how they map to their strategic sub-sectors [3]. Do existing assets already provide satisfactory indirect exposure to the focus theme [3]? Discuss Saxo Bank limitations and Sharia compliance filters as boundaries, and explain how they can diversify across different underlying economic drivers if direct options are limited.
        
        <h2>SECTION 2 — DEEP-DIVE CONSULTANT ANALYSIS (UP TO 10 SCREENED CANDIDATES)</h2>
        For each candidate, write an elegant card covering:
        1. <strong>Investment Case</strong> (How it balances Wazir's current assets).
        2. <strong>Financial Highlights</strong> (Omsætningsvækst, marginer, cash flow based on live data).
        3. <strong>Future Outlook & Pipeline</strong>.
        4. <strong>Risk Assessment</strong>.
        5. <strong>Momentum & Trend Analysis</strong> (3-month momentum vs. 3-year growth trajectory).
        6. <strong>Analyst Insight & Sources</strong>: Insert exactly 2 clickable links structured beautifully in HTML (e.g. `<a href="https://seekingalpha.com/symbol/TICKER" style="color: #C5A880; text-decoration: none; font-weight: bold;">Seeking Alpha</a>`).
        
        <h2>SECTION 3 — THE ASYNCHRONOUS COUNCIL DEBAT (TOP-3)</h2>
        Select the top 3 assets. Moderate a high-stakes, dramatic debate among the 5 financial advisers using the styled left-bordered divs [3]. Show conflict, arguments on valuation, capex, and macro timing.
        
        <h2>SECTION 4 — THE CHAIRMAN'S DEKRET (RECOMMENDATION)</h2>
        The Chairman's final clear directive enclosed in the gold callout box [3]. Conclude with a highly precise, step-by-step action plan for Wazir's Saxo Investor account over the next 7 days [3].
        
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

    def generate_podcast_audio(self, report_html: str) -> str:
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
                "Create a high-energy Bloomberg-style financial show moderated by Sarah and Mark. "
                "The show MUST open with Sarah and Mark introducing themselves and welcoming our VIP client, Wazir [3]. "
                "Then, they introduce and interview our 5 resident advisers: "
                "Contrarian (the risk-obsessed skeptic who must interrupt with: 'But what if the market turns tomorrow?'), "
                "First-Principles (the logical mathematician using raw numbers), "
                "Expansionist (the highly bullish growth hunter wanting to deploy capital), "
                "Outsider (the big-picture strategist analyzing indirect exposures like NKT/FLS and favoring royalty models), "
                "and Executor (the pragmatic guy checking Saxo tradeability and Dollar-Cost Averaging) [3]. "
                "The show must conclude with Sarah and Mark summarizing the Chairman's final recommendation and "
                "giving Wazir a highly clear, actionable next step for his Saxo account."
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
    def send_email(subject: str, html_content: str, attachment_path: str = None):
        if not EMAIL_PASSWORD or EMAIL_PASSWORD.strip() == "":
            print("EMAIL_PASSWORD mangler i GitHub Secrets. Udskriver HTML i konsol:")
            print(html_content)
            return

        msg = MIMEMultipart()
        msg["From"] = EMAIL_SENDER
        msg["To"] = EMAIL_RECEIVER
        msg["Subject"] = subject
        msg.attach(MIMEText(html_content, "html", "utf-8"))

        if attachment_path and os.path.exists(attachment_path):
            print(f"Vedhæfter lydfil: {attachment_path}...")
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {os.path.basename(attachment_path)}",
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

# Custom CSS til styling (Slate & Gold tema)
st.markdown("""
    <style>
    .main-title {
        color: #0F172A;
        font-family: 'Georgia', serif;
        font-size: 38px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 5px;
    }
    .subtitle {
        color: #C5A880;
        font-family: 'Helvetica Neue', sans-serif;
        font-size: 16px;
        text-align: center;
        text-transform: uppercase;
        letter-spacing: 2px;
        margin-bottom: 30px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🗳️ LLM Council</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Premium Investment Newsletter & Podcast Service</div>', unsafe_allow_html=True)

st.write("Welcome, Investor. This portal configures your personal, automated investment advisory department. "
         "Once submitted, you will receive your **first investment dossier and audio podcast in your inbox within 60 seconds**.")

# =====================================================================
#  STEP 1: BRUGEROPLYSNINGER & DYNAMISKE MÅLVÆGTE
# =====================================================================
st.subheader("Step 1: Your Delivery Information")
user_email = st.text_input("Enter your Email Address to receive the briefings:", placeholder="your.name@gmail.com")

st.subheader("Step 1.5: Customize Your Target Allocations")
st.write("Define your target weighting for the major asset classes. The system will automatically normalize them to sum to 100%.")

col1, col2, col3, col4 = st.columns(4)
with col1:
    target_stocks = st.number_input("Equities (%)", min_value=0.0, max_value=100.0, value=25.0)
with col2:
    target_sukuk = st.number_input("Sukuk (%)", min_value=0.0, max_value=100.0, value=25.0)
with col3:
    target_commodities = st.number_input("Commodities (%)", min_value=0.0, max_value=100.0, value=25.0)
with col4:
    target_cash = st.number_input("Cash/Private (%)", min_value=0.0, max_value=100.0, value=25.0)

# Normaliser mål-vægte så de summer til 100%
total_target = target_stocks + target_sukuk + target_commodities + target_cash
if total_target > 0:
    custom_targets = {
        "Aktier": (target_stocks / total_target) * 100.0,
        "Sukuk": (target_sukuk / total_target) * 100.0,
        "Råvarer": (target_commodities / total_target) * 100.0,
        "Kontanter/Private": (target_cash / total_target) * 100.0
    }
else:
    custom_targets = {"Aktier": 25.0, "Sukuk": 25.0, "Råvarer": 25.0, "Kontanter/Private": 25.0}

# =====================================================================
#  STEP 2: ENKELT, FULDAUTOMATISERET HOLDINGS ENTRANCE
# =====================================================================
st.subheader("Step 2: Your Active Portfolio")
st.write("Enter your active positions. Enter ONLY Ticker and Shares; our backend automatically detects Category and Sector live [3]!")

default_holdings = [
    {"Ticker": "NOVO-B.CO", "Shares": 10, "Category Override (Optional)": None, "Sector Override (Optional)": None},
    {"Ticker": "MSAU.L", "Shares": 10, "Category Override (Optional)": None, "Sector Override (Optional)": None},
    {"Ticker": "IGDA.L", "Shares": 23, "Category Override (Optional)": None, "Sector Override (Optional)": None},
    {"Ticker": "SKUK", "Shares": 100, "Category Override (Optional)": None, "Sector Override (Optional)": None},
    {"Ticker": "WPM", "Shares": 5, "Category Override (Optional)": None, "Sector Override (Optional)": None}
]

df_default = pd.DataFrame(default_holdings)

# Den interaktive tabel - kun to felter er reelt nødvendige for brugeren!
edited_df = st.data_editor(
    df_default,
    num_rows="dynamic",
    column_config={
        "Ticker": st.column_config.TextColumn("Ticker", help="e.g. MSFT, NOVO-B.CO, SKUK"),
        "Shares": st.column_config.NumberColumn("Shares", min_value=1),
        "Category Override (Optional)": st.column_config.SelectboxColumn("Category Override (Optional)", options=["Aktier", "Sukuk", "Råvarer", "Kontanter/Private"]),
        "Sector Override (Optional)": st.column_config.SelectboxColumn("Sector Override (Optional)", options=TARGET_SUBSECTORS)
    },
    use_container_width=True
)

# =====================================================================
#  STEP 3: WATCHLIST (VALGFRI)
# =====================================================================
st.subheader("Step 3: Your Watchlist (Optional)")
st.write("Enter tickers to monitor. If left empty, the system will dynamically screen our global Sharia-growth pool [3].")
watchlist_input = st.text_input(
    "Enter Tickers (comma-separated):",
    "TRMB, SAP, SPSK, AEM, NEM"
)

watchlist_list = [t.strip().upper() for t in watchlist_input.split(",") if t.strip()]

# =====================================================================
#  FUNKTION TIL AT SKABE LIVE-RAPPORT OG PODCAST AUTOMATISK PÅ STREAMLIT
# =====================================================================
async def process_instant_briefing(receiver_email, holdings_df, watchlist, target_allocations):
    """
    Kører hele investerings-motoren asynkront direkte på Streamlits cloud-server.
    Detekterer automatisk sektorer og kategorier live ved hjælp af yfinance [3].
    """
    total_assets_count = len(holdings_df)
    
    portfolio_distribution = {"Aktier": 0.0, "Sukuk": 0.0, "Råvarer": 0.0, "Kontanter/Private": 0.0}
    sector_distribution = {}
    
    # Midlertidig screener til at kortlægge live sektorer
    screener = ScreenerComplianceAgent([], target_category="Aktier")
    
    # Looper over indtastede positioner og detekterer kategorier/sektorer helt automatisk live!
    for row in holdings_df.itertuples():
        symbol = str(row.Ticker).strip().upper()
        
        # Hent valgfri overstyringer
        cat_override = getattr(row, "Category_Override_Optional", None)
        sec_override = getattr(row, "Sector_Override_Optional", None)
        
        category = cat_override if (cat_override and not pd.isna(cat_override)) else None
        subsector = sec_override if (sec_override and not pd.isna(sec_override)) else None
        
        # Hvis ikke manuelt overstyret, slår vi det op live på yfinance!
        if not category or not subsector:
            try:
                t = yf.Ticker(symbol)
                info = t.info
                sec = info.get("sector", "Other")
                ind = info.get("industry", "Other")
                
                cat_detect, sub_detect = screener.map_to_category_and_sector(symbol, sec, ind)
                category = category if category else cat_detect
                subsector = subsector if subsector else sub_detect
            except Exception:
                category = category if category else "Aktier"
                subsector = subsector if subsector else "Other"
                
        # Læg vægtene sammen
        weight_chunk = (100.0 / total_assets_count)
        if category in portfolio_distribution:
            portfolio_distribution[category] += weight_chunk
        if subsector not in sector_distribution:
            sector_distribution[subsector] = 0.0
        sector_distribution[subsector] += weight_chunk

    # 2. Find fokus-kategori baseret på de dynamiske målvægte!
    pm = PortfolioManagerAgent(portfolio_distribution, target_allocations)
    focus_category, deficit = pm.identify_underweighted_focus()
    
    # 3. Proaktiv søgning
    growth_pool = GLOBAL_COMPLIANT_GROWTH_POOL.get(focus_category, [])
    combined_candidates = list(set(watchlist + growth_pool))
    
    # 4. Kør screening
    screener.target_category = focus_category
    screener.tickers = combined_candidates
    approved_stocks = screener.run_screening(focus_category)
    target_candidates = approved_stocks[:10]
    
    if not target_candidates:
        return False, "No compliant assets could be found in your focused category."

    # 5. Indhent detaljerede yfinance metrics
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

    # 6. Kør Gemini 3.5 Flash til nyhedsbrevet
    current_weights_str = json.dumps(portfolio_distribution, indent=2, ensure_ascii=False)
    current_holdings = [{"ticker": r.Ticker, "asset_class": r.Aktivklasse if hasattr(r, 'Aktivklasse') else "Aktier"} for r in holdings_df.itertuples()]
    current_holdings_str = json.dumps(current_holdings, indent=2, ensure_ascii=False)
    sector_distribution_str = json.dumps(sector_distribution, indent=2, ensure_ascii=False)
    
    council_agent = CouncilAgent(GEMINI_API_KEY)
    report_html = council_agent.run_proactive_analysis(
        candidates_data=detailed_candidates_data,
        category=focus_category,
        deficit=deficit,
        current_portfolio_str=current_weights_str,
        current_holdings_str=current_holdings_str,
        sector_distribution_str=sector_distribution_str
    )

    # 7. Kør Podcastfy til MP3-kompileringen
    output_mp3 = "llm_council_podcast.mp3"
    podcast_compiled = False
    
    podcast_agent = PodcastAgent(GEMINI_API_KEY)
    generated_file = podcast_agent.generate_podcast_audio(report_html)
    
    if generated_file and os.path.exists(generated_file):
        import shutil
        shutil.copyfile(generated_file, output_mp3)
        podcast_compiled = True

    # 8. Send e-mailen (Vi overskriver EMAIL_RECEIVER med brugerens indtastede mail)
    os.environ["EMAIL_RECEIVER"] = receiver_email
    subject = f"[LLM Council] Your Personal Strategic Briefing - Focus on {DISPLAY_CATEGORIES.get(focus_category, focus_category)}"
    
    # Sætter modtager-adressen dynamisk inden afsendelse
    global EMAIL_RECEIVER
    EMAIL_RECEIVER = receiver_email
    
    DeliveryAgent.send_email(
        subject=subject,
        html_content=report_html,
        attachment_path=output_mp3 if podcast_compiled else None
    )
    return True, "Your briefing and audio podcast have been sent to your email!"

# =====================================================================
#  SAAS START-KNAP
# =====================================================================
st.subheader("Step 4: Activate Your Personal Council")

if st.button("🚀 Start My LLM Council & Send First Report"):
    if not user_email or "@" not in user_email:
        st.error("Please enter a valid email address.")
    elif edited_df.empty:
        st.error("Please configure at least one active holding.")
    elif not GEMINI_API_KEY:
        st.error("SaaS master API key is missing on the server.")
    else:
        with st.spinner("Processing your holdings, screening candidates and generating your Bloomberg-style podcast... This takes about 60 seconds."):
            success, msg = asyncio.run(process_instant_briefing(user_email, edited_df, watchlist_list, custom_targets))
            if success:
                st.success(f"Boom! {msg}")
                st.balloons()
            else:
                st.error(f"Failed to generate briefing: {msg}")
