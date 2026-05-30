import streamlit as st
import pandas as pd
import openpyxl
import io
import os
import json
import asyncio
import requests
import datetime

# Importér dine agenter direkte fra dit hovedscript!
from llm_council import (
    GoogleSheetsAgent,
    ScreenerComplianceAgent,
    CouncilAgent,
    PodcastAgent,
    DeliveryAgent,
    PortfolioManagerAgent,
    TARGET_PORTFOLIO,
    GLOBAL_COMPLIANT_GROWTH_POOL,
    DISPLAY_CATEGORIES,
    TARGET_SUBSECTORS
)

# Sæt sidens opsætning med CNBC/Bloomberg tema
st.set_page_config(
    page_title="LLM Council - Premium Onboarding",
    page_icon="🗳️",
    layout="centered"
)

# Hent dine master-nøgler fra Streamlit Secrets
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1EnE2XkQySaGsdaxR5KySZZ924LT66ICo")

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
""", unsafe_allow_html=True) # Her er fejlen rettet fra limits til html!

st.markdown('<div class="main-title">🗳️ LLM Council</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Premium Investment Newsletter & Podcast Service</div>', unsafe_allow_html=True)

st.write("Welcome, Investor. This portal configures your personal, automated investment advisory department. "
         "Once submitted, you will receive your **first investment dossier and audio podcast in your inbox within 60 seconds**.")

# =====================================================================
#  STEP 1: BRUGEROPLYSNINGER & EMAIL
# =====================================================================
st.subheader("Step 1: Your Delivery Information")
user_email = st.text_input("Enter your Email Address to receive the briefings:", placeholder="your.name@gmail.com")

# =====================================================================
#  STEP 2: AKTIER & BEHOLDNINGER
# =====================================================================
st.subheader("Step 2: Your Active Portfolio")
st.write("Enter your active positions. This data calibrates the mathematical weightings of your council.")

default_holdings = [
    {"Ticker": "NOVO-B.CO", "Position_Name": "Novo Nordisk B", "Shares": 10, "Aktivklasse": "Aktier", "Sektor": "Pharma"},
    {"Ticker": "MSAU.L", "Position_Name": "Saudi Arabia ETF", "Shares": 10, "Aktivklasse": "Aktier", "Sektor": "ETF - regional"},
    {"Ticker": "IGDA.L", "Position_Name": "Invesco Islamic Global", "Shares": 23, "Aktivklasse": "Aktier", "Sektor": "ETF - global"},
    {"Ticker": "SKUK", "Position_Name": "iShares USD Sukuk", "Shares": 100, "Aktivklasse": "Sukuk", "Sektor": "Sukuk"},
    {"Ticker": "WPM", "Position_Name": "Wheaton Precious Metals", "Shares": 5, "Aktivklasse": "Råvarer", "Sektor": "Mining royalty"}
]

df_default = pd.DataFrame(default_holdings)

edited_df = st.data_editor(
    df_default,
    num_rows="dynamic",
    column_config={
        "Ticker": st.column_config.TextColumn("Ticker (Yahoo Finance format)", help="e.g. MSFT, NOVO-B.CO, SKUK"),
        "Position_Name": st.column_config.TextColumn("Asset Name"),
        "Shares": st.column_config.NumberColumn("Number of Shares", min_value=1),
        "Aktivklasse": st.column_config.SelectboxColumn("Aktivklasse (4x25% Category)", options=["Aktier", "Sukuk", "Råvarer", "Kontanter/Private"]),
        "Sektor": st.column_config.SelectboxColumn("Sub-sector (21 Sectors)", options=TARGET_SUBSECTORS)
    },
    use_container_width=True
)

# =====================================================================
#  STEP 3: WATCHLIST / HULLER
# =====================================================================
st.subheader("Step 3: Your Watchlist")
watchlist_input = st.text_input(
    "Enter the Tickers you want the council to monitor (comma-separated):",
    "TRMB, SAP, SPSK, AEM, NEM"
)

watchlist_list = [t.strip().upper() for t in watchlist_input.split(",") if t.strip()]

# =====================================================================
#  FUNKTION TIL AT SKABE LIVE-RAPPORT OG PODCAST AUTOMATISK PÅ STREAMLIT
# =====================================================================
async def process_instant_briefing(receiver_email, holdings_df, watchlist):
    """
    Kører hele investerings-motoren asynkront direkte på Streamlits cloud-server.
    Udløser den øjeblikkelige første mail og danner podcasten via Podcastfy [3].
    """
    total_assets_count = len(holdings_df)
    
    # Beregn estimerede procenter (en simpel fordeling til den første rapport)
    portfolio_distribution = {"Aktier": 0.0, "Sukuk": 0.0, "Råvarer": 0.0, "Kontanter/Private": 0.0}
    sector_distribution = {s: 0.0 for s in TARGET_SUBSECTORS}
    
    for row in holdings_df.itertuples():
        cat = row.Aktivklasse
        portfolio_distribution[cat] += (100.0 / total_assets_count)
        
        sec = row.Sektor
        if sec in sector_distribution:
            sector_distribution[sec] += (100.0 / total_assets_count)

    # 2. Find fokus-kategori
    pm = PortfolioManagerAgent(portfolio_distribution, TARGET_PORTFOLIO)
    focus_category, deficit = pm.identify_underweighted_focus()
    
    # 3. Proaktiv søgning
    growth_pool = GLOBAL_COMPLIANT_GROWTH_POOL.get(focus_category, [])
    combined_candidates = list(set(watchlist + growth_pool))
    
    # 4. Kør screening
    screener = ScreenerComplianceAgent(combined_candidates, target_category=focus_category)
    approved_stocks = screener.run_screening(focus_category)
    target_candidates = approved_stocks[:10]
    
    if not target_candidates:
        retur
