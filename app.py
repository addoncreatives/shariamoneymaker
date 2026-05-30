import streamlit as st
import pandas as pd
import openpyxl
import io
import re

# Sæt sidens opsætning med CNBC/Bloomberg tema
st.set_page_config(
    page_title="LLM Council - Onboarding Portal",
    page_icon="🗳️",
    layout="centered"
)

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
    .card {
        background-color: #F8FAFC;
        padding: 20px;
        border-radius: 8px;
        border: 1px solid #E2E8F0;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_limits=True)

st.markdown('<div class="main-title">🗳️ LLM Council</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">SaaS Onboarding & Portfolio Generator</div>', unsafe_allow_html=True)

st.markdown("""
Welcome, Investor. This onboarding portal automatically generates the live-updating Google Sheets template 
required to power your personal **LLM Council** automated investment assistant. 

Fill in your active assets and watchlist below to generate your customized workbook.
""")

# =====================================================================
#  STEP 1: CURRENT HOLDINGS (INTERAKTIV DATA EDITOR)
# =====================================================================
st.subheader("Step 1: Your Active Portfolio")
st.write("Add your current holdings. Ticker and Shares are the most critical pillars.")

# Standard eksempler på dine egne start-aktiver, så nye brugere kan se formatet
default_holdings = [
    {"Ticker": "NOVO-B.CO", "Position_Name": "Novo Nordisk B", "Shares": 10, "Aktivklasse": "Aktier", "Sektor": "Pharma"},
    {"Ticker": "MSAU.L", "Position_Name": "Saudi Arabia ETF", "Shares": 10, "Aktivklasse": "Aktier", "Sektor": "ETF - regional"},
    {"Ticker": "IGDA.L", "Position_Name": "Invesco Islamic Global", "Shares": 23, "Aktivklasse": "Aktier", "Sektor": "ETF - global"},
    {"Ticker": "SKUK", "Position_Name": "iShares USD Sukuk", "Shares": 100, "Aktivklasse": "Sukuk", "Sektor": "Sukuk"},
    {"Ticker": "WPM", "Position_Name": "Wheaton Precious Metals", "Shares": 5, "Aktivklasse": "Råvarer", "Sektor": "Mining royalty"}
]

df_default = pd.DataFrame(default_holdings)

# Den interaktive tabel, hvor brugeren kan tilføje/slette rækker direkte på skærmen
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
#  STEP 2: WATCHLIST / HULLER
# =====================================================================
st.subheader("Step 2: Your Watchlist")
watchlist_input = st.text_input(
    "Enter the Tickers you want to monitor (comma-separated):",
    "TRMB, SAP, SPSK, AEM, NEM"
)

# Rens indtastede Watchlist-tickers
watchlist_list = [t.strip().upper() for t in watchlist_input.split(",") if t.strip()]

# =====================================================================
#  EXCEL-GENERATOR FUNKTION MED LIVE GOOGLE-FORMELER
# =====================================================================
def generate_excel_template(holdings_df, watchlist):
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
    
    for idx, row in enumerate(holdings_df.itertuples(), start=2):
        ticker = str(row.Ticker).strip().upper()
        name = str(row.Position_Name).strip()
        shares = int(row.Shares)
        aktivklasse = str(row.Aktivklasse)
        sektor = str(row.Sektor)
        
        # Skriv værdier og indbyg de dynamiske GOOGLEFINANCE formler
        ws1.cell(row=idx, column=1, value=name)
        ws1.cell(row=idx, column=2, value=ticker)
        ws1.cell(row=idx, column=3, value="Ejer")
        ws1.cell(row=idx, column=4, value=shares)
        
        # LIVE FORMELER: Google Sheets henter selv prisen live og udregner markedsværdien
        ws1.cell(row=idx, column=5, value=f'=GOOGLEFINANCE(B{idx})') # Live pris
        ws1.cell(row=idx, column=6, value=f'=D{idx}*E{idx}')         # Antal x Live pris
        
        ws1.cell(row=idx, column=7, value=aktivklasse)
        ws1.cell(row=idx, column=8, value="")  # Drivkraft (valgfri)
        ws1.cell(row=idx, column=9, value=sektor)
        ws1.cell(row=idx, column=10, value="Global")
        
        # Vægt-formel: Værdien divideret med summen af hele kolonne F
        ws1.cell(row=idx, column=11, value=f'=F{idx}/SUM(F$2:F$100)')
        ws1.cell(row=idx, column=12, value="")
        ws1.cell(row=idx, column=13, value="")

    # 2. FANEN: Opsummering
    ws2 = wb.create_sheet(title="Opsummering")
    
    # Skriv headers så de passer til dit arks struktur (Watchlist ligger i Kolonne N)
    headers2 = ["4x25-overblik", "", "", "", "", "Økonomiske drivere", "", "", "", "Sektorere", "", "", "", "Huller / Watchlist"]
    ws2.append(headers2)
    
    # Indsæt Watchlist-tickers i Kolonne N (Kolonne 14)
    for idx, ticker in enumerate(watchlist, start=2):
        ws2.cell(row=idx, column=14, value=ticker)

    # Gem projektmappen i en hukommelses-buffer
    excel_data = io.BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)
    return excel_data

# =====================================================================
#  GENERER & DOWNLOAD KNAP
# =====================================================================
st.subheader("Step 3: Generate and Deploy")

if st.button("Generate My Live Sheet Template"):
    if edited_df.empty:
        st.error("Please add at least one holding before generating.")
    else:
        with st.spinner("Generating your live Excel workbook..."):
            excel_file = generate_excel_template(edited_df, watchlist_list)
            
            st.success("Success! Your live portfolio spreadsheet has been generated.")
            
            # Download knap til brugeren
            st.download_button(
                label="📥 Download Onboarding_Portfolio.xlsx",
                data=excel_file,
                file_name="Onboarding_Portfolio.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.markdown("""
            ### 🚀 Next Steps to start your LLM Council:
            1. **Upload to Google Drive:** Upload the downloaded `Onboarding_Portfolio.xlsx` to your Google Drive and open it as a Google Sheet.
            2. **Share the Sheet:** Click **Share** (Del), and set permissions to **"Anyone with the link can view"** (Alle med linket kan se).
            3. **Save your Sheet ID:** Copy your Sheet ID from the URL (the string between `/d/` and `/edit`). Save it as a GitHub Secret named `GOOGLE_SHEET_ID`.
            4. **Set up Secrets:** Add your `GEMINI_API_KEY`, `EMAIL_PASSWORD` to your GitHub secrets, and run the action!
            """)
