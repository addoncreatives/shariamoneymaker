import os
import sys
import time
import io
import re
import json
import traceback
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ---------------------------------------------------------------------
#  SIKKERHEDSNET: Automatisk installation af openpyxl, hvis det mangler
# ---------------------------------------------------------------------
try:
    import openpyxl
except ImportError:
    import subprocess
    print("Sikkerhedsnet: openpyxl mangler i dit miljø. Installerer automatisk via pip...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

import yfinance as yf
import requests
import pandas as pd

# =====================================================================
#  KONFIGURATION OG MILJØVARIABLER
# =====================================================================

TARGET_PORTFOLIO = {
    "Tech & B2B Software": 20.0,
    "Defensivt Forbrug & Healthcare": 20.0,
    "Infrastruktur & Grøn Omstilling": 20.0,
    "Råvarer": 10.0,
    "ETFer & Sukuk": 30.0
}

# GLOBAL ISLAMIC GROWTH UNIVERSE (Proaktiv søgebase)
GLOBAL_COMPLIANT_GROWTH_POOL = {
    "Tech & B2B Software": [
        "TRMB", "SAP", "IFX.DE", "MSFT", "ASML", "NVDA", "ADBE", "CRM", "SNPS", 
        "ANSS", "CSCO", "AMAT", "LRCX", "NOW", "PANW", "FTNT", "ORCL"
    ],
    "Defensivt Forbrug & Healthcare": [
        "ORK.OL", "NOVO-B.CO", "6869.T", "AZN.ST", "REGN", "ISRG", "LLY", "VRTX", 
        "SYK", "MRK", "ZTS", "MDT", "GILD", "EL.PA", "NSRGY"
    ],
    "Infrastruktur & Grøn Omstilling": [
        "VWS.CO", "NKT.CO", "FLS.CO", "ROCK-B.CO", "ENPH", "FSLR", "ETN", "ABB", 
        "ALB", "ORSTED.CO", "SIE.DE", "GE", "NEE", "SRE"
    ],
    "Råvarer": [
        "WPM", "NEM", "GOLD", "AEM", "FNV", "RGLD", "BHP", "RIO", "FCX", "VALE"
    ],
    "ETFer & Sukuk": [
        "IGDA.L", "SPSK", "HLAL", "UMMA", "ISWD.L", "ISUS.L", "HIWS.L"
    ]
}

# Hent Sheet ID og API-nøgler
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
if not GOOGLE_SHEET_ID or GOOGLE_SHEET_ID.strip() == "":
    GOOGLE_SHEET_ID = "1EnE2XkQySaGsdaxR5KySZZ924LT66ICo"

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# E-mail indstillinger (med standard fallbacks)
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
if not EMAIL_SENDER or EMAIL_SENDER.strip() == "":
    EMAIL_SENDER = "wazir.ilyas@gmail.com"

EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")
if not EMAIL_RECEIVER or EMAIL_RECEIVER.strip() == "":
    EMAIL_RECEIVER = "addoncreatives@gmail.com"

EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")


# =====================================================================
#  GOOGLE SHEETS / EXCEL AGENT
# =====================================================================
class GoogleSheetsAgent:
    def __init__(self, sheet_id: str):
        self.sheet_id = sheet_id

    def _read_tab_as_df(self, tab_name: str) -> pd.DataFrame:
        url = f"https://docs.google.com/spreadsheets/d/{self.sheet_id}/export?format=xlsx"
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            df = pd.read_excel(io.BytesIO(response.content), sheet_name=tab_name, engine='openpyxl')
            return df
        except Exception as e:
            raise RuntimeError(f"Kunne ikke indlæse fanen '{tab_name}': {str(e)}")

    def _clean_and_align_df(self, df: pd.DataFrame, key_header_word: str) -> pd.DataFrame:
        for col in df.columns:
            if key_header_word.lower() in str(col).lower():
                return df
                
        for idx, row in df.head(10).iterrows():
            row_values = [str(val).lower() for val in row.values]
            if any(key_header_word.lower() in val for val in row_values):
                new_columns = df.iloc[idx].values
                df.columns = new_columns
                df = df.iloc[idx+1:].reset_index(drop=True)
                return df
        return df

    def _find_column_by_keyword(self, df: pd.DataFrame, keyword: str) -> str:
        for col in df.columns:
            if keyword.lower() in str(col).lower():
                return col
        return None

    def get_current_weights(self) -> dict:
        try:
            raw_df = self._read_tab_as_df("Beholdninger")
            df = self._clean_and_align_df(raw_df, "drivkraft")
            
            drivkraft_col = self._find_column_by_keyword(df, "drivkraft")
            weight_col = self._find_column_by_keyword(df, "vægt")

            if not drivkraft_col or not weight_col:
                raise KeyError("Kunne ikke lokalisere Drivkraft- eller Vægt-kolonne.")

            def clean_weight(val):
                if pd.isna(val) or val == "":
                    return 0.0
                val_str = str(val).replace('%', '').replace(',', '.').strip()
                try:
                    return float(val_str)
                except ValueError:
                    return 0.0

            df['Cleaned_Weight'] = df[weight_col].apply(clean_weight)
            grouped = df.groupby(drivkraft_col)['Cleaned_Weight'].sum().to_dict()

            normalized_portfolio = {}
            for target_key in TARGET_PORTFOLIO.keys():
                sum_val = 0.0
                for g_key, g_val in grouped.items():
                    if target_key.lower() in str(g_key).lower() or str(g_key).lower() in target_key.lower():
                        sum_val += g_val
                normalized_portfolio[target_key] = sum_val

            return normalized_portfolio
        except Exception as e:
            print(f"Advarsel under indlæsning af 'Beholdninger': {str(e)}")
            return {k: 0.0 for k in TARGET_PORTFOLIO.keys()}

    def get_watchlist_tickers(self) -> list:
        try:
            raw_df = self._read_tab_as_df("Opsummering")
            df = self._clean_and_align_df(raw_df, "huller")
            
            watchlist_col = self._find_column_by_keyword(df, "huller") or self._find_column_by_keyword(df, "watchlist")
            
            tickers = []
            if watchlist_col:
                raw_series = df[watchlist_col]
            else:
                if len(df.columns) >= 14:
                    raw_series = df.iloc[:, 13]
                else:
                    return []

            for val in raw_series:
                if pd.isna(val):
                    continue
                val_str = str(val).strip().upper()
                if val_str and len(val_str) < 12 and re.match(r'^[A-Z0-9\.\-]+$', val_str):
                    if val_str not in ["TICKER", "STATUS", "POSITION", "HULLER"]:
