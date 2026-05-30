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

# Standard målvægte (hvis ingen sendes fra appen)
TARGET_PORTFOLIO = {
    "Aktier": 25.0,
    "Sukuk": 25.0,
    "Råvarer": 25.0,
    "Kontanter/Private": 25.0
}

# Oversættelse til den engelske nyhedsbrevs-rapport
DISPLAY_CATEGORIES = {
    "Aktier": "Equities",
    "Sukuk": "Sukuk (Islamic Bonds)",
    "Råvarer": "Commodities",
    "Kontanter/Private": "Cash / Private Sector"
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
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

EMAIL_SENDER = os.getenv("EMAIL_SENDER")
if not EMAIL_SENDER or EMAIL_SENDER.strip() == "":
    EMAIL_SENDER = "wazir.ilyas@gmail.com"

EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")
if not EMAIL_RECEIVER or EMAIL_RECEIVER.strip() == "":
    EMAIL_RECEIVER = "addoncreatives@gmail.com"

# Hent og rens adgangskoden (Sikkerhedsnet mod mellemrum)
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
#  GOOGLE SHEETS / EXCEL AGENT (MED FULL-AUTO DETEKTERING)
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

    def get_current_weights_and_sectors(self) -> tuple:
        """
        Læser fanen 'Beholdninger', identificerer vægtene og fordeler dem 
        både på hovedkasserne og på de specifikke delsektorer helt automatisk.
        """
        try:
            raw_df = self._read_tab_as_df("Beholdninger")
            df = self._clean_and_align_df(raw_df, "ticker")
            
            ticker_col = self._find_column_by_keyword(df, "ticker")
            drivkraft_col = self._find_column_by_keyword(df, "drivkraft")
            aktivklasse_col = self._find_column_by_keyword(df, "aktivklasse")
            sektor_col = self._find_column_by_keyword(df, "sektor")
            weight_col = self._find_column_by_keyword(df, "vægt")
            mv_col = self._find_column_by_keyword(df, "markedsværdi")

            if not weight_col:
                raise KeyError("Kunne ikke lokalisere de nødvendige kolonner i Google Sheet.")

            def clean_number(val):
                if pd.isna(val) or val == "":
                    return 0.0
                val_str = str(val).replace('%', '').replace('kr', '').replace('$', '').replace(' ', '').replace('\xa0', '').strip()
                if ',' in val_str and '.' in val_str:
                    if val_str.find('.') < val_str.find(','):
                        val_str = val_str.replace('.', '').replace(',', '.')
                    else:
                        val_str = val_str.replace(',', '').replace('.', '.')
                elif ',' in val_str:
                    val_str = val_str.replace(',', '.')
                try:
                    return float(val_str)
                except ValueError:
                    return 0.0

            df['Cleaned_Weight'] = df[weight_col].apply(clean_number)
            df['Cleaned_MV'] = df[mv_col].apply(clean_number) if mv_col else 0.0

            total_mv = df['Cleaned_MV'].sum()
            total_weight = df['Cleaned_Weight'].sum()

            if total_weight < 1.0 and total_mv > 0.0:
                df['Cleaned_Weight'] = (df['Cleaned_MV'] / total_mv) * 100.0
            elif total_mv > 0.0:
                df['Cleaned_Weight'] = (df['Cleaned_MV'] / total_mv) * 100.0

            portfolio_distribution = {k: 0.0 for k in TARGET_PORTFOLIO.keys()}
            sector_distribution = {}

            # Opret en midlertidig screener til at udføre auto-detektering på dine nuværende positioner
            temp_screener = ScreenerComplianceAgent([])

            for _, row in df.iterrows():
                row_weight = row['Cleaned_Weight']
                if row_weight <= 0.0:
                    continue

                symbol = str(row.get(ticker_col, "")).strip().upper()
                drivkraft_val = normalize_string(row.get(drivkraft_col, "")) if drivkraft_col else ""
                aktivklasse_val = normalize_string(row.get(aktivklasse_col, "")) if aktivklasse_col else ""
                sektor_val = normalize_string(row.get(sektor_col, "")) if sektor_col else ""

                # Hvis brugeren har overstyret i arket, bruger vi det
                combined_text = f"{drivkraft_val} {aktivklasse_val} {sektor_val}"
                
                category = None
                subsector = None

                if "sukuk" in combined_text:
                    category = "Sukuk"
                    subsector = "Sukuk"
                elif any(word in combined_text for word in ["råvarer", "ravarer", "guld", "gold", "commodities"]):
                    category = "Råvarer"
                    subsector = "Commodities"
                elif any(word in combined_text for word in ["kontant", "cash", "private"]):
                    category = "Kontanter/Private"
                    subsector = "Cash"
                elif "aktie" in combined_text:
                    category = "Aktier"

                # Hvis intet er overstyret i Excel-arket, slår vi det op live på yfinance!
                if not category or not subsector:
                    try:
                        t = yf.Ticker(symbol)
                        info = t.info
                        sec = info.get("sector", "Other")
                        ind = info.get("industry", "Other")
                        
                        cat_detect, sub_detect = temp_screener.map_to_category_and_sector(symbol, sec, ind)
                        category = category if category else cat_detect
                        subsector = subsector if subsector else sub_detect
                    except Exception:
                        category = category if category else "Aktier"
                        subsector = subsector if subsector else "Other"

                # Læg vægtene sammen
                if category in portfolio_distribution:
                    portfolio_distribution[category] += row_weight
                
                if subsector not in sector_distribution:
                    sector_distribution[subsector] = 0.0
                sector_distribution[subsector] += row_weight

            return portfolio_distribution, sector_distribution
        except Exception as e:
            print(f"Fejl under indlæsning: {str(e)}")
            return {k: 0.0 for k in TARGET_PORTFOLIO.keys()}, {}

    def get_current_holdings_details(self) -> list:
        try:
            raw_df = self._read_tab_as_df("Beholdninger")
            df = self._clean_and_align_df(raw_df, "ticker")
            
            ticker_col = self._find_column_by_keyword(df, "ticker")
            position_col = self._find_column_by_keyword(df, "position") or self._find_column_by_keyword(df, "navn")
            mv_col = self._find_column_by_keyword(df, "markedsværdi")
            sektor_col = self._find_column_by_keyword(df, "sektor")
            aktivklasse_col = self._find_column_by_keyword(df, "aktivklasse")
            
            holdings = []
            for _, row in df.iterrows():
                ticker = str(row.get(ticker_col, "")).strip().upper()
                if ticker and not pd.isna(row.get(ticker_col)) and ticker not in ["TICKER", "STATUS", "POSITION", "HULLER"]:
                    holdings.append({
                        "ticker": ticker,
                        "name": str(row.get(position_col, ticker)).strip(),
                        "market_value": str(row.get(mv_col, "0.00 DKK")).strip(),
                        "sector": str(row.get(sektor_col, "N/A")).strip(),
                        "asset_class": str(row.get(aktivklasse_col, "N/A")).strip()
                    })
            return holdings
        except Exception as e:
            print(f"Fejl: {str(e)}")
            return []

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
                        tickers.append(val_str)
            return list(set(tickers))
        except Exception as e:
            return []


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
#  SCREENER & COMPLIANCE AGENT (MED AUTOMATISK ZOYA-CRAWLER & DYN-MAPPING)
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

        # Kontanter / short-term
        if "cash" in sym or "money market" in sec_l:
            return "Kontanter/Private", "Cash Equivalents"
