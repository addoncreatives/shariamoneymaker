import os
import sys
import time
import io
import re
import json
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
    subprocess.check_call([sys.executable, "-m", "pip", "install", "edge-tts"])
    import edge_tts

import yfinance as yf
import requests
import pandas as pd

# =====================================================================
#  KONFIGURATION OG MILJØVARIABLER (STRATEGISK MÅLBILLEDE)
# =====================================================================

TARGET_PORTFOLIO = {
    "Aktier": 25.0,
    "Sukuk": 25.0,
    "Råvarer": 25.0,
    "Kontanter/Private": 25.0
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
        "SPSK"
    ]
}

GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1EnE2XkQySaGsdaxR5KySZZ924LT66ICo")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

EMAIL_SENDER = os.getenv("EMAIL_SENDER", "wazir.ilyas@gmail.com")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER", "addoncreatives@gmail.com")


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

    def get_current_weights_and_sectors(self) -> tuple:
        """
        Læser fanen 'Beholdninger', identificerer vægtene og fordeler dem 
        både på 4x25% hovedkasserne og på dine 21 specifikke delsektorer.
        """
        try:
            raw_df = self._read_tab_as_df("Beholdninger")
            df = self._clean_and_align_df(raw_df, "ticker")
            
            drivkraft_col = self._find_column_by_keyword(df, "drivkraft")
            aktivklasse_col = self._find_column_by_keyword(df, "aktivklasse")
            sektor_col = self._find_column_by_keyword(df, "sektor")
            weight_col = self._find_column_by_keyword(df, "vægt")
            mv_col = self._find_column_by_keyword(df, "markedsværdi")

            if not aktivklasse_col or not weight_col:
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

            # 1. Beregn 4x25% Hovedkasser
            portfolio_distribution = {k: 0.0 for k in TARGET_PORTFOLIO.keys()}
            # 2. Beregn 21 Delsektorer
            sector_distribution = {s: 0.0 for s in TARGET_SUBSECTORS}

            for _, row in df.iterrows():
                row_weight = row['Cleaned_Weight']
                if row_weight <= 0.0:
                    continue

                drivkraft_val = normalize_string(row.get(drivkraft_col, "")) if drivkraft_col else ""
                aktivklasse_val = normalize_string(row.get(aktivklasse_col, ""))
                sektor_val = normalize_string(row.get(sektor_col, "")) if sektor_col else ""

                combined_text = f"{drivkraft_val} {aktivklasse_val} {sektor_val}"

                # 4x25% Fordeling
                if "sukuk" in combined_text:
                    portfolio_distribution["Sukuk"] += row_weight
                elif any(word in combined_text for word in ["råvarer", "ravarer", "guld", "gold", "commodities"]):
                    portfolio_distribution["Råvarer"] += row_weight
                elif any(word in combined_text for word in ["kontant", "cash", "private"]):
                    portfolio_distribution["Kontanter/Private"] += row_weight
                else:
                    portfolio_distribution["Aktier"] += row_weight

                # 21 Delsektor-fordeling
                matched_sector = False
                for target_s in TARGET_SUBSECTORS:
                    norm_target = normalize_string(target_s)
                    parts = [normalize_string(p) for p in re.split(r'[/|]', target_s)]
                    if norm_target in combined_text or any(p in combined_text for p in parts if p):
                        sector_distribution[target_s] += row_weight
                        matched_sector = True
                        break

            return portfolio_distribution, sector_distribution
        except Exception as e:
            print(f"Fejl under indlæsning: {str(e)}")
            return {k: 0.0 for k in TARGET_PORTFOLIO.keys()}, {s: 0.0 for s in TARGET_SUBSECTORS}

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
#  PORTFOLIO MANAGER AGENT
# =====================================================================
class PortfolioManagerAgent:
    def __init__(self, current: dict, target: dict):
        self.current = current
        self.target = target

    def identify_underweighted_focus(self) -> tuple:
        max_deficit = -999.0
        focus_category = None
        for category, target_val in self.target.items():
            curr_val = self.current.get(category, 0.0)
            deficit = target_val - curr_val
            if deficit > max_deficit:
                max_deficit = deficit
                focus_category = category
        return focus_category, max_deficit


# =====================================================================
#  SCREENER & COMPLIANCE AGENT (MAPPING AF DELSEKTORER)
# =====================================================================
class ScreenerComplianceAgent:
    PROHIBITED_SECTORS = ["Financial Services", "Financial"]
    PROHIBITED_INDUSTRIES = [
        "Banks", "Insurance", "Aerospace & Defense", "Gambling", 
        "Tobacco", "Distillers & Vintners", "Breweries"
    ]

    def __init__(self, tickers: list):
        self.tickers = tickers

    def map_to_category_and_sector(self, symbol: str, sector: str, industry: str) -> tuple:
        """
        Mapper selskabet til både 4x25% hovedklassen og en af dine 21 delsektorer.
        """
        sym = symbol.upper()
        sec_l = sector.lower() if sector else ""
        ind_l = industry.lower() if industry else ""

        # 1. Sukuk
        if "sukuk" in sym or sym in ["SPSK", "SKUK"]:
            return "Sukuk", "Sukuk"
            
        # 2. Råvarer
        if sym in ["WPM", "FNV", "RGLD"]:
            return "Råvarer", "Mining royalty"
        if sym in ["NEM", "GOLD", "AEM", "BHP", "RIO", "FCX", "VALE"] or \
           any(w in ind_l for w in ["gold", "silver", "precious metals", "copper", "aluminum"]):
            return "Råvarer", "Industrielle metaller / kobber"

        # 3. Aktier
        if sym == "VWS.CO" or "wind" in ind_l:
            return "Aktier", "Vind"
        if sym == "NKT.CO" or "cable" in ind_l or "electrical" in ind_l:
            return "Aktier", "Elektrificering"
        if sym in ["ASML", "NVDA", "IFX.DE"] or "semiconductor" in ind_l:
            return "Aktier", "Halvleder"
        if sym in ["NOVO-B.CO", "AZN.ST", "LLY"] or "pharmaceutical" in ind_l or "biotechnology" in ind_l:
            return "Aktier", "Pharma"
        if sym in ["6869.T", "REGN", "ISRG", "VRTX", "SYK", "MDT"] or "medical" in ind_l:
            return "Aktier", "Medico"
        if sym in ["SAP", "MSFT", "ADBE", "CRM", "SNPS", "ANSS", "NOW", "PANW", "FTNT", "ORCL"] or "software" in ind_l or "software" in sec_l:
            return "Aktier", "Software / platform"
        if sym in ["ROCK-B.CO"] or "building" in ind_l or "construction" in ind_l:
            return "Aktier", "Byggeri"
        if sym in ["FLS.CO", "ABB", "SIE.DE", "GE"] or "machinery" in ind_l or "industrials" in sec_l:
            return "Aktier", "Industri"
        if "chemical" in ind_l or "materials" in sec_l:
            return "Aktier", "Materialer"
        if "infrastructure" in ind_l or "utilities" in sec_l:
            return "Aktier", "Infrastruktur"
        if sym in ["ORK.OL", "NSRGY"] or "food" in ind_l or "consumer defensive" in sec_l:
            return "Aktier", "Consumer staples / dagligvarer"
        if "fertilizer" in ind_l or "agriculture" in ind_l:
            return "Aktier", "Landbrug / gødning"
        if "logistics" in ind_l or "shipping" in ind_l:
            return "Aktier", "Logistik"
            
        # Global/Regional ETF'er
        if sym in ["IGDA.L", "ISWD.L", "UMMA"]:
            return "Aktier", "ETF - global"
        if sym in ["HLAL", "ISUS.L", "MSAU.L"]:
            return "Aktier", "ETF - regional"

        return "Aktier", "Industri"

    def screen_ticker(self, symbol: str) -> dict:
        try:
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
            result = self.screen_ticker(ticker)
            if result["passed"] and result["category"] == target_category:
                approved_stocks.append(result)
        return approved_stocks


# =====================================================================
#  COUNCIL AGENT (GEMINI 3.5 FLASH - FULDT STYLET HTML RAPPORT)
# =====================================================================
class CouncilAgent:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={self.api_key}"

    def run_proactive_analysis(self, candidates_data: list, category: str, deficit: float, current_portfolio_str: str, current_holdings_str: str, sector_distribution_str: str) -> str:
        candidates_json = json.dumps(candidates_data, indent=2, ensure_ascii=False)
        
        prompt = f"""
        Du er en elitesammenslutning af 5 finansielle rådgivere og en formand ("LLM Council"), der rådgiver en Sharia-bevidst nordisk investor.
        
        SITUATIONSBILLEDE & INVESTORPROFIL:
        - Strategisk overordnet model: 4x25% (Aktier, Sukuk, Råvarer, Kontanter/Private) [3].
        - Under evaluering i nat: {category} (Mangler: {deficit:.2f}%) [3].
        - Aktuelle overordnede vægtninger: {current_portfolio_str} [3].
        
        INVESTORS 21 STRATEGISKE DELSEKTORER (ALLOKERING ER DYNAMISK BEREGNET FRA SHEET):
        {sector_distribution_str}
        
        INVESTORS REELLE BEHOLDNINGER:
        {current_holdings_str}
        
        NYE SCREENEDE KANDIDATER KLAR TIL EVALUERING:
        {candidates_json}
        
        DIN OPGAVE (LEVER RAPPORTEN SOM REN HTML MED INLINE-CSS):
        Generer en dyb, fængslende og professionel HTML-investeringsrapport. 
        Brug udelukkende inline CSS-styling til e-mails. Anvend en minimalistisk palette med mørk slate overskrifter (`#0F172A`) og dæmpede guld-detaljer (`#C5A880`) [3].
        
        DEL 1 — DELSEKTOR DIAGNOSE & INDIREKTE EKSPONERING
        Analyser investors beholdninger mod de 21 delsektorer [3]. Ejer investor f.eks. selskaber (som FLSmidth eller NKT), der reelt allerede giver indirekte beskyttelse/råvareeksponering? Er visse delsektorer tomme? Diskuter Saxo Bank og Sharia som en begrænsende barriere.
        
        DEL 2 — KONSULENT-ANALYSE AF DE NYE KANDIDATER (OP TIL 10 SCREENEDE)
        For hver kandidat skal du udarbejde: Investeringscase, Økonomisk gennemgang, Pipeline/Udsigter, Risici, Grafisk analyse (momentum og langsigtet kurve), samt præcis 2 klikbare links (fx til Seeking Alpha: `<a href="https://seekingalpha.com/symbol/TICKER" style="color:#C5A880;text-decoration:none;">Seeking Alpha</a>`).
        
        DEL 3 — LLM COUNCIL DEBAT (TOP-3)
        Udvælg de Top-3 mest lovende aktiver. Kør en dyb, intens og karakterdreven debat baseret på dine 5 rådgivere (Contrarian, First-Principles, Expansionist, Outsider, Executor) [3]. Sørg for at bruge de farvekodede venstregrund-rammer for hver rådgivers boks.
        
        DEL 4 — FORMANDENS AFGØRENDE DEKRET
        Formanden tager den endelige beslutning i en guld-indrammet boks (`border-left: 6px solid #C5A880; background: #FDFBF7;`) [3]. Specificer de næste præcise 7-dages handlingstrin på Saxo Investor [3].
        
        Returner KUN den rå HTML-kode, uden nogen "```html" eller kommentarer uden for HTML-koden.
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
            return "<h3>Fejl</h3><p>Modtog uventet format fra Gemini.</p>"
        except Exception as e:
            return f"<h3>Systemfejl</h3><p>{str(e)}</p>"


# =====================================================================
#  PODCAST AGENT (FREE NEURAL TTS COPMILER)
# =====================================================================
class PodcastAgent:
    """
    Beder Gemini omskrive HTML-rapporten til et samtalemanuskript (JSON),
    og genererer derefter en ægte flerstems MP3-podcast ved brug af edge-tts.
    """
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.gemini_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={self.api_key}"

    def generate_script(self, report_html: str) -> list:
        prompt = f"""
        Du er en elitesurfer og podcast-producer. Omskriv denne HTML-investeringsrapport til et skarpt, utroligt medrivende og naturligt samtalemanuskript på DANSK:
        {report_html}
        
        KRAV TIL ROLLERNE:
        - "Vært": En nysgerrig finansvært.
        - "Contrarian": Den skeptiske, forsigtige djævlens advokat fra rådet [3].
        - "Formand": Den vise formand, der skærer igennem [3].
        
        Dialogen skal flyde som en Bloomberg-podcast (talesprog, korte bemærkninger, let ping-pong). De skal diskutere Saxo-begrænsninger, Sharia-regler, de 21 delsektorer, og de konkrete top-aktier [3].
        
        Du SKAL levere resultatet som en rå, valid JSON-liste af ordbøger uden forklaringer før eller efter.
        Eksempel:
        [
          {{"speaker": "Vært", "text": "Velkommen til rådets ugentlige briefing..."}},
          {{"speaker": "Contrarian", "text": "Ja, og lad os nu lige trække vejret..."}},
          {{"speaker": "Formand", "text": "Formandskabet har vurderet tallene..."}}
        ]
        """
        headers = {'Content-Type': 'application/json'}
        payload = {"contents": [{"parts": [{"text": prompt}]}]}
        try:
            response = requests.post(self.gemini_url, headers=headers, json=payload, timeout=90)
            response.raise_for_status()
            data = response.json()
            if 'candidates' in data and len(data['candidates']) > 0:
                raw_text = data['candidates'][0]['content']['parts'][0]['text']
                raw_text = raw_text.replace("```json", "").replace("```", "").strip()
                return json.loads(raw_text)
            return []
        except Exception as e:
            print(f"Fejl under manuskriptgenerering: {str(e)}")
            return []

    async def compile_audio(self, script_json: list, output_filename: str = "llm_council_podcast.mp3"):
        voice_map = {
            "Vært": "da-DK-JeppeNeural",
            "Contrarian": "da-DK-JeppeNeural",
            "Formand": "da-DK-ChristelNeural"
        }
        print("Kompilerer lydfiler via edge-tts...")
        combined_audio = b""
        for turn in script_json:
            speaker = turn.get("speaker", "Vært")
            text = turn.get("text", "")
            voice = voice_map.get(speaker, "da-DK-JeppeNeural")
            
            rate = "-5%" if speaker == "Contrarian" else "+0%"
            communicate = edge_tts.Communicate(text, voice, rate=rate)
            async for chunk in communicate.stream():
                if chunk["type"] == "audio":
                    combined_audio += chunk["data"]
            time.sleep(0.3)
            
        with open(output_filename, "wb") as f:
            f.write(combined_audio)
        print(f"Podcast klar: '{output_filename}'")


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
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
            server.quit()
            print("Succes: Rapport og podcast-afspilning sendt til din indbakke.")
        except Exception as e:
            print(f"E-mail fejl: {str(e)}")


# =====================================================================
#  ORCHESTRATOR / SYSTEM FLOW
# =====================================================================
def main():
    try:
        print("Opstarter LLM Council 4x25% med delsektorer og Podcast...")
        sheets_agent = GoogleSheetsAgent(GOOGLE_SHEET_ID)
        
        # 1. Hent investors aktuelle vægte samt de 21 delsektorer
        current_portfolio_weights, sector_distribution = sheets_agent.get_current_weights_and_sectors()
        print(f"Beregnet 4x25% hovedfordeling: {current_portfolio_weights}")
        print(f"Beregnet delsektor-fordeling: {sector_distribution}")
        
        # 2. Hent de konkrete positioner
        current_holdings = sheets_agent.get_current_holdings_details()
        print(f"Hentet {len(current_holdings)} konkrete positioner.")
        
        # 3. Hent Watchlist-tickers
        watchlist_tickers = sheets_agent.get_watchlist_tickers()
        print(f"Watchlist: {watchlist_tickers}")

        # 4. Find den mest undervægtede kasse i din 4x25%-struktur
        pm = PortfolioManagerAgent(current_portfolio_weights, TARGET_PORTFOLIO)
        focus_category, deficit = pm.identify_underweighted_focus()
        print(f"Nattens strategiske fokus: {focus_category} (Underskud: {deficit:.2f}%)")

        # 5. PROAKTIV SØGNING: Kombiner personlig Watchlist med vores globale vækst-pool
        growth_pool = GLOBAL_COMPLIANT_GROWTH_POOL.get(focus_category, [])
        combined_candidates = list(set(watchlist_tickers + growth_pool))
        print(f"Kombineret søgebase ({len(combined_candidates)} aktiver): {combined_candidates}")

        # 6. Kør compliance screening (Sharia & Gælds-barrierer)
        print("Screener kandidater mod Sharia- og gældskrav...")
        screener = ScreenerComplianceAgent(combined_candidates)
        approved_stocks = screener.run_screening(focus_category)
        print(f"Godkendte kandidater fundet efter screening: {[s['symbol'] for s in approved_stocks]}")

        target_candidates = approved_stocks[:10]
        
        if not target_candidates:
            error_html = f"<h2>Ingen godkendte kandidater fundet i kategorien {focus_category}</h2>"
            DeliveryAgent.send_email(f"[LLM Council] Alert - Ingen godkendte kandidater i {focus_category}", error_html)
            return

        # 7. Indhent detaljerede yfinance tal
        print("Indhenter kvartalstal for de 10 godkendte kandidater...")
        detailed_candidates_data = []
        for stock in target_candidates:
            symbol = stock["symbol"]
            try:
                t = yf.Ticker(symbol)
                info = t.info
                time.sleep(0.5)
                
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
                    "free_cash_flow": f"{free_cashflow / 1e6:.2f}M USD/DKK" if isinstance(free_cashflow, (int, float)) else "N/A",
                    "current_price": info.get("currentPrice", info.get("regularMarketPrice", "N/A")),
                    "currency": info.get("currency", "N/A")
                })
            except Exception as e:
                detailed_candidates_data.append(stock)

        # 8. Aktiver Gemini 3.5 Flash til den proaktive HTML-rapport
        council_report_html = "<h3>LLM Council fejl</h3>"
        output_mp3 = "llm_council_podcast.mp3"
        podcast_compiled = False

        if GEMINI_API_KEY:
            print("Aktiverer Gemini 3.5 Flash...")
            current_weights_str = json.dumps(current_portfolio_weights, indent=2, ensure_ascii=False)
            current_holdings_str = json.dumps(current_holdings, indent=2, ensure_ascii=False)
            sector_distribution_str = json.dumps(sector_distribution, indent=2, ensure_ascii=False)
            
            council_agent = CouncilAgent(GEMINI_API_KEY)
            council_report_html = council_agent.run_proactive_analysis(
                candidates_data=detailed_candidates_data,
                category=focus_category,
                deficit=deficit,
                current_portfolio_str=current_weights_str,
                current_holdings_str=current_holdings_str,
                sector_distribution_str=sector_distribution_str
            )
            
            # 9. Generer automatisk lyd-podcast
            print("Igangsætter podcast-produktion...")
            podcast_agent = PodcastAgent(GEMINI_API_KEY)
            script = podcast_agent.generate_script(council_report_html)
            if script:
                try:
                    asyncio.run(podcast_agent.compile_audio(script, output_mp3))
                    podcast_compiled = True
                except Exception as tts_err:
                    print(f"Fejl under lydkompilering: {str(tts_err)}")
        else:
            print("Advarsel: GEMINI_API_KEY mangler.")

        # 10. Send HTML-rapport og MP3-podcast til din indbakke
        subject = f"[LLM Council] Strategisk Rapport - Fokus: {focus_category}"
        DeliveryAgent.send_email(
            subject=subject, 
            html_content=council_report_html, 
            attachment_path=output_mp3 if podcast_compiled else None
        )

    except Exception as e:
        error_msg = f"Kritisk systemfejl under kørslen:\n\n{traceback.format_exc()}"
        print(error_msg, file=sys.stderr)
        
        error_html = f"""
        <div style="font-family: sans-serif; max-width: 600px; margin: 0 auto; padding: 25px; border: 1px solid #FCA5A5; background-color: #FEF2F2; border-radius: 8px;">
            <h2 style="color: #B91C1C; margin-top: 0; border-bottom: 2px solid #FCA5A5; padding-bottom: 10px;">🚨 Systemfejl i LLM Council</h2>
            <pre style="background-color: #FFFFFF; border: 1px solid #FEE2E2; padding: 15px; border-radius: 5px; overflow-x: auto; color: #991B1B; font-size: 13px;">{error_msg}</pre>
        </div>
        """
        DeliveryAgent.send_email("[System Error] LLM Council fejlede", error_html)


if __name__ == "__main__":
    main()
