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
# Dette sikrer, at systemet aldrig løber tør for kandidater, selv hvis din fane er tom.
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

# Google Sheet ID (Delt link)
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
if not GOOGLE_SHEET_ID or GOOGLE_SHEET_ID.strip() == "":
    GOOGLE_SHEET_ID = "1EnE2XkQySaGsdaxR5KySZZ924LT66ICo"

# Den nye Gemini 3.5 Flash Model
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
            print(f"Advarsel under indlæsning af 'Beholdninger' (bruger standardvægte): {str(e)}")
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
                        tickers.append(val_str)

            return list(set(tickers))
        except Exception as e:
            print(f"Advarsel under indlæsning af 'Opsummering' (Watchlist): {str(e)}")
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
#  SCREENER & COMPLIANCE AGENT (PROAKTIV OG ETF-FEJLSIKRET)
# =====================================================================
class ScreenerComplianceAgent:
    PROHIBITED_SECTORS = ["Financial Services", "Financial"]
    PROHIBITED_INDUSTRIES = [
        "Banks", "Insurance", "Aerospace & Defense", "Gambling", 
        "Tobacco", "Distillers & Vintners", "Breweries"
    ]

    def __init__(self, tickers: list):
        self.tickers = tickers

    def map_to_category(self, symbol: str, sector: str, industry: str) -> str:
        symbol_upper = symbol.upper()
        sector_lower = sector.lower() if sector else ""
        industry_lower = industry.lower() if industry else ""

        if symbol_upper in ["WPM", "NEM", "GOLD", "AEM"]:
            return "Råvarer"
        if symbol_upper in ["IGDA.L", "SPSK", "HLAL", "UMMA", "ISWD.L", "ISUS.L", "HIWS.L"]:
            return "ETFer & Sukuk"

        if "technology" in sector_lower or "software" in industry_lower:
            return "Tech & B2B Software"
        elif "healthcare" in sector_lower or "defensive" in sector_lower or "medical" in industry_lower:
            return "Defensivt Forbrug & Healthcare"
        elif "industrial" in sector_lower or "utilities" in sector_lower or "energy" in sector_lower:
            return "Infrastruktur & Grøn Omstilling"
        elif "materials" in sector_lower:
            return "Råvarer"
            
        return "Infrastruktur & Grøn Omstilling"

    def screen_ticker(self, symbol: str) -> dict:
        try:
            ticker_obj = yf.Ticker(symbol)
            info = ticker_obj.info
            
            if not info:
                return {"symbol": symbol, "passed": False, "reason": "Ingen data fundet på yfinance"}

            # Tjek om det er en ETF eller en Sukuk-fond (de skal omgå traditionel gældsscreening)
            quote_type = info.get("quoteType", "").upper()
            is_etf = quote_type in ["ETF", "MUTUALFUND"] or symbol in ["IGDA.L", "SPSK", "HLAL", "UMMA", "ISWD.L"]

            if is_etf:
                mapped_cat = "ETFer & Sukuk"
                return {
                    "symbol": symbol,
                    "passed": True,
                    "name": info.get("longName", symbol),
                    "pe_ratio": info.get("trailingPE", "N/A"),
                    "debt_ratio": "N/A (ETF/Sukuk)",
                    "sector": "ETF / Fond",
                    "industry": "ETF",
                    "category": mapped_cat,
                    "is_etf": True
                }

            # Sharia Branche-screening for almindelige aktier
            sector = info.get("sector", "")
            industry = info.get("industry", "")
            
            for p_sector in self.PROHIBITED_SECTORS:
                if p_sector.lower() in sector.lower():
                    return {"symbol": symbol, "passed": False, "reason": f"Ikke-tilladt sektor: {sector}"}
                    
            for p_ind in self.PROHIBITED_INDUSTRIES:
                if p_ind.lower() in industry.lower():
                    return {"symbol": symbol, "passed": False, "reason": f"Ikke-tilladt branche: {industry}"}

            # Gældsscreening (< 30% Gæld til Markedsværdi eller Gæld til Egenkapital)
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
                return {"symbol": symbol, "passed": False, "reason": "Kunne ikke beregne gældskvotient"}

            if debt_ratio_pct > 30.0:
                return {
                    "symbol": symbol, 
                    "passed": False, 
                    "reason": f"Gældskvoten ({debt_ratio_pct:.2f}%) overskrider grænsen på 30% ({method_used})"
                }

            mapped_cat = self.map_to_category(symbol, sector, industry)

            return {
                "symbol": symbol,
                "passed": True,
                "name": info.get("longName", symbol),
                "pe_ratio": info.get("trailingPE", "N/A"),
                "debt_ratio": f"{debt_ratio_pct:.2f}% ({method_used})",
                "sector": sector,
                "industry": industry,
                "category": mapped_cat,
                "is_etf": False
            }

        except Exception as e:
            return {"symbol": symbol, "passed": False, "reason": f"Fejl under screening: {str(e)}"}

    def run_screening(self, target_category: str) -> list:
        approved_stocks = []
        for ticker in self.tickers:
            time.sleep(1.5)
            result = self.screen_ticker(ticker)
            if result["passed"] and result["category"] == target_category:
                approved_stocks.append(result)
        return approved_stocks


# =====================================================================
#  COUNCIL AGENT (GEMINI 3.5 FLASH - PROAKTIV KONSULENT)
# =====================================================================
class CouncilAgent:
    """
    Opretter forbindelse til Gemini 3.5 Flash og udfører den komplette analyse:
    1. Detaljeret konsulent-analyse af de 10 godkendte kandidater.
    2. Den asynkrone LLM-debat (5 rådgivere) baseret på de Top-3 mest lovende aktier.
    3. Formandens endelige beslutningsrapport.
    """
    def __init__(self, api_key: str):
        self.api_key = api_key
        # Vi anvender den nye gemini-3.5-flash model til komplekse, agentiske og dybe analyser
        self.url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={self.api_key}"

    def run_proactive_analysis(self, candidates_data: list, category: str, deficit: float, current_portfolio_str: str) -> str:
        # Konverterer vores rå data til en klar JSON-struktur som Gemini kan forstå fejlfrit
        candidates_json = json.dumps(candidates_data, indent=2, ensure_ascii=False)
        
        prompt = f"""
        Du er en elitesammenslutning af 5 finansielle rådgivere og en formand ("LLM Council"), der fungerer som den personlige analyseafdeling for en langsigtet, muslimsk investor i Norden.
        
        MÅLET MED ANALYSEN:
        Vi skal finde de absolut stærkeste aktiver til at dække et underskud på {deficit:.2f}% i porteføljekassen "{category}".
        
        SITUATIONSBILLEDE (DIN REFERENCE-RAMME):
        Aktuel porteføljefordeling udlæst direkte fra investor's Google Sheet:
        {current_portfolio_str}
        
        DE 10 GODKENDTE KANDIDATER (KLARGJORT VIA DYNAMISK SCREENING):
        Nedenfor er de 10 selskaber, der har bestået alle Sharia- og gældskrav (<30% gældskvote). Her er deres rå nøgletal trukket live fra Yahoo Finance:
        {candidates_json}
        
        DIN OPGAVE (DU MÅ IKKE SPRINGE NOGET OVER ELLER FORKORTE):
        
        ---------------------------------------------------------------------
        DEL 1 — DYBDEGÅENDE KONSULENT-ANALYSE AF DE 10 KANDIDATER
        For hver af de 10 kandidater ovenfor skal du udarbejde en skarp, professionel investeringsanalyse indeholdende:
        
        1. Investeringscase: Hvorfor er denne aktie interessant lige nu i forhold til investors nuværende balance?
        2. Økonomisk gennemgang: Opsummér seneste kvartalsrapport (vækst, marginer, cash flow). Brug din viden kombineret med de leverede rå nøgletal.
        3. Fremtidsudsigter & Pipeline: Hvilke projekter, markeder eller produkter satser selskabet på?
        4. Risikovurdering: Hvad er de største trusler (makro, konkurrenter, renter osv.)?
        5. Grafisk Analyse (Tekstbaseret): Beskriv trenden over de sidste 3 måneder (momentum) og 3 år (langsigtet vækstrejse).
        6. Analytiker-indsigt & Kilder: Indsæt nøjagtigt 2 direkte og klikbare URL-links til troværdige finansielle analyser (f.eks. Seeking Alpha, Yahoo Finance eller Reuters) udgivet inden for de seneste 30 dage. Brug det rigtige ticker-symbol til at forme URL'en (fx `https://seekingalpha.com/symbol/TICKER` eller `https://finance.yahoo.com/quote/TICKER`).
        
        ---------------------------------------------------------------------
        DEL 2 — LLM COUNCIL DEBAT (TOP-3)
        Udvælg de TOP-3 absolut mest lovende kandidater ud af de 10. Kør nu en intens debat mellem dine 5 faste rådgivere baseret udelukkende på disse Top-3 aktiver:
        
        - Rådgiver 1 — THE CONTRARIAN: Ser kun fejl. Leder efter hvad der knækker først. Viser det absolut værste scenarie.
        - Rådgiver 2 — THE FIRST-PRINCIPLES THINKER: Stripper problemet ned til det mest fundamentale. Riber alle antagelser fra hinanden.
        - Rådgiver 3 — THE EXPANSIONIST: Finder det uventede potentiale. Hvad åbner det her op for i det helt store perspektiv?
        - Rådgiver 4 — THE OUTSIDER: Ved intet om branchen. Stiller de åbenlyse, "dumme" spørgsmål, som folk i branchen er holdt op med at stille.
        - Rådgiver 5 — THE EXECUTOR: Går kun op i handling. Hvad skal investoren gøre på mandag kl. 09:00?
        
        Udfør den anonyme peer-review, hvor hver rådgiver kort anmelder og rater de andre 4 anonymiserede svar (Response A, B, C, D) med karakterer fra 1-4.
        
        ---------------------------------------------------------------------
        DEL 3 — FORMANDENS ENDELIGE BESLUTNING (CHAIRMAN'S CALL)
        Formanden træffer nu den endelige beslutning baseret på de 10 analyser og debatten. Skriv en klar, uforbeholden anbefaling:
        - Hvilken specifik aktie (eller ETF) investoren skal købe nu.
        - Den absolut stærkeste grund til beslutningen.
        - Den største risiko, som investoren skal holde skarpt øje med.
        - Det konkrete næste skridt, der skal tages inden for de næste 7 dage.
        
        FORMATERING:
        Skriv det hele i flot, overskuelig Markdown på dansk. Hold tonen skarp, finansielt funderet og analytisk udfordrende.
        """

        headers = {'Content-Type': 'application/json'}
        payload = {
            "contents": [
                {
                    "parts": [{"text": prompt}]
                }
            ]
        }
        
        try:
            response = requests.post(self.url, headers=headers, json=payload, timeout=90)
            response.raise_for_status()
            data = response.json()
            
            if 'candidates' in data and len(data['candidates']) > 0:
                parts = data['candidates'][0]['content']['parts']
                if len(parts) > 0:
                    return parts[0]['text']
            
            return "### Fejl\n*Modtog uventet format fra Gemini 3.5 Flash.*"
        except Exception as e:
            return f"### Systemfejl under kørsel af LLM Council\nFejlbesked: {str(e)}"


# =====================================================================
#  DELIVERY AGENT (SMTP NOTIFIKATION)
# =====================================================================
class DeliveryAgent:
    @staticmethod
    def send_email(subject: str, content: str):
        if not EMAIL_PASSWORD or EMAIL_PASSWORD.strip() == "":
            print("E-mail adgangskode (EMAIL_PASSWORD) mangler i GitHub Secrets. Udskriver rapporten her:")
            print(f"\n=== {subject} ===\n")
            print(content)
            return

        msg = MIMEMultipart()
        msg["From"] = EMAIL_SENDER
        msg["To"] = EMAIL_RECEIVER
        msg["Subject"] = subject
        msg.attach(MIMEText(content, "plain", "utf-8"))

        try:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
            server.quit()
            print("Succes: Strategisk e-mailrapport afsendt.")
        except Exception as e:
            print(f"Fejl under afsendelse af e-mail: {str(e)}")


# =====================================================================
#  ORCHESTRATOR / SYSTEM FLOW
# =====================================================================
def main():
    try:
        print("Henter porteføljedata fra Google Sheet...")
        sheets_agent = GoogleSheetsAgent(GOOGLE_SHEET_ID)
        
        # 1. Hent investors aktuelle vægte
        current_portfolio_weights = sheets_agent.get_current_weights()
        print(f"Beregnet porteføljebalance: {current_portfolio_weights}")
        
        # 2. Hent investors personlige Watchlist
        watchlist_tickers = sheets_agent.get_watchlist_tickers()
        print(f"Investors personlige Watchlist: {watchlist_tickers}")

        # 3. Find den mest undervægtede kasse, som skal have fokus i nat
        pm = PortfolioManagerAgent(current_portfolio_weights, TARGET_PORTFOLIO)
        focus_category, deficit = pm.identify_underweighted_focus()
        print(f"Nattens strategiske fokus: {focus_category} (Underskud: {deficit:.2f}%)")

        # 4. PROAKTIV SØGNING: Kombiner personlig Watchlist med vores globale vækst-pool
        # Dette sikrer, at vi altid har stærke kandidater til screening!
        growth_pool = GLOBAL_COMPLIANT_GROWTH_POOL.get(focus_category, [])
        combined_candidates = list(set(watchlist_tickers + growth_pool))
        print(f"Kombineret søgebase ({len(combined_candidates)} aktiver): {combined_candidates}")

        # 5. Kør compliance screening (Sharia & Gælds-barrierer)
        print("Screener kombineret søgebase mod Sharia- og gældskrav...")
        screener = ScreenerComplianceAgent(combined_candidates)
        approved_stocks = screener.run_screening(focus_category)
        print(f"Godkendte kandidater fundet efter screening: {[s['symbol'] for s in approved_stocks]}")

        # Tag de op til 10 bedste godkendte kandidater til den dybdegående analyse
        target_candidates = approved_stocks[:10]
        
        if not target_candidates:
            raise ValueError(f"Kunne ikke finde nogen godkendte kandidater til kategorien: {focus_category}")

        # 6. Indhent detaljerede yfinance nøgletal for hver af de 10 kandidater
        print("Indhenter detaljerede kvartalstal og finansielle metrics for de 10 kandidater...")
        detailed_candidates_data = []
        for stock in target_candidates:
            symbol = stock["symbol"]
            try:
                t = yf.Ticker(symbol)
                info = t.info
                time.sleep(0.5) # Undgå yfinance blokering
                
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
                    "free_cash_flow": f"{free_cashflow / 1e6:.2f}M DKK/USD" if isinstance(free_cashflow, (int, float)) else "N/A",
                    "current_price": info.get("currentPrice", info.get("regularMarketPrice", "N/A")),
                    "currency": info.get("currency", "N/A")
                })
            except Exception as e:
                print(f"Kunne ikke hente udvidede data for {symbol}: {str(e)}")
                # Fallback til screenede data
                detailed_candidates_data.append(stock)

        # 7. Aktiver Gemini 3.5 Flash til den proaktive strategiske rådgivning
        council_report = "### LLM Council kunne ikke genereres pga. manglende API-nøgle."
        if GEMINI_API_KEY:
            print("Aktiverer Gemini 3.5 Flash til dybdegående investeringsanalyse og rådsdebat...")
            current_portfolio_str = json.dumps(current_portfolio_weights, indent=2, ensure_ascii=False)
            
            council_agent = CouncilAgent(GEMINI_API_KEY)
            council_report = council_agent.run_proactive_analysis(
                candidates_data=detailed_candidates_data,
                category=focus_category,
                deficit=deficit,
                current_portfolio_str=current_portfolio_str
            )

        # 8. Send den færdige rapport afsted via e-mail
        subject = f"[LLM Council] Strategisk Investeringsrapport - Fokus: {focus_category}"
        DeliveryAgent.send_email(subject, council_report)

    except Exception as e:
        error_msg = f"Der opstod en kritisk systemfejl under kørslen af LLM Council-workflowet:\n\n{traceback.format_exc()}"
        print(error_msg, file=sys.stderr)
        DeliveryAgent.send_email("[System Error] LLM Council fejlede", error_msg)


if __name__ == "__main__":
    main()
