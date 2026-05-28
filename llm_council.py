import os
import sys
import time
import json
import traceback
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import yfinance as yf
import requests
import gspread
import pandas as pd

# =====================================================================
#  KONFIGURATION OG MILJØVARIABLER
# =====================================================================

# Mål-vægtninger (disse kan du beholde her, eller lade systemet sammenligne med)
TARGET_PORTFOLIO = {
    "Tech & B2B Software": 20.0,
    "Defensivt Forbrug & Healthcare": 20.0,
    "Infrastruktur & Grøn Omstilling": 20.0,
    "Råvarer": 10.0,
    "ETFer & Sukuk": 30.0
}

# Hent credentials og API-nøgler fra GitHub Secrets
GOOGLE_SHEETS_CREDENTIALS = os.getenv("GOOGLE_SHEETS_CREDENTIALS")  # JSON-streng fra din Google Service-konto
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")                      # ID'et fra dit Google Sheet URL
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")


# =====================================================================
#  GOOGLE SHEETS INTEGRATION AGENT
# =====================================================================
class GoogleSheetsAgent:
    """
    Forbinder til Google Sheets via gspread ved brug af en Service Account.
    Henter beholdningsdata og Watchlist/Huller dynamisk.
    """
    def __init__(self, credentials_str: str, sheet_id: str):
        self.credentials_str = credentials_str
        self.sheet_id = sheet_id
        self.client = None
        self.spreadsheet = None

    def connect(self):
        if not self.credentials_str or not self.sheet_id:
            raise ValueError("Mangler GOOGLE_SHEETS_CREDENTIALS eller GOOGLE_SHEET_ID i miljøvariablerne.")
        
        try:
            creds_dict = json.loads(self.credentials_str)
            # Log ind via gspreads indbyggede service account login
            self.client = gspread.service_account_from_dict(creds_dict)
            self.spreadsheet = self.client.open_by_key(self.sheet_id)
        except Exception as e:
            raise RuntimeError(f"Kunne ikke oprette forbindelse til Google Sheets: {str(e)}")

    def get_current_weights(self) -> dict:
        """
        Læser fanen 'Beholdninger', grupperer efter 'H: Drivkraft' 
        og beregner de aktuelle porteføljevægte.
        """
        try:
            worksheet = self.spreadsheet.worksheet("Beholdninger")
            data = worksheet.get_all_records()
            df = pd.DataFrame(data)

            # Sikrer at de nødvendige kolonner findes
            required_cols = ['Drivkraft', 'Porteføljevægt']
            for col in required_cols:
                if col not in df.columns:
                    raise KeyError(f"Kolonnen '{col}' blev ikke fundet i fanen 'Beholdninger'.")

            # Rens 'Porteføljevægt'-kolonnen (fjerner %, konverterer til float)
            def clean_weight(val):
                if pd.isna(val) or val == "":
                    return 0.0
                val_str = str(val).replace('%', '').replace(',', '.').strip()
                try:
                    return float(val_str)
                except ValueError:
                    return 0.0

            df['Cleaned_Weight'] = df['Porteføljevægt'].apply(clean_weight)

            # Gruppér efter Drivkraft (som svarer til dine kasser i Excel)
            grouped = df.groupby('Drivkraft')['Cleaned_Weight'].sum().to_dict()

            # Normaliser navne så de matcher TARGET_PORTFOLIO nøgler
            normalized_portfolio = {}
            for k, v in TARGET_PORTFOLIO.items():
                # Lav en simpel case-insensitive søgning
                found_val = 0.0
                for g_key, g_val in grouped.items():
                    if k.lower() in g_key.lower() or g_key.lower() in k.lower():
                        found_val += g_val
                normalized_portfolio[k] = found_val

            return normalized_portfolio

        except Exception as e:
            raise RuntimeError(f"Fejl under behandling af fanen 'Beholdninger': {str(e)}")

    def get_watchlist_tickers(self) -> list:
        """
        Læser fanen 'Opsummering' og isolerer kolonnen 'N: Huller / Watchlist'
        til brug for screening.
        """
        try:
            worksheet = self.spreadsheet.worksheet("Opsummering")
            
            # Da 'Opsummering' kan have tomme celler i toppen eller være asymmetrisk,
            # henter vi værdierne for hele kolonne N direkte
            all_cols = worksheet.get_all_values()
            
            # Find indeks for kolonne N (14. kolonne)
            col_n_index = 13  # 0-baseret indeks for kolonne 14
            
            tickers = []
            for row in all_cols[1:]:  # Spring headeren over
                if len(row) > col_n_index:
                    ticker = row[col_n_index].strip()
                    # Ignorer tomme celler, overskrifter eller forklarende tekst
                    if ticker and len(ticker) < 15 and "/" not in ticker and "Huller" not in ticker:
                        # Rengør tickeren (f.eks. store bogstaver)
                        tickers.append(ticker.upper())
            
            return list(set(tickers)) # Returner unikke tickers
        except Exception as e:
            raise RuntimeError(f"Fejl under behandling af 'Opsummering' (Watchlist): {str(e)}")


# =====================================================================
#  PORTFOLIO MANAGER AGENT
# =====================================================================
class PortfolioManagerAgent:
    """
    Analyserer afvigelsen mellem nuværende vægtning (fra Google Sheets) 
    og målvægtning (fra din model).
    """
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
#  SCREENER & COMPLIANCE AGENT
# =====================================================================
class ScreenerComplianceAgent:
    """
    Validerer selskaber mod Sharia- og gældsregler (Gæld/Egenkapital eller Gæld/MarketCap < 30%).
    Mapper også dynamisk Watchlist-kandidater til de rigtige portefølje-kasser.
    """
    PROHIBITED_SECTORS = ["Financial Services", "Financial"]
    PROHIBITED_INDUSTRIES = [
        "Banks", "Insurance", "Aerospace & Defense", "Gambling", 
        "Tobacco", "Distillers & Vintners", "Breweries"
    ]

    def __init__(self, tickers: list):
        self.tickers = tickers

    def map_to_category(self, symbol: str, sector: str, industry: str) -> str:
        """
        Mapper automatisk en aktie til din Excel-datastrukturs kasser baseret på sektordata.
        """
        symbol_upper = symbol.upper()
        sector_lower = sector.lower() if sector else ""
        industry_lower = industry.lower() if industry else ""

        # Særlige hårde mappings
        if symbol_upper in ["WPM", "NEM", "GOLD", "AEM"]:
            return "Råvarer"
        if symbol_upper in ["IGDA.L", "SPSK", "HLAL"]:
            return "ETFer & Sukuk"

        if "technology" in sector_lower or "software" in industry_lower:
            return "Tech & B2B Software"
        elif "healthcare" in sector_lower or "defensive" in sector_lower or "medical" in industry_lower:
            return "Defensivt Forbrug & Healthcare"
        elif "industrial" in sector_lower or "utilities" in sector_lower or "energy" in sector_lower:
            return "Infrastruktur & Grøn Omstilling"
        elif "materials" in sector_lower:
            return "Råvarer"
            
        return "Infrastruktur & Grøn Omstilling"  # Standard fallback

    def screen_ticker(self, symbol: str) -> dict:
        try:
            ticker_obj = yf.Ticker(symbol)
            info = ticker_obj.info
            
            if not info:
                return {"symbol": symbol, "passed": False, "reason": "Ingen data fundet på yfinance"}

            # 1. Sharia Branche-screening
            sector = info.get("sector", "")
            industry = info.get("industry", "")
            
            for p_sector in self.PROHIBITED_SECTORS:
                if p_sector.lower() in sector.lower():
                    return {"symbol": symbol, "passed": False, "reason": f"Ikke-tilladt sektor: {sector}"}
                    
            for p_ind in self.PROHIBITED_INDUSTRIES:
                if p_ind.lower() in industry.lower():
                    return {"symbol": symbol, "passed": False, "reason": f"Ikke-tilladt branche: {industry}"}

            # 2. Gældsscreening
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
                "category": mapped_cat
            }

        except Exception as e:
            return {"symbol": symbol, "passed": False, "reason": f"Systemfejl under screening: {str(e)}"}

    def run_screening(self, target_category: str) -> list:
        """
        Screener alle kandidater fra din Watchlist, og returnerer dem, 
        der både overholder kravene og passer ind i nattens fokus-kasse.
        """
        approved_stocks = []
        for ticker in self.tickers:
            time.sleep(1.5)
            result = self.screen_ticker(ticker)
            if result["passed"] and result["category"] == target_category:
                approved_stocks.append(result)
        return approved_stocks


# =====================================================================
#  FUNDAMENTAL AGENT
# =====================================================================
class FundamentalAgent:
    """
    Henter de seneste tre nyheder for det screenede selskab.
    """
    @staticmethod
    def get_latest_news(symbol: str) -> list:
        try:
            ticker_obj = yf.Ticker(symbol)
            news = ticker_obj.news
            headlines = []
            if news:
                for item in news[:3]:
                    title = item.get("title", "Ingen titel")
                    link = item.get("link", "#")
                    headlines.append(f"- [{title}]({link})")
                return headlines
            return ["Ingen nyheder fundet for nylig."]
        except Exception:
            return ["Kunne ikke hente nyheder."]


# =====================================================================
#  COUNCIL AGENT (GEMINI INTEGRATION)
# =====================================================================
class CouncilAgent:
    """
    Opretter forbindelse til Google Gemini API (gemini-1.5-flash / gemini-2.5-flash)
    og genererer en dybdegående 5-personers rådsevaluering.
    """
    def __init__(self, api_key: str):
        self.api_key = api_key
        # Vi anvender standarden gemini-1.5-flash, der er hurtig og gratis under AI Studio grænserne
        self.url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={self.api_key}"

    def run_council(self, stock_info: dict, category: str, deficit: float, news: list) -> str:
        news_str = "\n".join(news)
        
        context = f"""
        PROJEKT CONTEXT:
        - Aktiv til evaluering: {stock_info.get('name')} ({stock_info.get('symbol')})
        - Branche/Sektor: {stock_info.get('sector')} / {stock_info.get('industry')}
        - P/E Nøgletal: {stock_info.get('pe_ratio')}
        - Gældskvotient: {stock_info.get('debt_ratio')}
        - Porteføljekasse: {category}
        - Aktuel underallokering i dit Google Sheet: {deficit:.2f}%
        - Seneste nyheder:
        {news_str}
        
        Du rådgiver en langsigtet muslimsk investor i Norden, der følger en striks Sharia- og gældsdisciplin (<30% gæld).
        """

        prompt = f"""
        {context}
        
        I want you to act as a five-person decision council. Do not skip steps. Do not blend the advisers together. Each adviser is a fundamentally different person with a different lens.
        
        Respond entirely in Danish.
        
        STEP 1 — Each adviser answers separately.
        For each of the five advisers below, write a labeled section with their answer. Stay in character. Different language, different priorities, different blind spots.
        
        Adviser 1 — THE CONTRARIAN. Looks only for what will fail. Does not balance. Lists every reason this investment is wrong, what breaks first, and the worst plausible outcome.
        
        Adviser 2 — THE FIRST-PRINCIPLES THINKER. Rips apart my assumptions. Asks what I would do if I couldn't use any obvious framework. Strips the problem down to fundamentals and rebuilds.
        
        Adviser 3 — THE EXPANSIONIST. Finds the upside I'm missing. Looks at the asymmetric outcome if this works. What does the bigger version of this open up?
        
        Adviser 4 — THE OUTSIDER. Knows nothing about my industry. Asks the dumb questions only an outsider asks. Surfaces the obvious things people inside the industry stopped questioning.
        
        Adviser 5 — THE EXECUTOR. Doesn't care about strategy. Cares about Monday morning. Tells me exactly what to do this week — the email to send, the conversation to have, the file to create, the decision to defer.
        
        STEP 2 — Anonymous peer review.
        Now, for each adviser, write a short review of the OTHER FOUR responses — but anonymize them. Refer to them only as "Response A," "Response B," etc. Do not let an adviser know which response is which. Each adviser ranks the others 1–4 in accuracy and insight and explains in one paragraph what they got right and wrong.
        
        STEP 3 — The Chairman's final call.
        Finally, act as the Chairman. You have read all five original answers and all five anonymous reviews. Synthesize a single clear recommendation. No hedging. No "both sides." Tell me:
        - What the right decision actually is
        - The one strongest reason for it
        - The one biggest risk to watch for
        - The specific next step I should take in the next 7 days
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
            response = requests.post(self.url, headers=headers, json=payload, timeout=45)
            response.raise_for_status()
            data = response.json()
            
            if 'candidates' in data and len(data['candidates']) > 0:
                parts = data['candidates'][0]['content']['parts']
                if len(parts) > 0:
                    return parts[0]['text']
            
            return "Fejl: Kunne ikke parse rådets svar korrekt."
        except Exception as e:
            return f"Fejl under generering af LLM Council: {str(e)}"


# =====================================================================
#  DOSSIER GENERATOR & DELIVERY
# =====================================================================
class DossierGenerator:
    @staticmethod
    def generate_report(focus_category: str, deficit: float, approved_data: list, council_report: str = None) -> str:
        report = []
        report.append(f"# Investeringsdossier: LLM Council\n")
        report.append(f"**Analysefokus:** {focus_category} (Dynamisk identificeret via Google Sheets med et underskud på **{deficit:.2f}%**).\n")
        report.append("---")

        if not approved_data:
            report.append(f"\n### Ingen godkendte Watchlist-aktiver fundet i dag.")
            report.append(f"De screenede kandidater fra din fane 'Opsummering' (Kolonne N) passede enten ikke ind i kassen '{focus_category}', eller overholdt ikke dine Sharia- og gældskrav.")
            return "\n".join(report)

        for stock in approved_data:
            symbol = stock["symbol"]
            report.append(f"\n## {stock['name']} ({symbol})")
            report.append(f"- **Sektor/Branche:** {stock['sector']} / {stock['industry']}")
            report.append(f"- **P/E Ratio:** {stock['pe_ratio']}")
            report.append(f"- **Gældsratio:** {stock['debt_ratio']}")
            
            report.append(f"\n### Portefølje-Fit (Excel-integration)")
            report.append(
                f"Selskabet matcher din kasse **{focus_category}**. "
                f"Aktien vil bidrage til at dække din aktuelle underallokering på {deficit:.2f}% registreret i din fane 'Beholdninger'."
            )
            
            report.append(f"\n### Seneste Nyhedsoverskrifter")
            news_items = FundamentalAgent.get_latest_news(symbol)
            for item in news_items:
                report.append(item)
            
            if council_report:
                report.append(f"\n\n## 🗳️ LLM Council Beslutningsrapport")
                report.append(council_report)
                
            report.append("\n" + "-"*40)

        return "\n".join(report)


class DeliveryAgent:
    @staticmethod
    def send_email(subject: str, content: str):
        if not all([EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER]):
            print(f"--- {subject} ---")
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
            print("Rapport sendt pr. e-mail.")
        except Exception as e:
            print(f"E-mail fejl: {str(e)}")


# =====================================================================
#  HOVEDKØRSEL (ORCHESTRATOR)
# =====================================================================
def main():
    try:
        print("Henter data fra dit Google Sheet...")
        
        # 1. Forbind til Google Sheets og læs data
        sheets_agent = GoogleSheetsAgent(GOOGLE_SHEETS_CREDENTIALS, GOOGLE_SHEET_ID)
        sheets_agent.connect()
        
        # Udregn dynamiske vægte fra fanen 'Beholdninger'
        current_portfolio_weights = sheets_agent.get_current_weights()
        print(f"Aktuelle vægte beregnet: {current_portfolio_weights}")
        
        # Hent Watchlist-tickers fra 'Opsummering' (Kolonne N)
        watchlist_tickers = sheets_agent.get_watchlist_tickers()
        print(f"Watchlist-kandidater fundet: {watchlist_tickers}")

        # 2. Find mest undervægtede kategori
        pm = PortfolioManagerAgent(current_portfolio_weights, TARGET_PORTFOLIO)
        focus_category, deficit = pm.identify_underweighted_focus()
        print(f"Fokus i nat: {focus_category} (Underskud: {deficit:.2f}%)")

        # 3. Kør screener på alle Watchlist-tickers
        print("Screener kandidater mod Sharia- og gældsregler...")
        screener = ScreenerComplianceAgent(watchlist_tickers)
        approved_stocks = screener.run_screening(focus_category)
        
        # 4. Kør LLM Council på den øverste godkendte aktie
        council_report = None
        if approved_stocks and GEMINI_API_KEY:
            target_stock = approved_stocks[0]
            print(f"Starter LLM Council på: {target_stock['symbol']}...")
            
            news = FundamentalAgent.get_latest_news(target_stock["symbol"])
            council_agent = CouncilAgent(GEMINI_API_KEY)
            
            council_report = council_agent.run_council(
                stock_info=target_stock,
                category=focus_category,
                deficit=deficit,
                news=news
            )
        elif not GEMINI_API_KEY:
            print("GEMINI_API_KEY mangler. Springer over LLM Council.")

        # 5. Generer og afsend rapport
        report_md = DossierGenerator.generate_report(focus_category, deficit, approved_stocks, council_report)
        subject = f"[LLM Council] Strategisk Rapport - Fokus: {focus_category}"
        DeliveryAgent.send_email(subject, report_md)

    except Exception as e:
        error_msg = f"Der opstod en systemfejl i LLM Council-workflowet:\n\n{traceback.format_exc()}"
        print(error_msg, file=sys.stderr)
        DeliveryAgent.send_email("[System Error] LLM Council fejlede", error_msg)


if __name__ == "__main__":
    main()
