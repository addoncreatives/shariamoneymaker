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
import yfinance as yf
import requests
import pandas as pd

# =====================================================================
#  KONFIGURATION OG MILJØVARIABLER
# =====================================================================

# Strategiske målvægte for dine kasser i dit Excel-styringsark
TARGET_PORTFOLIO = {
    "Tech & B2B Software": 20.0,
    "Defensivt Forbrug & Healthcare": 20.0,
    "Infrastruktur & Grøn Omstilling": 20.0,
    "Råvarer": 10.0,
    "ETFer & Sukuk": 30.0
}

# Hent ID og API-nøgler fra GitHub Secrets
# Standard-ID er sat til dit delte link: 1EnE2XkQySaGsdaxR5KySZZ924LT66ICo
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1EnE2XkQySaGsdaxR5KySZZ924LT66ICo")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Dine e-mailindstillinger (Hentes fra Secrets, hvis de findes, ellers bruges dine standardadresser)
EMAIL_SENDER = os.getenv("EMAIL_SENDER", "wazir.ilyas@gmail.com")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD") # Skal fortsat ligge sikkert som en GitHub Secret
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER", "addoncreatives@gmail.com")


# =====================================================================
#  GOOGLE SHEETS / EXCEL AGENT (EKSORT FILMETODE)
# =====================================================================
class GoogleSheetsAgent:
    """
    Downloader hele dit Google Sheet / Excel-ark som en .xlsx-fil direkte i hukommelsen.
    Dette sikrer, at vi kan tilgå fanerne fejlfrit uanset filtype.
    """
    def __init__(self, sheet_id: str):
        self.sheet_id = sheet_id

    def _read_tab_as_df(self, tab_name: str) -> pd.DataFrame:
        url = f"https://docs.google.com/spreadsheets/d/{self.sheet_id}/export?format=xlsx"
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            # Indlæser fanen ved hjælp af openpyxl
            df = pd.read_excel(io.BytesIO(response.content), sheet_name=tab_name, engine='openpyxl')
            return df
        except Exception as e:
            raise RuntimeError(
                f"Kunne ikke indlæse fanen '{tab_name}' fra dit Google Sheet link. "
                f"Er delingsindstillingerne sat til 'Alle med linket kan se'? Fejl: {str(e)}"
            )

    def _clean_and_align_df(self, df: pd.DataFrame, key_header_word: str) -> pd.DataFrame:
        """
        Søger ned gennem de første rækker for at finde den rigtige header,
        hvis der er tomme rækker eller titelfelter i toppen af dit Excel-ark.
        """
        # Hvis søgeordet allerede findes i de indlæste kolonner, returneres df direkte
        for col in df.columns:
            if key_header_word.lower() in str(col).lower():
                return df
                
        # Ellers søger vi i de første 10 rækker for at finde header-rækken
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
        """
        Læser fanen 'Beholdninger', identificerer vægtkolonnerne og grupperer efter drivkraft.
        """
        raw_df = self._read_tab_as_df("Beholdninger")
        df = self._clean_and_align_df(raw_df, "drivkraft")
        
        drivkraft_col = self._find_column_by_keyword(df, "drivkraft")
        weight_col = self._find_column_by_keyword(df, "vægt")

        if not drivkraft_col or not weight_col:
            raise KeyError(
                f"Kunne ikke finde kolonnerne 'Drivkraft' og 'Porteføljevægt' i 'Beholdninger'. "
                f"Tjek om de er stavet korrekt i dit ark."
            )

        def clean_weight(val):
            if pd.isna(val) or val == "":
                return 0.0
            val_str = str(val).replace('%', '').replace(',', '.').strip()
            try:
                return float(val_str)
            except ValueError:
                return 0.0

        df['Cleaned_Weight'] = df[weight_col].apply(clean_weight)

        # Gruppér efter drivkraft
        grouped = df.groupby(drivkraft_col)['Cleaned_Weight'].sum().to_dict()

        # Tilpas til TARGET_PORTFOLIO's definerede kategorier
        normalized_portfolio = {}
        for target_key in TARGET_PORTFOLIO.keys():
            sum_val = 0.0
            for g_key, g_val in grouped.items():
                if target_key.lower() in str(g_key).lower() or str(g_key).lower() in target_key.lower():
                    sum_val += g_val
            normalized_portfolio[target_key] = sum_val

        return normalized_portfolio

    def get_watchlist_tickers(self) -> list:
        """
        Læser fanen 'Opsummering', lokaliserer kolonne N (Huller / Watchlist) og udtager rene tickers.
        """
        raw_df = self._read_tab_as_df("Opsummering")
        df = self._clean_and_align_df(raw_df, "huller")
        
        watchlist_col = self._find_column_by_keyword(df, "huller") or self._find_column_by_keyword(df, "watchlist")
        
        tickers = []
        if watchlist_col:
            raw_series = df[watchlist_col]
        else:
            # Falder tilbage til kolonne index 13 (kolonne N) hvis ingen kolonnenavne matcher
            if len(df.columns) >= 14:
                raw_series = df.iloc[:, 13]
            else:
                raise KeyError("Kunne ikke lokalisere kolonne N (Huller / Watchlist) i fanen 'Opsummering'.")

        for val in raw_series:
            if pd.isna(val):
                continue
            val_str = str(val).strip().upper()
            # Tjekker om værdien ligner en legitim børsticker (ingen lange tekster, kun bogstaver, tal, . og -)
            if val_str and len(val_str) < 12 and re.match(r'^[A-Z0-9\.\-]+$', val_str):
                if val_str not in ["TICKER", "STATUS", "POSITION", "HULLER"]:
                    tickers.append(val_str)

        return list(set(tickers))


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
#  SCREENER & COMPLIANCE AGENT
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
            
        return "Infrastruktur & Grøn Omstilling"

    def screen_ticker(self, symbol: str) -> dict:
        try:
            ticker_obj = yf.Ticker(symbol)
            info = ticker_obj.info
            
            if not info:
                return {"symbol": symbol, "passed": False, "reason": "Ingen data fundet på yfinance"}

            sector = info.get("sector", "")
            industry = info.get("industry", "")
            
            for p_sector in self.PROHIBITED_SECTORS:
                if p_sector.lower() in sector.lower():
                    return {"symbol": symbol, "passed": False, "reason": f"Ikke-tilladt sektor: {sector}"}
                    
            for p_ind in self.PROHIBITED_INDUSTRIES:
                if p_ind.lower() in industry.lower():
                    return {"symbol": symbol, "passed": False, "reason": f"Ikke-tilladt branche: {industry}"}

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
#  FUNDAMENTAL AGENT
# =====================================================================
class FundamentalAgent:
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
    def __init__(self, api_key: str):
        self.api_key = api_key
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
            print("E-mail konfiguration mangler eller er ufuldstændig i GitHub Secrets. Udskriver rapporten her:")
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
            print("Succes: Rapport er blevet sendt pr. e-mail.")
        except Exception as e:
            print(f"Fejl ved afsendelse af e-mail: {str(e)}")


# =====================================================================
#  HOVEDKØRSEL (ORCHESTRATOR)
# =====================================================================
def main():
    try:
        print("Henter data fra dit Google Sheet som Excel-projektmappe...")
        
        # 1. Indlæs data via Excel-eksport URL
        sheets_agent = GoogleSheetsAgent(GOOGLE_SHEET_ID)
        
        # Udregn dynamiske vægte fra fanen 'Beholdninger'
        current_portfolio_weights = sheets_agent.get_current_weights()
        print(f"Aktuelle vægte beregnet fra Google Sheet: {current_portfolio_weights}")
        
        # Hent Watchlist-tickers fra 'Opsummering'
        watchlist_tickers = sheets_agent.get_watchlist_tickers()
        print(f"Watchlist-kandidater fundet i Google Sheet: {watchlist_tickers}")

        if not watchlist_tickers:
            print("Advarsel: Ingen gyldige tickers fundet i Watchlist.")

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
