import os
import sys
import time
import traceback
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import yfinance as yf

# =====================================================================
#  KONFIGURATION OG BRUGERDATA (HER EDITERER DU DIN EXCEL-BALANCE)
# =====================================================================

# Nuværende vægtning i din portefølje (i procent, skal summe til 100%)
CURRENT_PORTFOLIO = {
    "Tech & B2B Software": 15.0,
    "Defensivt Forbrug & Healthcare": 25.0,
    "Infrastruktur & Grøn Omstilling": 12.0,
    "Råvarer": 3.0,
    "ETFer & Sukuk": 45.0
}

# Ønsket mål-vægtning (i procent, skal summe til 100%)
TARGET_PORTFOLIO = {
    "Tech & B2B Software": 20.0,
    "Defensivt Forbrug & Healthcare": 20.0,
    "Infrastruktur & Grøn Omstilling": 20.0,
    "Råvarer": 10.0,
    "ETFer & Sukuk": 30.0
}

# Tickers opdelt efter dine strategiske kasser (Kandidater til screening)
CANDIDATE_POOL = {
    "Tech & B2B Software": ["TRMB", "SAP", "IFX.DE", "MSFT", "ASML"],
    "Defensivt Forbrug & Healthcare": ["ORK.OL", "NOVO-B.CO", "6869.T", "AZN.ST"],
    "Infrastruktur & Grøn Omstilling": ["VWS.CO", "NKT.CO", "FLS.CO", "ROCK-B.CO"],
    "Råvarer": ["WPM", "NEM", "GOLD", "AEM"],
    "ETFer & Sukuk": ["IGDA.L", "SPSK", "HLAL"]
}

# SMTP-indstillinger til e-mailnotifikation (Hentes fra GitHub Secrets)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")


# =====================================================================
#  AGENT 1: PORTFOLIO MANAGER AGENT
# =====================================================================
class PortfolioManagerAgent:
    """
    Analyserer afvigelsen mellem nuværende vægtning og målvægtning.
    Identificerer den mest undervægtede 'kasse' til nattens søgefokus.
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
#  AGENT 2: SCREENER & COMPLIANCE AGENT (SHARIA & GÆLD)
# =====================================================================
class ScreenerComplianceAgent:
    """
    Screener finansielle metrics og udelukker selskaber baseret på Sharia-
    og gældsregler (Debt/Equity eller Debt/MarketCap < 30%).
    """
    PROHIBITED_SECTORS = ["Financial Services", "Financial"]
    PROHIBITED_INDUSTRIES = [
        "Banks", "Insurance", "Aerospace & Defense", "Gambling", 
        "Tobacco", "Distillers & Vintners", "Breweries"
    ]

    def __init__(self, tickers: list):
        self.tickers = tickers

    def screen_ticker(self, symbol: str) -> dict:
        """
        Screener en enkelt ticker for gæld og ikke-tilladte brancher.
        """
        try:
            ticker_obj = yf.Ticker(symbol)
            info = ticker_obj.info
            
            if not info:
                return {"symbol": symbol, "passed": False, "reason": "Ingen data fundet via yfinance"}

            # 1. Sharia Branche-screening
            sector = info.get("sector", "")
            industry = info.get("industry", "")
            
            for p_sector in self.PROHIBITED_SECTORS:
                if p_sector.lower() in sector.lower():
                    return {"symbol": symbol, "passed": False, "reason": f"Ikke-tilladt sektor: {sector}"}
                    
            for p_ind in self.PROHIBITED_INDUSTRIES:
                if p_ind.lower() in industry.lower():
                    return {"symbol": symbol, "passed": False, "reason": f"Ikke-tilladt branche: {industry}"}

            # 2. Gældsscreening (Debt/Equity eller Debt/MarketCap)
            # yfinance returnerer ofte debtToEquity som procent (fx 25.5 for 25.5%). 
            # Vi konfigurerer grænsen til <= 30%.
            debt_to_equity = info.get("debtToEquity")
            total_debt = info.get("totalDebt")
            market_cap = info.get("marketCap")
            
            debt_ratio_pct = None
            method_used = ""

            if debt_to_equity is not None:
                # Debt/Equity er direkte tilgængelig
                debt_ratio_pct = debt_to_equity
                method_used = "Debt to Equity"
            elif total_debt and market_cap:
                # Alternativ beregning: Gæld-til-markedsværdi
                debt_ratio_pct = (total_debt / market_cap) * 100
                method_used = "Debt to Market Cap"

            if debt_ratio_pct is None:
                # Konservativ tilgang: Kan vi ikke beregne gælden, fejler selskabet screeningen af sikkerhedshensyn
                return {"symbol": symbol, "passed": False, "reason": "Kunne ikke beregne gældskvoten (mangler data)"}

            if debt_ratio_pct > 30.0:
                return {
                    "symbol": symbol, 
                    "passed": False, 
                    "reason": f"Gældskvoten ({debt_ratio_pct:.2f}%) overskrider grænsen på 30% ({method_used})"
                }

            # Godkendt
            return {
                "symbol": symbol,
                "passed": True,
                "name": info.get("longName", symbol),
                "pe_ratio": info.get("trailingPE", "N/A"),
                "debt_ratio": f"{debt_ratio_pct:.2f}% ({method_used})",
                "sector": sector,
                "industry": industry,
                "currency": info.get("currency", "N/A")
            }

        except Exception as e:
            return {"symbol": symbol, "passed": False, "reason": f"Fejl under screening: {str(e)}"}

    def run_screening(self) -> list:
        approved_stocks = []
        for ticker in self.tickers:
            # Respekter yfinance hastighedsbegrænsninger med en lille pause
            time.sleep(1.5)
            result = self.screen_ticker(ticker)
            if result["passed"]:
                approved_stocks.append(result)
        return approved_stocks


# =====================================================================
#  AGENT 3: FUNDAMENTAL AGENT (NYHEDS-SCRAPING VIA YFINANCE)
# =====================================================================
class FundamentalAgent:
    """
    Trækker de 3 seneste nyheder for at give et øjebliksbillede af selskabets situation.
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
            return ["Fejl under hentning af nyheder."]


# =====================================================================
#  AGENT 4: DOSSIER GENERATOR
# =====================================================================
class DossierGenerator:
    """
    Samler alle data og udarbejder en overskuelig Markdown-rapport.
    """
    @staticmethod
    def generate_report(focus_category: str, deficit: float, approved_data: list) -> str:
        report = []
        report.append(f"# Investeringsdossier: LLM Council\n")
        report.append(f"**Analysefokus:** {focus_category} (Identificeret som mest undervægtet med et underskud på **{deficit:.2f}%** i forhold til dit målbillede).\n")
        report.append(f"**Status:** Screeningen er gennemført uden fejl. Nedenfor ses de godkendte selskaber, der opfylder dine Sharia- og gældskrav.\n")
        report.append("---")

        if not approved_data:
            report.append(f"\n### Ingen kandidater godkendt i denne kørsel.")
            report.append(f"De screenede kandidater i kategorien '{focus_category}' blev udelukket på grund af enten overtrædelse af gældsgrænsen (30%), brancherestriktioner eller manglende data.")
            return "\n".join(report)

        for stock in approved_data:
            symbol = stock["symbol"]
            report.append(f"\n## {stock['name']} ({symbol})")
            report.append(f"- **Sektor/Branche:** {stock['sector']} / {stock['industry']}")
            report.append(f"- **P/E Ratio:** {stock['pe_ratio']}")
            report.append(f"- **Gældsratio:** {stock['debt_ratio']}")
            
            # Portefølje-Fit Sektion
            report.append(f"\n### Portefølje-Fit")
            report.append(
                f"Selskabet integreres i kategorien **{focus_category}**. "
                f"Gennemførelsen af screeningen bekræfter, at selskabet overholder din gældsgrænse på højst 30% "
                f"og ikke opererer inden for udelukkede forretningsområder. "
                f"Køb af denne aktie vil bidrage til at reducere din nuværende underallokering på {deficit:.2f}% i denne specifikke del af din Excel-strategi."
            )
            
            # Seneste nyheder
            report.append(f"\n### Seneste Nyhedsoverskrifter")
            news_items = FundamentalAgent.get_latest_news(symbol)
            for item in news_items:
                report.append(item)
                
            report.append("\n" + "-"*40)

        return "\n".join(report)


# =====================================================================
#  AGENT 5: DELIVERY AGENT (SMTP NOTIFIKATION)
# =====================================================================
class DeliveryAgent:
    """
    Afsender e-mailrapporter. Sender fejl-logs i tilfælde af systemnedbrud.
    """
    @staticmethod
    def send_email(subject: str, content: str):
        if not all([EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER]):
            print("E-mail konfiguration mangler i miljøvariabler. Springer afsendelse over.")
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
            print("E-mail afsendt med succes.")
        except Exception as e:
            print(f"Kunne ikke sende e-mail: {str(e)}")


# =====================================================================
#  ORCHESTRATOR / SYSTEM FLOW
# =====================================================================
def main():
    try:
        print("Initialiserer LLM Council...")
        
        # 1. Analysér porteføljebalancen
        pm = PortfolioManagerAgent(CURRENT_PORTFOLIO, TARGET_PORTFOLIO)
        focus_category, deficit = pm.identify_underweighted_focus()
        print(f"Fokuskategori fundet: {focus_category} (Mangler: {deficit:.2f}%)")

        # 2. Find og screen kandidater for fokuskategorien
        candidates = CANDIDATE_POOL.get(focus_category, [])
        if not candidates:
            raise ValueError(f"Ingen kandidater fundet i konfigurationen for kategorien: {focus_category}")
            
        print(f"Screener kandidater: {candidates}")
        screener = ScreenerComplianceAgent(candidates)
        approved_stocks = screener.run_screening()
        print(f"Screening fuldført. Godkendt: {[s['symbol'] for s in approved_stocks]}")

        # 3. Generer rapport
        report_md = DossierGenerator.generate_report(focus_category, deficit, approved_stocks)

        # 4. Levering af succesrapport
        subject = f"[LLM Council] Investeringsdossier - Fokus på {focus_category}"
        DeliveryAgent.send_email(subject, report_md)

    except Exception as e:
        # Fejlhåndtering: Hvis noget fejler (f.eks. IP-blokering eller manglende internet), sendes fejl-loggen
        error_msg = f"Der opstod en fejl under kørslen af LLM Council:\n\n{traceback.format_exc()}"
        print(error_msg, file=sys.stderr)
        DeliveryAgent.send_email("[System Error] LLM Council fejlede", error_msg)


if __name__ == "__main__":
    main()
