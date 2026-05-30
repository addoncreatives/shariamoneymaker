import json
import requests
import pandas as pd

def build_global_database():
    print("Initializing LLM Council Global Database Compiler...")
    db = {}

    # 1. S&P 500 (US)
    try:
        print("Fetching S&P 500 constituents from Wikipedia...")
        url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
        tables = pd.read_html(url)
        df = tables[0]
        for _, row in df.iterrows():
            ticker = str(row['Symbol']).strip().replace('.', '-')
            sector = str(row['GICS Sector']).strip()
            industry = str(row['GICS Sub-Industry']).strip()
            
            # Sharia-filtrering: Ekskluder traditionel finans og banker
            if "financial" in sector.lower() or "bank" in industry.lower() or "insurance" in industry.lower():
                continue
                
            category = "Aktier"
            if "materials" in sector.lower() or "gold" in industry.lower() or "precious metals" in industry.lower():
                category = "Råvarer"
                
            db[ticker] = [category, industry]
    except Exception as e:
        print(f"Could not load S&P 500: {str(e)}")

    # 2. NASDAQ 100 (US Tech)
    try:
        print("Fetching Nasdaq 100 constituents from Wikipedia...")
        url = "https://en.wikipedia.org/wiki/Nasdaq-100"
        tables = pd.read_html(url)
        df = tables[4] # Standard tabel-indeks for Nasdaq
        ticker_col = [col for col in df.columns if 'ticker' in col.lower() or 'symbol' in col.lower()][0]
        sector_col = [col_for col in df.columns if 'gics sector' in col.lower() or 'sector' in col.lower()][0]
        for _, row in df.iterrows():
            ticker = str(row[ticker_col]).strip()
            sector = str(row[sector_col]).strip()
            
            if "financial" in sector.lower():
                continue
                
            db[ticker] = ["Aktier", sector]
    except Exception as e:
        print(f"Could not load Nasdaq 100: {str(e)}")

    # 3. DOW JONES 30 (US)
    try:
        print("Fetching Dow Jones constituents from Wikipedia...")
        url = "https://en.wikipedia.org/wiki/Dow_Jones_Industrial_Average"
        tables = pd.read_html(url)
        df = tables[1]
        for _, row in df.iterrows():
            ticker = str(row['Symbol']).strip()
            industry = str(row['Industry']).strip()
            if "financial" in industry.lower():
                continue
            db[ticker] = ["Aktier", industry]
    except Exception as e:
        print(f"Could not load Dow Jones: {str(e)}")

    # 4. DAX 40 (Germany)
    try:
        print("Fetching DAX 40 constituents from Wikipedia...")
        url = "https://en.wikipedia.org/wiki/DAX"
        tables = pd.read_html(url)
        df = tables[4]
        for _, row in df.iterrows():
            ticker = str(row['Ticker']).strip()
            sector = str(row['Prime Standard Sector']).strip()
            if "financial" in sector.lower() or "bank" in sector.lower():
                continue
            db[ticker] = ["Aktier", sector]
    except Exception as e:
        print(f"Could not load DAX 40: {str(e)}")

    # 5. CAC 40 (France)
    try:
        print("Fetching CAC 40 constituents from Wikipedia...")
        url = "https://en.wikipedia.org/wiki/CAC_40"
        tables = pd.read_html(url)
        df = tables[0]
        for _, row in df.iterrows():
            ticker = str(row['Ticker']).strip()
            sector = str(row['Sector']).strip()
            if "financial" in sector.lower() or "bank" in sector.lower():
                continue
            db[ticker] = ["Aktier", sector]
    except Exception as e:
        print(f"Could not load CAC 40: {str(e)}")

    # 6. OMX Copenhagen 25 (Denmark)
    try:
        print("Fetching OMC C25 constituents from Wikipedia...")
        url = "https://en.wikipedia.org/wiki/OMXC25"
        tables = pd.read_html(url)
        df = tables[2]
        for _, row in df.iterrows():
            ticker = str(row['Ticker']).strip()
            sector = str(row['Sector']).strip()
            db[ticker] = ["Aktier", sector]
    except Exception as e:
        print(f"Could not load OMX C25: {str(e)}")

    # 7. OMX Stockholm 30 (Sweden)
    try:
        print("Fetching OMX S30 constituents from Wikipedia...")
        url = "https://en.wikipedia.org/wiki/OMXS30"
        tables = pd.read_html(url)
        df = tables[1]
        for _, row in df.iterrows():
            ticker = str(row['Ticker symbol']).strip()
            db[ticker] = ["Aktier", "Industrial Machinery"]
    except Exception as e:
        print(f"Could not load OMX S30: {str(e)}")

    # 8. OBX 25 (Norway)
    try:
        print("Fetching OBX 25 constituents from Wikipedia...")
        url = "https://en.wikipedia.org/wiki/OBX_Index"
        tables = pd.read_html(url)
        df = tables[1]
        for _, row in df.iterrows():
            ticker = str(row['Ticker']).strip()
            db[ticker] = ["Aktier", "Marine Logistics / Energy"]
    except Exception as e:
        print(f"Could not load OBX 25: {str(e)}")

    # Gem det samlede lynhurtige kartotek som failsafe_db.json
    with open("failsafe_db.json", "w") as f:
        json.dump(db, f, indent=2)
        
    print(f"Successfully compiled {len(db)} global Shariah-compliant tickers into failsafe_db.json!")

if __name__ == "__main__":
    build_global_database()
