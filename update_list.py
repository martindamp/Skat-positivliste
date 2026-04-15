import pandas as pd
import requests
import json
from io import BytesIO
from datetime import datetime
import pytz
from bs4 import BeautifulSoup

# Den officielle landingsside hvor linket findes
SOURCE_PAGE_URL = "https://skat.dk/erhverv/ekapital/vaerdipapirer/beviser-og-aktier-i-investeringsforeninger-og-selskaber-ifpa"

def find_excel_url(page_url):
    """Scraper landingssiden for at finde det aktuelle link til Excel-filen."""
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(page_url, headers=headers, timeout=15)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Vi leder efter alle <a> tags
        links = soup.find_all('a', href=True)
        
        for link in links:
            href = link['href']
            # Vi kigger efter links der slutter på .xlsx og indeholder abis eller liste
            if href.endswith('.xlsx') and ('abis' in href.lower() or 'liste' in href.lower()):
                # Håndter relative links
                if href.startswith('/'):
                    return f"https://skat.dk{href}"
                return href
                
        return None
    except Exception as e:
        print(f"Fejl ved scraping af landingsside: {e}")
        return None

def download_and_convert():
    print(f"Leder efter nyeste Excel-fil på {SOURCE_PAGE_URL}...")
    excel_url = find_excel_url(SOURCE_PAGE_URL)
    
    if not excel_url:
        print("Kritisk fejl: Kunne ikke finde et Excel-link på siden.")
        return

    print(f"Fandt fil: {excel_url}")
    
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(excel_url, headers=headers)
    excel_file = pd.ExcelFile(BytesIO(response.content))
    
    all_sheets_data = []
    # iterer over alle sheets, et for hvert år
    for sheet_name in excel_file.sheet_names:
        # Spring altid forsiden over
        if "forside" in sheet_name.lower(): 
            continue
        
        # Find header-række dynamisk
        raw_df = excel_file.parse(sheet_name, header=None)
        header_row_index = 0
        found_header = False
        
        for i, row in raw_df.iterrows():
            row_str = row.astype(str).values
            if any('ISIN' in s.upper() for s in row_str) or any('NAVN' in s.upper() for s in row_str):
                header_row_index = i
                found_header = True
                break
        
        if found_header:
            df = excel_file.parse(sheet_name, skiprows=header_row_index)
            # Rens kolonnenavne for linjeskift
            df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
            all_sheets_data.append(df)

    if not all_sheets_data:
        print("Ingen data fundet i arkene.")
        return

    # Saml alle ark til ét DataFrame
    full_df = pd.concat(all_sheets_data, ignore_index=True)
    
    # Fjern helt tomme rækker
    full_df = full_df.dropna(how='all')
    
    # Fjern dubletter
    full_df = full_df.drop_duplicates()
    
    # Fjern dubletter baseret på ISIN-kode
    isin_col = next((c for c in full_df.columns if 'ISIN' in str(c).upper()), None)
    if isin_col:
        full_df = full_df.drop_duplicates(subset=[isin_col], keep='last')
    
    # Erstat NaN med tom streng
    full_df = full_df.fillna('')

    # Sæt tidszonen til dansk tid
    dk_tz = pytz.timezone('Europe/Copenhagen')
    now_dk = datetime.now(dk_tz)
    last_updated_str = now_dk.strftime("%d. %B %Y kl. %H:%M")

    output = {
        "metadata": {
            "last_updated": last_updated_str,
            "source_url": SOURCE_PAGE_URL,
            "excel_url": excel_url,
            "description": "Automatisk udtrukket fra Skats IFPA-liste."
        },
        "records": full_df.to_dict(orient='records')
    }

    # Gem som data.json
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    print(f"Succes! Database opdateret kl. {last_updated_str} med {len(output['records'])} unikke rækker.")

if __name__ == "__main__":
    download_and_convert()
