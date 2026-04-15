import pandas as pd
import requests
import json
from io import BytesIO

URL = "https://skat.dk/media/5bldctyv/marts-2026-abis-liste-2021-2026.xlsx"

def download_and_convert():
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(URL, headers=headers)
    excel_file = pd.ExcelFile(BytesIO(response.content))
    
    all_sheets_data = []
    
    for sheet_name in excel_file.sheet_names:
        if "forside" in sheet_name.lower():
            continue
            
        # Læs arket og find overskrifterne automatisk
        raw_df = excel_file.parse(sheet_name, header=None)
        header_row_index = 0
        for i, row in raw_df.iterrows():
            row_str = row.astype(str).values
            if any('ISIN' in s.upper() for s in row_str) or any('NAVN' in s.upper() for s in row_str):
                header_row_index = i
                break
        
        df = excel_file.parse(sheet_name, skiprows=header_row_index)
        all_sheets_data.append(df)

    # Saml alle ark til ét DataFrame
    full_df = pd.concat(all_sheets_data, ignore_index=True)
    
    # --- RENSNING AF DUBLETTER START ---
    
    # 1. Fjern helt tomme rækker
    full_df = full_df.dropna(how='all')
    
    # 2. Fjern rækker der er 100% identiske
    full_df = full_df.drop_duplicates()
    
    # 3. Fjern dubletter baseret på ISIN-kode (hvis ISIN findes)
    # Vi beholder den sidste forekomst, da den ofte er den mest opdaterede
    isin_col = next((c for c in full_df.columns if 'ISIN' in str(c).upper()), None)
    if isin_col:
        full_df = full_df.drop_duplicates(subset=[isin_col], keep='last')
    
    # 4. Erstat NaN med tom streng for JSON-kompatibilitet
    full_df = full_df.fillna('')
    
    # --- RENSNING AF DUBLETTER SLUT ---

    combined_data = full_df.to_dict(orient='records')

    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(combined_data, f, ensure_ascii=False, indent=2)
    
    print(f"Succes! Gemte {len(combined_data)} unikke rækker i data.json")

if __name__ == "__main__":
    download_and_convert()
import pandas as pd
import requests
import json
from io import BytesIO
from datetime import datetime

URL = "https://skat.dk/media/5bldctyv/marts-2026-abis-liste-2021-2026.xlsx"

def download_and_convert():
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(URL, headers=headers)
    excel_file = pd.ExcelFile(BytesIO(response.content))
    
    all_sheets_data = []
    for sheet_name in excel_file.sheet_names:
        if "forside" in sheet_name.lower(): continue
        df = excel_file.parse(sheet_name)
        all_sheets_data.append(df)

    full_df = pd.concat(all_sheets_data, ignore_index=True)
    
    # Rensning
    full_df = full_df.dropna(how='all').drop_duplicates()
    isin_col = next((c for c in full_df.columns if 'ISIN' in str(c).upper()), None)
    if isin_col:
        full_df = full_df.drop_duplicates(subset=[isin_col], keep='last')
    
    full_df = full_df.fillna('')

    # Ny struktur: Metadata + Data
    output = {
        "metadata": {
            "last_updated": datetime.now().strftime("%d. %B %Y kl. %H:%M"),
            "source_url": URL,
            "description": "Baseret på Skats liste over beviser og aktier i investeringsforeninger og selskaber (IFPA)."
        },
        "records": full_df.to_dict(orient='records')
    }

    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    print(f"Succes! Data opdateret {output['metadata']['last_updated']}")

if __name__ == "__main__":
    download_and_convert()
