import json
import csv
import sys
from typing import Any, Dict 
import requests
from bs4 import BeautifulSoup
import pandas as pd

def get_data(ticker_symbol: Any) -> Dict[str, Any]: 
    print('Getting profile data of ', ticker_symbol)

    # Set user agent to avoid detection as a scraper
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:122.0) Gecko/20100101 Firefox/122.0'}

    url = f'https://finance.yahoo.com/quote/{ticker_symbol}/profile' 

    # Make a request to the URL
    r = requests.get(url, headers=headers) 
    soup = BeautifulSoup(r.text, 'html.parser')
    # description_element = soup.find('section', {'class': 'quote-sub-section'})
    # description_paragraph = description_element.find('p', {'class': 'Mt(15px) Lh(1.6)'}) if description_element else None
    # description = description_paragraph.text.strip() if description_paragraph else ''
        # Extract stock profile data from the HTML

    profile = {
        'stock_name': soup.find('div', {'class':'D(ib) Mt(-5px) Maw(38%)--tab768 Maw(38%) Mend(10px) Ov(h) smartphone_Maw(85%) smartphone_Mend(0px)'}).find_all('div')[0].text.strip(),
        'address': soup.find('p', {'class':'D(ib) W(47.727%) Pend(40px)'}).text.strip(),
        'key_executive1': soup.find_all('td', class_ = 'Ta(start)')[0].text.strip(),
        'key_executive2': soup.find_all('td', class_ = 'Ta(start)')[2].text.strip(),
        'key_executive3': soup.find_all('td', class_ = 'Ta(start)')[4].text.strip(),
        'key_executive4': soup.find_all('td', class_ = 'Ta(start)')[6].text.strip(),
        'key_executive5': soup.find_all('td', class_ = 'Ta(start)')[8].text.strip(),
        'description': soup.find('p', {'class':'Mt(15px) Lh(1.6)'}).text.strip(),
        'governance_data': soup.find('p', {'class':'Fz(s)'}).text.strip()
    }
    return profile

# Check if ticker symbols are provided as command line arguments
if len(sys.argv) < 2:
    print("Usage: python script.py <ticker_symbol1> <ticker_symbol2> ...") 
    sys.exit(1)

# Extract ticker symbols from command line arguments
ticker_symbols = sys.argv[1:]

# Get profile data for each ticker symbol
profiledata = [get_data(symbol) for symbol in ticker_symbols]

# Writing profile data to a JSON file
with open('CARBONE_stock_profile_data.json', 'w', encoding='utf-8') as f: 
    json.dump(profiledata, f)

# Writing profile data to a CSV file with aligned values
CSV_FILE_PATH = 'CARBONE_stock_profile_data.csv'
with open(CSV_FILE_PATH, 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = profiledata[0].keys()
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore') 
    writer.writeheader()
    writer.writerows(profiledata)

# Writing profile data to an Excel file
EXCEL_FILE_PATH = 'CARBONE_stock_profile_data.xlsx'
df = pd.DataFrame(profiledata) 
df.to_excel(EXCEL_FILE_PATH, index=False)

print('Done!')