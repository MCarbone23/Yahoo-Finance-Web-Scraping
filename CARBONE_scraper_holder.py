import json
import csv
import sys
from typing import Any, Dict 
import requests
from bs4 import BeautifulSoup
import pandas as pd

def get_data(ticker_symbol: Any) -> Dict[str, Any]: 
    print('Getting holder data of ', ticker_symbol)

    # Set user agent to avoid detection as a scraper
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:122.0) Gecko/20100101 Firefox/122.0'}

    url = f'https://finance.yahoo.com/quote/{ticker_symbol}/holders'  

    # Make a request to the URL
    r = requests.get(url, headers=headers) 
    soup = BeautifulSoup(r.text, 'html.parser')
    # description_element = soup.find('section', {'class': 'quote-sub-section'})
    # description_paragraph = description_element.find('p', {'class': 'Mt(15px) Lh(1.6)'}) if description_element else None
    # description = description_paragraph.text.strip() if description_paragraph else ''
        # Extract stock holder data from the HTML

    holder = {
        'ticker_symbol': ticker_symbol,
        '%_shares_all_insiders': soup.find_all('td', class_ = 'Py(10px) Va(m) Fw(600) W(15%)')[0].text.strip(),
        '%_shares_institutions': soup.find_all('td', class_ = 'Py(10px) Va(m) Fw(600) W(15%)')[1].text.strip(),
        '%_float_institutions': soup.find_all('td', class_ = 'Py(10px) Va(m) Fw(600) W(15%)')[2].text.strip(),
        '#_institutions': soup.find_all('td', class_ = 'Py(10px) Va(m) Fw(600) W(15%)')[3].text.strip(),
        'holder1': soup.find_all('td', class_ = 'Ta(start) Pend(10px)')[0].text.strip(),
        'holder2': soup.find_all('td', class_ = 'Ta(start) Pend(10px)')[1].text.strip(),
        'holder3': soup.find_all('td', class_ = 'Ta(start) Pend(10px)')[2].text.strip(),
        'holder4': soup.find_all('td', class_ = 'Ta(start) Pend(10px)')[3].text.strip(),
        'holder5': soup.find_all('td', class_ = 'Ta(start) Pend(10px)')[4].text.strip()
        }
    return holder

# Check if ticker symbols are provided as command line arguments
if len(sys.argv) < 2:
    print("Usage: python script.py <ticker_symbol1> <ticker_symbol2> ...") 
    sys.exit(1)

# Extract ticker symbols from command line arguments
ticker_symbols = sys.argv[1:]

# Get holder data for each ticker symbol
holderdata = [get_data(symbol) for symbol in ticker_symbols]

# Writing holder data to a JSON file
with open('CARBONE_stock_holder_data.json', 'w', encoding='utf-8') as f: 
    json.dump(holderdata, f)

# Writing holder data to a CSV file with aligned values
CSV_FILE_PATH = 'CARBONE_stock_holder_data.csv'
with open(CSV_FILE_PATH, 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = holderdata[0].keys()
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore') 
    writer.writeheader()
    writer.writerows(holderdata)

# Writing holder data to an Excel file
EXCEL_FILE_PATH = 'CARBONE_stock_holder_data.xlsx'
df = pd.DataFrame(holderdata) 
df.to_excel(EXCEL_FILE_PATH, index=False)

print('Done!')