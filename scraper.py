'''
Philip Massouh || massouh.3@osu.edu

This script gets a list of the top 500 compaies in the S&P from wikipedia
then iterates through them getting their market cap and ESG number then
writes it to an excel sheet.
'''

import time
from bs4 import BeautifulSoup
import pandas as pd
import requests
import xlwt
from xlwt import Workbook

# Header for Yahoo finance requests
headers = { 
    'User-Agent'      : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36', 
    'Accept'          : 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 
    'Accept-Language' : 'en-US,en;q=0.5',
    'DNT'             : '1',
    'Connection'      : 'close'
}

def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    if iteration == total: 
        print()

def getMarketCap(company):
    html = BeautifulSoup(requests.get(f'https://finance.yahoo.com/quote/{company}?p={company}', headers=headers).text, 'html.parser')
    market_cap = html.find('td', {'Ta(end) Fw(600) Lh(14px)'})
    if market_cap:
        return market_cap.text
    return float('nan')

def getESGOverall(company):
    html = BeautifulSoup(requests.get(f'http://finance.yahoo.com/quote/{company}/sustainability?p={company}', headers=headers).text, 'html.parser')
    scraped = html.find('div', {'Fz(36px) Fw(600) D(ib) Mend(5px)'})
    # Sometimes we dont have an esg score
    if scraped:
        return scraped.text
    return float('nan')

# (names, tickers)
print(f'--> Fetching tickers for S&P Top 500 Companies')
start = time.time()
page = pd.read_html(requests.get('https://en.wikipedia.org/wiki/List_of_S%26P_500_companies').text)[0]
companies = list(zip(page['Security'], page['Symbol']))
end = time.time()
print(f'--> {round(end-start, 3)}s | Data collected.')

# Make a new excel workbook and sheet
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

# Instantiate the column headers and excel sheet
l = len(companies)
print(f'--> Note: expect this process to be slow because we are circumventing the yahoo finance API.\n\n')

df = pd.DataFrame(columns = "Name Market_Cap Tot_ESG".split(), index = companies[1])

printProgressBar(0, l, prefix = 'Progress:', suffix = 'Complete', length = 50)
for i, company in enumerate(companies):

    df.loc[company[1]] = [company[0], getMarketCap(company[1]), getESGOverall(company[1])]
    printProgressBar(i + 1, l, prefix = 'Progress:', suffix = 'Complete', length = 50)

df.to_excel("esg.xlsx")
df.to_csv("esg.csv")
