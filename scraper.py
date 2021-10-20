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

# Progress bar
# https://stackoverflow.com/questions/3173320/text-progress-bar-in-the-console
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    if iteration == total: 
        print()

# Header for Yahoo finance requests
headers = { 
    'User-Agent'      : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36', 
    'Accept'          : 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 
    'Accept-Language' : 'en-US,en;q=0.5',
    'DNT'             : '1',
    'Connection'      : 'close'
}

# Change this if you are so inclined
filename = 'out'

# Get list of (names, tickers)
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
print(f'--> Pulling and saving {l} companies to {filename}.xls:')
print(f'--> Note: expect this process to be slow because we are circumventing the yahoo finance API.\n\n')
sheet1.write(0, 0, 'Name')
sheet1.write(0, 1, 'Ticker')
sheet1.write(0, 2, 'Market Cap')
sheet1.write(0, 3, 'Total ESG')

# Use a progress bar so you have something to stare at
printProgressBar(0, l, prefix = 'Progress:', suffix = 'Complete', length = 50)

# Run through the list of tickers and get the market cap and esg score for them
for i, company in enumerate(companies):

    # Name
    sheet1.write(i+1, 0, company[0])

    # Ticker
    sheet1.write(i+1, 1, company[1])

    # Market cap
    html = BeautifulSoup(requests.get('http://finance.yahoo.com/quote/'+company[1]+'?p='+company[1], headers=headers).text, 'html.parser')
    market_cap = html.find('span', {'Trsdu(0.3s)'})
    # We should have always have a market cap but why not include this check on a script that takes 20 min to run
    if market_cap:
        sheet1.write(i+1, 2, market_cap.text)
    else:
        sheet1.write(i+1, 2, 'Not found')

    # ESG score
    html = BeautifulSoup(requests.get('http://finance.yahoo.com/quote/'+company[1]+'/sustainability?p='+company[1], headers=headers).text, 'html.parser')
    scraped = html.find('div', {'Fz(36px) Fw(600) D(ib) Mend(5px)'})
    # Sometimes we dont have an esg score
    if scraped:
        sheet1.write(i+1, 3, scraped.text)
    else:
        sheet1.write(i+1, 3, 'Not found')

    # Update the progress bar
    printProgressBar(i + 1, l, prefix = 'Progress:', suffix = 'Complete', length = 50)

# Save and export the excel doc
wb.save(f'{filename}.xls')
