import pandas as pd
from openpyxl import load_workbook
wb = load_workbook(filename=r'C:\Users\bowss\Downloads\U.S.DividendChampions.xlsx', read_only=True)
ws = wb['All CCC']
data_rows = []
for row in ws['B7':'B732']:
    data_cols = []
    for cell in row:
        data_cols.append(cell.value)
    data_rows.append(data_cols)
    df = pd.DataFrame(data_rows)
print (df)
#import yfinance as yf
from yahoo_fin import stock_info as si
print(si.get_live_price("AAPL"))
from yahoo_fin import options
chain = options.get_options_chain("AAPL")
print(chain)