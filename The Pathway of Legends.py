import pandas as pd; import yfinance as yf
from openpyxl import load_workbook
Rocco_hash = []
for column in load_workbook(r'C:\Users\bowss\Downloads\Stocks.xlsx', read_only=True)['Symbols']['A1':'A127']:
    for cell in column:
        try:
            Rocco_hash.append(yf.Ticker(cell.value).option_chain('2022-03-25').calls)
        except ValueError or TypeError:
            pass
pd.concat(Rocco_hash).to_excel('ContractsS.xlsx')
