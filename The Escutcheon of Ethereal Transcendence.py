import yfinance as yf
from openpyxl import load_workbook
wb = load_workbook(r'C:\Users\bowss\Downloads\U.S.DividendChampions.xlsx', read_only=True)
wb.active = 4
for column in wb.active['B7':'B20']:
    for cell in column:
        print(round(yf.Ticker(cell.value).history(period='1d')['Close'][0], 2))
        try:
            print(yf.Ticker(cell.value).option_chain('2021-11-19'))
        except Exception as e:
            print(cell.value, e)
