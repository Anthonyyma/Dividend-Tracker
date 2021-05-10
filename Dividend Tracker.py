import gspread
import yfinance as yf
from forex_python.converter import CurrencyRates

Stocks = ['BPY', 'ENB', 'BNS', 'ZWC', 'REI']

BPY = yf.Ticker("BPY-UN.TO")
ENB = yf.Ticker("ENB.TO")
BNS = yf.Ticker("BNS.TO")
ZWC = yf.Ticker("ZWC.TO")
REI = yf.Ticker("REI-UN.TO")

BPY_D = BPY.dividends[-1]
ENB_D = ENB.dividends[-1]
BNS_D = BNS.dividends[-1]
ZWC_D = ZWC.dividends[-1]
REI_D = REI.dividends[-1]

# from google_auth_oauthlib import ServiceAccountCredentials
from google_auth_oauthlib.flow import InstalledAppFlow


scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

gc = gspread.service_account(filename='creds.json')
sh = gc.open_by_key("key") # or by sheet name: gc.open("TestList")
sheet = sh.sheet1


def setDividends():
    for i in range(2, len(Stocks) + 2):
        if sheet.cell(i, 14).value == 'M':
            Freq = 12
        else:
            Freq = 4

        found = False
        while not found:
            for j in Stocks:
                Ticker = sheet.cell(i, 2).value
                if j in Ticker:
                    Stock = j
                    found = True
                    break

        div = Stock + '_D*{}'.format(Freq)
        if Stock == 'BPY':
            div = int(eval(div)) * CurrencyRates().get_rate('USD', 'CAD')
            sheet.update_cell(i, 12, '$' + str(div))
        else:
            sheet.update_cell(i, 12, '$' + str(eval(div)))

        print(str(i - 1) + '/' + str(len(Stocks)) + ': ' + Stock)

def countRows():
    Count = False
    i = 2
    while not Count:
        if sheet.cell(i, 1).value == 'Total':
            row_count = i - 2
            Count = True
        else:
            i = i + 1
    return row_count

setDividends()