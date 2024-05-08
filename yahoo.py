import xlwings as xw
import yfinance as yf
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta

# Set the default end date to today's date as default
end = date.today()

# Calculate the default start date, which is one year before the end date as default
start = end - relativedelta(years=1)

def yf_xw(ticker, start_date=None, end_date=None, output_path=None):
    # Create a new Excel workbook
    wb = xw.Book()
    
    if start_date is None:
        start_date = pd.to_datetime(start, format="%d.%m.%Y")
    
    if end_date is None:
        end_date = pd.to_datetime(end, format="%d.%m.%Y")
    
    data = yf.download(ticker, start_date, end_date)
    # if data empty, try with "^" symbol as ticket might be a fund
    if len(data) == 0:
        data = yf.download(f"^{ticker}", start_date, end_date)

    # Write dataframe to Excel
    wb.sheets[0].range("A1").value = data
    
# Run code
ticker = "SPY"
yf_xw(ticker)