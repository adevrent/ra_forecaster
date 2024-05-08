import xlwings as xw
import yfinance as yf
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta

# Set the default end date to today's date as default
end_date = date.today()

# Calculate the default start date, which is one year before the end date as default
start_date = end_date - relativedelta(years=1)

def yf_xw(ticker, start_date=None, end_date=None, isFund=False, output_path=None):
    # Create a new Excel workbook
    wb = xw.Book()
    
    if start_date is None:
        start_date = pd.to_datetime(start_date, format="%d.%m.%Y")
    
    if end_date is None:
        end_date = pd.to_datetime(end_date, format="%d.%m.%Y")
    
    if isFund:
        data = yf.download(f"^{ticker}", start_date, end_date)
    else:
        data = yf.download(ticker, start_date, end_date)
    
    # Write dataframe to Excel
    wb.sheets[0].range("A1").value = data
    
# Run code
yf_xw("TSLA")