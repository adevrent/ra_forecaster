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
    
    # if data empty, try with "^" symbol as ticker might be a fund
    if len(data) == 0:
        data = yf.download(f"^{ticker}", start_date, end_date)
    
    if len(data) == 0:
        # if data empty again, try futures
        commodity = ticker.removeprefix("F_").lower().split()[0]
        tickerdict = {
            "cotton":"CT=F",
            "sugar":"SB=F",
            "coffee":"KC=F", 
            "crude":"CL=F",
            "woil":"CL=F", 
            "brent":"BZ=F", 
            "cattle":"LE=F",
            "feeder":"LE=F",
            "cocoa":"CC=F", 
            "copper":"HG=F",
            "corn":"ZC=F",
            "ngas":"NG=F",
            "natural":"NG=F",
            "oat":"ZO=F",
            "platn":"PL=F",
            "plat":"PL=F",
            "soybean":"ZS=F",
            "soy":"ZS=F",
            "wheat":"KE=F",
            "kcbt":"KE=F",
            "silver":"SI=F",
            "gold":"GC=F",
            "gc":"GC=F"
            }
        
        for key in tickerdict.keys():
            if key in commodity:
                print(f"commodity, key pair found -> {commodity}, {key}" )
                data = yf.download(tickerdict[key], start_date, end_date)
                break
                
    # Write dataframe to Excel
    wb.sheets[0].range("A1").value = data
    
# Run code
ticker = "F_Sugar May24"
yf_xw(ticker)