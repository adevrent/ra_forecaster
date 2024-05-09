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
        start_date = start
    else:
        start_date = pd.to_datetime(start_date, format="%d.%m.%Y")
    
    if end_date is None:
        end_date = end
    else:
        end_date = pd.to_datetime(end_date, format="%d.%m.%Y")
    
    ticker = ticker.removeprefix("F_").lower().split()[0]
    
    # Try stocks
    data = yf.download(ticker.split("_")[0], start_date, end_date)
    
    # if data empty, try with "^" symbol as ticker might be a fund
    if len(data) == 0:
        data = yf.download(f"^{ticker}", start_date, end_date)
    
    if len(data) == 0:
        # if data empty again, try futures
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
            if key in ticker:
                print(f"commodity, key pair found -> {ticker}, {key}" )
                data = yf.download(tickerdict[key], start_date, end_date)
                break
                
    # Write dataframe to Excel
    wb.sheets[0].range("A1").value = data
    
    # Save the Excel workbook
    if output_path != None:
        wb.save(output_path + "YAHOO_" + f"{date.today()}" + ".xlsx")
        wb.close()
    
# Run code
ticker = "TSLA_US"
yf_xw(ticker)