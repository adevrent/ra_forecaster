import xlwings as xw
import yfinance as yf
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta

# Set the default end date to today's date as default
end = date.today()

# Calculate the default start date, which is one year before the end date as default
start = end - relativedelta(years=1)

def adjust_for_turkish_business_days(data, holidays_filepath):
    # Load the data from the first sheet
    hdays = pd.read_excel(holidays_filepath, sheet_name='Sheet1')

    # Correct the column name and convert it to pandas datetime objects
    hdays.rename(columns={'# holiday_date': 'holiday_date'}, inplace=True)
    hdays['holiday_date'] = pd.to_datetime(hdays['holiday_date'])

    # Set the index to date
    data.index = pd.to_datetime(data.index)

    # Drop rows that are in Turkish holidays
    data = data[~data.index.isin(hdays['holiday_date'])]

    # Reindex to include all business days between the start and end of the data
    all_days = pd.date_range(start=data.index.min(), end=data.index.max(), freq='B')
    data = data.reindex(all_days, method='ffill')
    
    return data

def yf_xw(ticker, start_date=None, end_date=None, output_path=None, holidays_filepath=None):
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
    
    # Debug
    print("normalized ticker =", ticker)
    
    # Try stocks
    data = yf.download(ticker.split("_")[0], start_date, end_date)
    
    # if data empty, try with "^" symbol as ticker might be a fund
    if len(data) == 0:
        data = yf.download(f"^{ticker}", start_date, end_date)
    
    # if data empty again, try futures
    if len(data) == 0:
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
            
    # if data empty again, try with ".CBT" suffix
    if len(data) == 0:
        data = yf.download(f"{ticker}.CBT", start_date, end_date)
        
    # if data empty again, try with ".CMX" suffix
    if len(data) == 0:
        data = yf.download(f"{ticker}.CMX", start_date, end_date)
        
     # if data empty again, raise ValueError
    if len(data) == 0:
        raise ValueError("ticker not found!")
    
    # Adjust for Turkish business days:
    data = adjust_for_turkish_business_days(data, holidays_filepath)
                
    # Write dataframe to Excel
    wb.sheets[0].range("A1").value = data
    
    # Save the Excel workbook
    if output_path != None:
        wb.save(output_path + f"{ticker}" + f"_{date.today()}" + ".xlsx")
        wb.close()
        
# Run code
holidays_filepath = r"C:\Users\adevr\ra_forecaster\yahoo\riskfree_holiday.xlsx"
output_path = "C:\\Users\\adevr\\OneDrive\\Belgeler\\Riskactive Portföy\\Historical data\\"
ticker = "TSLA"
yf_xw(ticker, output_path=output_path, holidays_filepath=holidays_filepath)