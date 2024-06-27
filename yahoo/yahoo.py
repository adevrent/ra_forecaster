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

    # Reindex to include all business days between the start and end of the data
    all_days = pd.date_range(start=data.index.min(), end=data.index.max(), freq='B')
    data = data.reindex(all_days, method='ffill')
    
    # Drop rows that are in Turkish holidays
    data = data[~data.index.isin(hdays['holiday_date'])]
    
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
    
    # Try BIST stocks first
    data = yf.download((ticker.split("_")[0] + ".IS"), start_date, end_date)
    currency = "TRY"
    label = "IMKB"
    data["ASSET_NAME"] = ticker.upper()
    
    # Try US stocks and ETFs
    if len(data) == 0:
        data = yf.download(ticker.split("_")[0], start_date, end_date)
        currency = "USD"
        label = "SP500"
        
        if "_us" in ticker:
            data["ASSET_NAME"] = f"{ticker.upper()}"
        else:
            data["ASSET_NAME"] = f"{ticker.upper()}_US"
    
    # if data empty, try with "^" symbol as ticker might be a US index
    if len(data) == 0:
        data = yf.download(f"^{ticker}", start_date, end_date)
        currency = "USD"
        label = "FOR_IND"
        data["ASSET_NAME"] = ticker.upper()
    
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
                currency = "USD"
                label = "COMDTY_MARKET"
                data["ASSET_NAME"] = key.upper()
                break
            
    # if data empty again, try with ".CBT" suffix
    if len(data) == 0:
        data = yf.download(f"{ticker}.CBT", start_date, end_date)
        currency = "USD"
        label = "COMDTY_MARKET"
        data["ASSET_NAME"] = ticker.upper()
        
    # if data empty again, try with ".CMX" suffix
    if len(data) == 0:
        data = yf.download(f"{ticker}.CMX", start_date, end_date)
        currency = "USD"
        label = "COMDTY_MARKET"
        data["ASSET_NAME"] = ticker.upper()
        
     # if data empty again, raise ValueError
    if len(data) == 0:
        raise ValueError("ticker not found!")
    
    # Adjust for Turkish business days:
    data = adjust_for_turkish_business_days(data, holidays_filepath)
    
    # Format adjustments
    data["CURRENCY_CODE"] = currency
    data["MARKET_NAME"] = label
    data["PRICE_BID"] = data["Adj Close"]
    data["PRICE_ASK"] = data["Adj Close"]
    data["PRICE_AVERAGE"] = data["Adj Close"]
    data["DATA_SOURCE"] = "BLOOMBERG"
    data["DATA_TYPE"] = "EOD"
    data["RECORD_TIME"] = None
    
    # Drop/Rename the columns
    data.drop(columns=["Close", "Volume"], inplace=True)
    data.rename(columns={
        "Open": "PRICE_OPEN",
        "High": "PRICE_HIGH",
        "Low": "PRICE_LOW",
        "Adj Close": "PRICE_CLOSE"
    }, inplace=True)
    
    # Re-order columns
    data = data[[
                "ASSET_NAME", "CURRENCY_CODE", "MARKET_NAME",
                "PRICE_OPEN", "PRICE_HIGH", "PRICE_LOW", "PRICE_CLOSE",
                "DATA_SOURCE", "DATA_TYPE", "RECORD_TIME"
            ]]
    
    # Rename the index column to "RECORD_DATE"
    data.rename_axis("RECORD_DATE", inplace=True)
                
    # Write dataframe to Excel
    wb.sheets[0].range("A1").value = data
    print("Excel export successful.")

    # Make the first row bold
    wb.sheets[0].range("1:1").api.Font.Bold = True
    
    # Resize columns for readability
    wb.sheets[0].autofit()
    
    # Set sheet name based on the label
    if label == "SP500":
        sheet_name = "EQUITY"
    elif label == "FOR_IND":
        sheet_name = "INDEX"
    else:
        sheet_name = "COMMODITY"
    wb.sheets[0].name = sheet_name
    
    # Save the Excel workbook
    if output_path != None:
        wb.save(output_path + f"{ticker}" + f"_{date.today()}" + ".xlsx")
        wb.close()
        
        
# Run code
holidays_filepath = r"C:\Users\adevr\ra_forecaster\yahoo\riskfree_holiday.xlsx"
output_path = r"C:\\Users\\adevr\\OneDrive\\Belgeler\\Riskactive Portföy\\Historical data\\"

ticker = "TSLA_US"

yf_xw(ticker, output_path=None, holidays_filepath=holidays_filepath)