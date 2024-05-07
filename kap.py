import json
import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlwings as xw

def european_to_float(value):
    """
    Convert a European formatted string to a float.
    E.g., "30.416.600" -> 30416600.00
        "30.416.600,99" -> 30416600.99
    """
    if isinstance(value, str):
        value = value.replace('.', '').replace(',', '.')
    try:
        return float(value)
    except ValueError:
        return None

def get_security_params(infolist):
    disclosureIndex = infolist[0]
    sukuk_flag = infolist[1]

    # Step 1: Fetch HTML content
    url = f"https://www.kap.org.tr/tr/Bildirim/{disclosureIndex}"
    response = requests.get(url)

    if response.status_code == 200:
        html_content = response.text
    else:
        raise Exception(f"Failed to fetch webpage: {response.status_code}")

    # Step 2: Parse the HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # Step 3: Extract several params
    paramdict = {}
    for label in ["ISIN Kodu", "Vade Tarihi", "Döviz Cinsi", "İhraç Fiyatı", "Faiz Oranı - Yıllık Basit (%)", "Satışı Gerçekleştirilen Nominal Tutar", "Satışa Başlanma Tarihi", "Kupon Sayısı", "Ek Getiri (%)", "Kupon Ödeme Sıklığı"]:
        param = None
        for row in soup.find_all("tr"):
            label_div = row.find("div", class_="bold font14")
            if label_div and label in label_div.text:
                param = row.find("div", class_="gwt-HTML control-label lineheight-32px").text.strip()
                paramdict[label] = param
                break
        if param is None:  # If the label isn't found, set None
            paramdict[label] = param
    
    # Security Coupon Sheet
    
    # If 0 or 1 coupon payment, there is no cash flow table
    if len(soup.find_all("table")) < 10:
        coupon = european_to_float(paramdict["Faiz Oranı - Yıllık Basit (%)"])
    
        fis_dict={}
        if int(paramdict["Kupon Sayısı"]) == 0:
            inst_type = "CORP_DISCOUNTED"
            coupon = 0
        else:
            inst_type = "CORP_FIXED_COUPON"
        
        discdict = {"ISIN_CODE":paramdict["ISIN Kodu"], "COUPON_DATE":pd.to_datetime(paramdict["Vade Tarihi"], format="%d.%m.%Y"), "COUPON_RATE":european_to_float(paramdict["Faiz Oranı - Yıllık Basit (%)"])}
        df_security_coupon = pd.Series(discdict)
        df_security_coupon = df_security_coupon.to_frame().T
        df_security_coupon.set_index("ISIN_CODE", inplace=True)
    # If more than 1 coupon payments:
    else:
        table = soup.find_all("table")[5]
        rows = table.find_all("tr")

        # Extract the table headers
        headers = [header.text.strip() for header in rows[0].find_all("td")]
        
        # Extract the table data
        data = []
        for row in rows[1:]:
            data.append([cell.text.strip() for cell in row.find_all("td")])

        # Create the DataFrame
        df = pd.DataFrame(data, columns=headers)
        # df["Ödeme Tutarı"] = df["Ödeme Tutarı"].apply(european_to_float)
        df["Faiz Oranı - Dönemsel (%)"] = df["Faiz Oranı - Dönemsel (%)"].apply(european_to_float)
        df["COUPON_DATE"] = pd.to_datetime(df["Ödeme Tarihi"], format="%d.%m.%Y")
        df["ISIN_CODE"] = paramdict["ISIN Kodu"]
        
        # Set coupon rate depending if FRN or not
        if paramdict["Faiz Oranı - Yıllık Basit (%)"] == None:
            coupon = european_to_float(df["Faiz Oranı - Yıllık Basit (%)"][0])
        else:
            coupon = european_to_float(paramdict["Faiz Oranı - Yıllık Basit (%)"])

        df_security_coupon = df.loc[:, ["ISIN_CODE", "COUPON_DATE", "Faiz Oranı - Dönemsel (%)"]].dropna()
        df_security_coupon.columns = ["ISIN_CODE", "COUPON_DATE", "COUPON_RATE"]
        df_security_coupon.set_index("ISIN_CODE", inplace=True)


    # Security Sheet
    # Basis
    if paramdict["Döviz Cinsi"] == "TRY":
        basis = "ACTL365"
    elif paramdict["Döviz Cinsi"] == "EUR":
        basis = "EU30360"
    else:
        basis = "US30360"
        
    # Spread
    if paramdict["Ek Getiri (%)"] in ["-", None]:
        spread = 0
    else:
        spread = european_to_float(paramdict["Ek Getiri (%)"])
        
    # Coupon Frequency
    if paramdict["Kupon Ödeme Sıklığı"] == None:
        frequency  = 0
    elif paramdict["Kupon Ödeme Sıklığı"].lower() == "tek kupon":
        frequency = 1
    elif paramdict["Kupon Ödeme Sıklığı"].lower() == "yıllık":
        frequency = 1
    elif paramdict["Kupon Ödeme Sıklığı"].lower() == "6 ayda bir":
        frequency = 2
    elif paramdict["Kupon Ödeme Sıklığı"].lower() == "3 ayda bir":
        frequency = 4
    elif paramdict["Kupon Ödeme Sıklığı"].lower() == "aylık":
        frequency = 12
        

    # Params dict
    fis_dict = {
        "ISIN_CODE": paramdict["ISIN Kodu"],
        "INSTRUMENT_TYPE": None,  # instrument type is assigned below
        "MATURITY_DATE": pd.to_datetime(paramdict["Vade Tarihi"], format="%d.%m.%Y"),
        "CURRENCY": paramdict["Döviz Cinsi"],
        "FREQUENCY": int(frequency),
        "COUPON": coupon,
        "SPREAD": spread,
        "ISSUER_CODE": infolist[2],  # third element of input list is issuer code
        "ISSUE_INDEX": 0,  # hard-coded, fix later
        "ISSUE_DATE": pd.to_datetime(paramdict["Satışa Başlanma Tarihi"], format="%d.%m.%Y"),
        "DAY_YEAR_BASIS": basis,
        "ISSUE_PRICE": european_to_float(paramdict["İhraç Fiyatı"]) * 100,
        "totalIssuedAmount": european_to_float(paramdict["Satışı Gerçekleştirilen Nominal Tutar"]),
        "securityType": None,  # hard-coded, fix later
        "fundUser": None  # hard-coded, fix later
    }

    # Instrument Type
    if len(soup.find_all("table")) < 10:
        fis_dict["INSTRUMENT_TYPE"] = inst_type
    else:
        if sukuk_flag:
            if fis_dict["FREQUENCY"] == 0:
                fis_dict["INSTRUMENT_TYPE"] = "CORP_SUKUK_DISCOUNTED"
            else:
                if paramdict["Ek Getiri (%)"] == None:
                    fis_dict["INSTRUMENT_TYPE"] = "CORP_SUKUK_FIXED_COUPON"
                else:
                    fis_dict["INSTRUMENT_TYPE"] = "CORP_SUKUK_FLOATING"
        else:
            if fis_dict["CURRENCY"] == "TRY":
                if fis_dict["FREQUENCY"] == 0:
                    fis_dict["INSTRUMENT_TYPE"] = "CORP_DISCOUNTED"
                else:
                    if fis_dict["ISSUE_INDEX"] == 0:
                        if paramdict["Ek Getiri (%)"] == None:  # there are actually 0 spread FRN's, fix this later
                            fis_dict["INSTRUMENT_TYPE"] = "CORP_FIXED_COUPON"
                        else:
                            fis_dict["INSTRUMENT_TYPE"] = "CORP_FLOATING"
                    else:
                        fis_dict["INSTRUMENT_TYPE"] = "TÜFEX"
            else:
                fis_dict["INSTRUMENT_TYPE"] = "EUROBOND"

    # Generate security dataframe from fis_dict
    df_security = pd.Series(fis_dict)
    df_security = df_security.to_frame().T
    df_security.set_index("ISIN_CODE", inplace=True)

    return df_security, df_security_coupon

def parse_disclosures():
    paramlist = []
    # Retrieve the data from the API
    url = "https://www.kap.org.tr/tr/api/disclosures"
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the JSON response
        disclosures = json.loads(response.text)
        
        # Extract 'basic' data into list of dictionaries
        basic_data = [disclosure['basic'] for disclosure in disclosures]
        # detail_data = [disclosure['detail'] for disclosure in disclosures]
    else:
        print("Failed to retrieve data from the API")
        
    for disc in basic_data:
        if disc["title"] == "Pay Dışında Sermaye Piyasası Aracı İşlemlerine İlişkin Bildirim (Faiz İçeren)":
            paramlist.append([disc["disclosureIndex"], False, disc["stockCodes"].split(',')[0].strip()])  # [sukuk flag, issuer code]
        # elif disc["title"] == "Pay Dışında Sermaye Piyasası Aracı İşlemlerine İlişkin Bildirim (Faizsiz)":
        #     paramlist.append([disc["disclosureIndex"], True, disc["stockCodes"].split(',')[0].strip()])  # [sukuk flag, issuer code]
    
    # print("type:", type(basic_data))
    # print("len:", len(basic_data))
    
    return paramlist

def merge_disclosures():
    flag = True
    for disclist in parse_disclosures():
        if flag:
            df_security, df_security_coupon = get_security_params(disclist)
            flag = False
        else:
            df_security = pd.concat([df_security, get_security_params(disclist)[0]])
            df_security_coupon = pd.concat([df_security_coupon, get_security_params(disclist)[1]])
    
    return df_security, df_security_coupon

def kap_xw():
    # Create a new Excel workbook
    wb = xw.Book()
    sht2 = wb.sheets.add("SecurityCoupon")
    sht1 = wb.sheets.add("Security")
    
    sht1.range("A1").value = merge_disclosures()[0]
    sht2.range("A1").value = merge_disclosures()[1]
    
    default_sheet = wb.sheets[2]
    default_sheet.delete()
    
    # Make the first row bold
    sht1.range("A1").expand('right').api.Font.Bold = True
    sht2.range("A1").expand('right').api.Font.Bold = True
    
    # Autofit columns to expand cells to fit their contents
    sht1.autofit()
    sht2.autofit()

# Run code    
kap_xw()