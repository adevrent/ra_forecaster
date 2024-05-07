import json
import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlwings as xw

def get_security_params(infolist):
    disclosureIndex = infolist[0]
    
    # SecurityCoupon Sheet
    
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
    for label in ["ISIN Kodu", "Döviz Cinsi", "İhraç Fiyatı", "Faiz Oranı - Yıllık Basit (%)", "Satışı Gerçekleştirilen Nominal Tutar", "Satışa Başlanma Tarihi", "Kupon Sayısı"]:
        param = None
        for row in soup.find_all("tr"):
            label_div = row.find("div", class_="bold font14")
            if label_div and label in label_div.text:
                param = row.find("div", class_="gwt-HTML control-label lineheight-32px").text.strip()
                paramdict[label] = param
                break
        if param == None:  # bad code, but works.
            paramdict[label] = param

    # Step 4: Extract tables
    tables = pd.read_html(str(soup))
    print("TABLES:", tables)

    # Step 5: Process and display tables
    df = tables[5]
    print("table 5:", df)
    

    # Create the DataFrame
    df = pd.DataFrame(data, columns=headers)
    
    new_header = df.iloc[0]  # save first row as column headers row
    df = df.iloc[1:, :]  # Take the data less the header row
    df.columns = new_header  # Set the header row as the df header
    df["ISIN_CODE"] = paramdict["ISIN Kodu"]
    df.set_index(df.iloc[:, -1], inplace=True)
    df["COUPON_DATE"] = pd.to_datetime(df["Ödeme Tarihi"], format="%d.%m.%Y")
    df["Ödeme Tutarı"] = df["Ödeme Tutarı"].apply(european_to_float)
    df["COUPON_RATE"] = df["Faiz Oranı - Dönemsel (%)"].apply(float) / 10000
    
    if paramdict["Faiz Oranı - Yıllık Basit (%)"] == None:
        coupon = european_to_float(df["Faiz Oranı - Yıllık Basit (%)"][0])
    else:
        coupon = european_to_float(paramdict["Faiz Oranı - Yıllık Basit (%)"])
    
    df = df.iloc[:, [-2, -1]]
    df_security_coupon = df.dropna()
    
    # Security Sheet
    # Basis
    if paramdict["Döviz Cinsi"] == "TRY":
        basis = "ACTL365"
    elif paramdict["Döviz Cinsi"] == "EUR":
        basis = "EU30360"
    else:
        basis = "US30360"
        
    fis_dict = {
        "ISIN_CODE":paramdict["ISIN Kodu"],
        "INSTRUMENT_TYPE":None,  # instrument type is assigned below
        "MATURITY_DATE":df["COUPON_DATE"][-1],
        "CURRENCY":paramdict["Döviz Cinsi"],
        "FREQUENCY":int(paramdict["Kupon Sayısı"]),
        "COUPON":coupon,
        "SPREAD":0,  # hard coded, fix later.
        "ISSUER_CODE":infolist[2],  # second element of input list is issuer code, by definition from parse_disclosures() function.
        "ISSUE_INDEX":0,  # hard coded, fix later.
        "ISSUE_DATE":pd.to_datetime(paramdict["Satışa Başlanma Tarihi"], format="%d.%m.%Y"),
        "DAY_YEAR_BASIS":basis,
        "ISSUE_PRICE":european_to_float(paramdict["İhraç Fiyatı"]) * 100,
        "totalIssuedAmount":european_to_float(paramdict["Satışı Gerçekleştirilen Nominal Tutar"]),
        "securityType":None,  # hard coded, fix later.
        "fundUser":None  # hard coded, fix later.
        }
    
    # Instrument Type
    if infolist[1]:
        if fis_dict["FREQUENCY"] == 0:
            fis_dict["INSTRUMENT_TYPE"] = "CORP_SUKUK_DISCOUNTED"
        else:
            if fis_dict["SPREAD"] == 0:
                fis_dict["INSTRUMENT_TYPE"] = "CORP_SUKUK_FIXED_COUPON"
            else:
                fis_dict["INSTRUMENT_TYPE"] = "CORP_SUKUK_FLOATING"
    else:
        if fis_dict["CURRENCY"] == "TRY":
            if fis_dict["FREQUENCY"] == 0:
                fis_dict["INSTRUMENT_TYPE"] = "CORP_DISCOUNTED"
            else:
                if fis_dict["ISSUE_INDEX"] == 0:
                    if fis_dict["SPREAD"] == 0:  # there are actually 0 spread FRN's, fix this later.
                        fis_dict["INSTRUMENT_TYPE"] = "CORP_FIXED_COUPON"
                    else:
                        fis_dict["INSTRUMENT_TYPE"] = "CORP_FLOATING"
                else:
                    fis_dict["INSTRUMENT_TYPE"] = "TÜFEX"
        else:
            fis_dict["INSTRUMENT_TYPE"] = "EUROBOND"
        
    # Generate security dataframe from fis_dict
    df_security = pd.Series(fis_dict)
    
    return df_security, df_security_coupon

def parse_disclosures():
    paramdict = {}
    # Retrieve the data from the API
    url = "https://www.kap.org.tr/tr/api/disclosures"
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the JSON response
        disclosures = json.loads(response.text)
        
        # Extract 'basic' and 'detail' data into separate lists of dictionaries
        basic_data = [disclosure['basic'] for disclosure in disclosures]
        # detail_data = [disclosure['detail'] for disclosure in disclosures]
    else:
        print("Failed to retrieve data from the API")
        
    for disc in basic_data:
        if disc["title"] in ["Pay Dışında Sermaye Piyasası Aracı İşlemlerine İlişkin Bildirim (Faiz İçeren)", "Pay Dışında Sermaye Piyasası Aracı İşlemlerine İlişkin Bildirim (Faizsiz)"]:
            paramdict[disc["disclosureIndex"]] = [disc["stockCodes"]]
    
    print("type:", type(basic_data))
    print("len:", len(basic_data))
    
    return paramdict

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
        elif disc["title"] == "Pay Dışında Sermaye Piyasası Aracı İşlemlerine İlişkin Bildirim (Faizsiz)":
            paramlist.append([disc["disclosureIndex"], True, disc["stockCodes"].split(',')[0].strip()])  # [sukuk flag, issuer code]
    
    # print("type:", type(basic_data))
    # print("len:", len(basic_data))
    
    return paramlist