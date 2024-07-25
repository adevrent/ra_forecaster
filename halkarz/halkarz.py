import requests
from bs4 import BeautifulSoup
import pandas as pd
import json
import os

def european_to_float(value):
    """
    Convert a European formatted string to a float.
    E.g., "30.416.600" -> 30416600.00
        "30.416.600,99" -> 30416600.99
    """
    if isinstance(value, str):
        # print("value =", value)
        value = value.split(" ")[0]  # splits the unit name
        # print("value after split =", value)
        value = value.replace('.', '').replace(',', '.')
        # print("value after replace =", value)
    try:
        return float(value)
    except:
        return None
    
def Pay_to_int(value):
    print("value =", value)
    value = value.split(" ")[0]
    print("value after split =", value)
    value = value.replace(",", "")
    print("value after replace =", value)
    
    return int(value)

def get_halkarz_info():
    # Define headers with a User-Agent
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    # Fetch the main page HTML with headers
    main_page_url = "https://halkarz.com/"
    response = requests.get(main_page_url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Failed to fetch the main page. Status code: {response.status_code}")

    main_page_html = response.text

    # Parse the main page HTML with BeautifulSoup
    soup = BeautifulSoup(main_page_html, 'html.parser')

    # Find all company entries
    company_entries = soup.find_all('article', class_='index-list')

    # Initialize empty dataframe and empty list for storing company links
    df = pd.DataFrame()
    company_links = {}
    
    # Iterate over each company entry
    for entry in company_entries:
        new_badge = entry.find('div', class_='il-new')
        if new_badge:
            # Extract the link to the company page
            company_link_tag = entry.find('a', href=True)
            if company_link_tag:
                company_link = company_link_tag['href']

                # Fetch the company page HTML with headers
                company_page_response = requests.get(company_link, headers=headers)
                if company_page_response.status_code == 200:
                    company_page_html = company_page_response.text

                    # Parse the company page HTML with BeautifulSoup
                    company_soup = BeautifulSoup(company_page_html, 'html.parser')

                    # Extract the desired information
                    company_article = company_soup.find('article', class_='single-page')
                    if company_article:
                        # Extract the rows of the table
                        rows = company_article.find_all('tr')

                        # Initialize a dictionary to hold the extracted information
                        company_info = {}

                        for row in rows:
                            cells = row.find_all('td')
                            if len(cells) == 2:
                                label = cells[0].text.strip()
                                value = cells[1].text.strip()
                                
                                if "arac" in label.lower():  # we do not need "Aracı Kurum" info
                                    continue
                                else:
                                    company_info[label[:-2]] = value

                        # Extract Bist Kodu for index
                        bist_kodu = company_info.get("Bist Kodu", "Unknown")
                        company_title_tag = entry.find('h3', class_='il-halka-arz-sirket')
                        company_title = company_title_tag.text.strip()
                        company_info["İhraççı"] = company_title
                        print(f"İhraççı: {company_title}")
                        
                        company_links[bist_kodu] = company_link
                        
                        # Create a Pandas Series and append it to the dataframe
                        series = pd.Series(company_info, name=bist_kodu)
                        df = pd.concat([df, pd.DataFrame(series).T])           

    # Re-formatting
    
    # Convert "Halka Arz Fiyatı/Aralığı" to float using the european_to_float function
    if "Halka Arz Fiyatı/Aralığı" in df.columns:
        df["Halka Arz Fiyatı/Aralığı"] = df["Halka Arz Fiyatı/Aralığı"].apply(european_to_float)
        
    # Convert "Pay" to float
    if "Pay" in df.columns:
        df["Pay"] = df["Pay"].apply(Pay_to_int)
        
    print("company links dict =", company_links)
    
    return df, company_links

def create_json(df, company_links):
    # Check if the JSON file already exists
    if os.path.exists('halkarz\halkarz.json'):
        with open('halkarz\halkarz.json', 'r', encoding='utf-8') as f:
            existing_data = json.load(f)
    else:
        existing_data = {}

    data_dict = df.to_dict(orient='index')
    final_dict = {key: [data_dict[key], company_links[key]] for key in data_dict}

    # Update existing data with new entries
    for key, value in final_dict.items():
        if key not in existing_data:
            existing_data[key] = value

    # Write updated data back to JSON file
    with open('halkarz\halkarz.json', 'w', encoding='utf-8') as f:
        json.dump(existing_data, f, ensure_ascii=False, indent=4)

# Test
df, company_links = get_halkarz_info()
create_json(df, company_links)

print(df)