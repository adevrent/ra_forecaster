import requests
from bs4 import BeautifulSoup
import pandas as pd

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

                        # Create a Pandas Series and append it to the dataframe
                        series = pd.Series(company_info, name=bist_kodu)
                        df = pd.concat([df, pd.DataFrame(series).T])

    return df

# Test
print(get_halkarz_info())