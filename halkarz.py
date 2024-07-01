import requests
from bs4 import BeautifulSoup
import xlwings as xw

# Define headers with a User-Agent
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

# Fetch the main page HTML with headers
main_page_url = "https://halkarz.com/"
response = requests.get(main_page_url, headers=headers)
if response.status_code == 200:
    print("Successfully fetched the main page.")
else:
    print(f"Failed to fetch the main page. Status code: {response.status_code}")

main_page_html = response.text

# Parse the main page HTML with BeautifulSoup
soup = BeautifulSoup(main_page_html, 'html.parser')

# Find all company entries
company_entries = soup.find_all('article', class_='index-list')
print(f"Found {len(company_entries)} company entries.")
print("-"*40)

# Iterate over each company entry
for entry in company_entries:
    new_badge = entry.find('div', class_='il-new')
    if new_badge:
        # Extract the link to the company page
        company_link_tag = entry.find('a', href=True)
        if company_link_tag:
            company_link = company_link_tag['href']
            print("-"*40)
            print(f"Company Link: {company_link}")

            # Fetch the company page HTML with headers
            company_page_response = requests.get(company_link, headers=headers)
            if company_page_response.status_code == 200:
                print(f"Successfully fetched the company page: {company_link}")
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
                            company_info[label] = value

                    # Print the extracted information
                    company_title_tag = entry.find('h3', class_='il-halka-arz-sirket')
                    if company_title_tag:
                        company_title = company_title_tag.text.strip()
                        print(f"Company: {company_title}")

                    for label, value in company_info.items():
                        print(f"{label} {value}")
                else:
                    print("No detailed company information found.")
            else:
                print(f"Failed to fetch the company page: {company_link}. Status code: {company_page_response.status_code}")
        else:
            print("Company link not found.")
    else:
        print("New badge not found.")