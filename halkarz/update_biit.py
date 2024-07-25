import requests
from bs4 import BeautifulSoup
import json

# Define headers with a User-Agent
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def update_bist_ilk_islem_tarihi(json_file):
    # Read the existing JSON file
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    for company_code, company_info in data.items():
        company_data, company_link = company_info

        if company_data.get("Bist İlk İşlem Tarihi") == "Hazırlanıyor...":
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

                    for row in rows:
                        cells = row.find_all('td')
                        if len(cells) == 2:
                            label = cells[0].text.strip()
                            value = cells[1].text.strip()

                            if "Bist İlk İşlem Tarihi" in label:
                                # Check if the value is a date
                                if value != "Hazırlanıyor...":
                                    # Update the "Bist İlk İşlem Tarihi"
                                    data[company_code][0]["Bist İlk İşlem Tarihi"] = value
                                    print(f"Updated 'Bist İlk İşlem Tarihi' for {company_code} to {value}")
                                else:
                                    print(f"Date not ready yet for {company_code}.")
                                break

    # Write the updated data back to the JSON file
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# Update the 'Bist İlk İşlem Tarihi' in the 'halkarz.json' file
update_bist_ilk_islem_tarihi('halkarz.json')