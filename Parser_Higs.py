

#(1)
#Необходимо скачать:
#pip install requests
#pip install openpyxl
#pip install tldextract
#pip install beautifulsoup4
#pip install google-api-python-client
#(2)
#проложить дерикторию документа (71 строка)


import os
import requests
from urllib.parse import urlparse
from openpyxl import load_workbook, Workbook
import tldextract
from bs4 import BeautifulSoup

def extract_base_domain(url):
    ext = tldextract.extract(url)
    return f"{ext.subdomain}.{ext.domain}.{ext.suffix}/"

def google_search(query, api_key, cse_id, num_pages=1):
    base_url = "https://www.googleapis.com/customsearch/v1"
    results_per_page = 10

    unique_base_domains = set()
    results = []

    for page in range(1, num_pages + 1):
        start_index = (page - 1) * results_per_page + 1
        params = {
            'key': api_key,
            'cx': cse_id,
            'q': query,
            'start': start_index,
        }

        response = requests.get(base_url, params=params)
        data = response.json()

        for item in data.get('items', []):
            link = item.get('link', '')
            base_domain = extract_base_domain(link)
            if base_domain not in unique_base_domains:
                unique_base_domains.add(base_domain)
                results.append(link)

    return results

def save_to_excel(results, excel_file):
    wb = Workbook()
    ws = wb.active
    ws.append(['Search Results'])

    for result in results:
        ws.append([result])

    wb.save(filename=excel_file)
    print(f'Data saved to {excel_file}')

def clear_previous_data(excel_file):
    if os.path.exists(excel_file):
        os.remove(excel_file)

def main():
    api_key = "AIzaSyA_Sew14nmjVT1hC6wmbjUT89E140pfL3k"
    cse_id = "264b5ae7c46d44c43"
    directory_path = r"C:\"
    excel_file = os.path.join(directory_path, "XLS Worksheet.xlsx")

    # Очищает данные прошлого запроса при новом текущим
    clear_previous_data(excel_file)

    query = input("Введите ключевые слова для поиска: ")
    num_pages = int(input("Введите количество страниц Google для анализа: "))

    search_results = google_search(query, api_key, cse_id, num_pages)
    save_to_excel(search_results, excel_file)

if __name__ == "__main__":
    main()