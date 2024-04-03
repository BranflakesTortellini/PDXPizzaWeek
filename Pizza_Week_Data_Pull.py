import requests
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm

def get_subpage_links(main_page_url):
    """Fetch subpage links from the main page."""
    response = requests.get(main_page_url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        links = [a['href'] for a in soup.select('h3 a')]
        return links
    return []

def fetch_subpage_content(url):
    """Fetch the content from the meta og:description and the descriptions block."""
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    meta_description = soup.find('meta', property='og:description')
    meta_content = meta_description['content'] if meta_description else 'Not Found'

    description_block = soup.select_one('.description')
    description_content = description_block.get_text(strip=True) if description_block else 'Not Found'

    return {'URL': url, 'Meta Description': meta_content, 'Description Content': description_content}

def save_to_excel(data, filepath, engine='openpyxl'):
    """Save the data to an Excel file."""
    df = pd.DataFrame(data)
    # Ensure the dataframe is correctly interpreted as containing Unicode characters
    df = df.applymap(lambda x: x.encode('utf-8').decode('utf-8') if isinstance(x, str) else x)
    df.to_excel(filepath, index=False, engine=engine)



def main(main_page_url):
    subpage_links = get_subpage_links(main_page_url)
    print(f"Found {len(subpage_links)} subpages to process.")

    data = [fetch_subpage_content(url) for url in tqdm(subpage_links, desc='Fetching subpages')]
    
    root = tk.Tk()
    root.withdraw()  # Hide the Tkinter GUI
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel files (*.xlsx)", "*.xlsx"),
                                                       ("Excel 97-2003 files (*.xls)", "*.xls"),
                                                       ("All files", "*.*")])
    if filepath:
        if filepath.endswith('.xlsx'):
            save_to_excel(data, filepath, engine='openpyxl')
        elif filepath.endswith('.xls'):
            save_to_excel(data, filepath, engine='xlwt')
        print(f"Data saved to {filepath}")
    else:
        print("Save file dialog cancelled.")


if __name__ == "__main__":
    MAIN_PAGE_URL = 'https://everout.com/portland/events/the-portland-mercurys-pizza-week-2024/e170026/'
    main(MAIN_PAGE_URL)


 