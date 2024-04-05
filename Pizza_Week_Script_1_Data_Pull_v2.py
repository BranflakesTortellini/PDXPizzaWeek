import requests
from bs4 import BeautifulSoup
import pandas as pd
import sqlite3
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog
import os

# Define database path
db_path = 'sublinks.db'

# Delete the existing database file to start fresh
if os.path.exists(db_path):
    os.remove(db_path)


# Initialize the database connection
conn = sqlite3.connect('sublinks.db')
c = conn.cursor()

# Create the table for storing sublinks and their processed status
c.execute('''
CREATE TABLE IF NOT EXISTS sublinks (
    url TEXT PRIMARY KEY,
    processed INTEGER DEFAULT 0,
    verified INTEGER DEFAULT 0
)''')
conn.commit()

def insert_sublinks_to_db(links):
    """Insert new sublinks into the database."""
    c.executemany('INSERT OR IGNORE INTO sublinks (url) VALUES (?)', [(link,) for link in links])
    conn.commit()

def get_unprocessed_links():
    """Retrieve a list of unprocessed sublinks from the database."""
    c.execute('SELECT url FROM sublinks WHERE processed = 0')
    return [row[0] for row in c.fetchall()]

def mark_link_as_processed(url):
    """Mark a sublink as processed in the database."""
    c.execute('UPDATE sublinks SET processed = 1 WHERE url = ?', (url,))
    conn.commit()

def get_subpage_links(main_page_url):
    """Fetch subpage links from the main page and store them in the database."""
    response = requests.get(main_page_url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        links = [a['href'] for a in soup.select('h3 a')]
        insert_sublinks_to_db(links)
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

def verify_all_links_processed():
    """Ensure all links have been processed and retry if necessary."""
    unprocessed_links = get_unprocessed_links()
    retries = 2  # Number of times to retry processing unprocessed links
    for _ in range(retries):
        if not unprocessed_links:
            break  # Exit if there are no unprocessed links
        for url in unprocessed_links:
            data.append(fetch_subpage_content(url))
            mark_link_as_processed(url)
        unprocessed_links = get_unprocessed_links()

def main(main_page_url):
    # Fetch subpage links and store them in the database
    get_subpage_links(main_page_url)

    # Process each subpage link
    data = []
    for url in tqdm(get_unprocessed_links(), desc='Fetching subpages'):
        data.append(fetch_subpage_content(url))
        mark_link_as_processed(url)
    
    # Verify and retry unprocessed links
    verify_all_links_processed()

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
    main_page_url = 'https://everout.com/portland/events/the-portland-mercurys-pizza-week-2024/e170026/'
    main(main_page_url)

    # Close the database connection when done
    conn.close()
