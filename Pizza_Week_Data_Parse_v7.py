import pandas as pd
import tkinter as tk
from tkinter import filedialog
import re
import html

def clean_text(text):
    # First, use html.unescape to convert HTML entities to their corresponding characters
    text = html.unescape(text)
    # Then, replace any additional specific entities if necessary
    return text.replace('&nbsp;', ' ').strip()

def extract_vendor(description):
    # Look for the pattern: 'Daily Availability Limit?' followed by any text, then the last 'at', then the vendor name, then 'in'
    # The regex '(?:.+\bat\b\s*)' finds the last 'at' because the non-capturing group '(?:...)' with '.+\b' repeatedly matches any 'at' followed by characters, until the last 'at' is reached
    match = re.search(r"Daily Availability Limit\?.+?(?:.+\bat\b\s*)(.*?)\s+in", description, re.DOTALL | re.IGNORECASE)
    return match.group(1).strip() if match else 'Unknown'


def extract_info_from_description(description, url):
    description = clean_text(description)

    # Patterns to extract the info
    patterns = {
        'name': r"What It's Called:\s*([^<>\n\r]+?)\s*(?:What's On It:|$)",
        'ingredients': r"What's On It:\s*([^<>\n\r]+)",
        'description': r"What They Say About It:\s*([^<>\n\r]+)",
        'location_time': r"Where and When to Get It:\s*([^<>\n\r]+)",
        'meat_or_veggie': r"Meat or Vegetarian\?\s*([^<>\n\r]+)",
        'vegetarian_substitute': r"Vegetarian Substitute\?\s*([^<>\n\r]+)",
        'vegan_substitute': r"Vegan\s*Substitute\?\s*([^<>\n\r]+)",
        'gluten_free': r"Gluten Free\?\s*([^<>\n\r]+)",
        'gluten_free_substitute': r"Gluten Free Substitute\?\s*([^<>\n\r]+)",
        'whole_pie_or_slice': r"Whole Pie or Slice\?\s*([^<>\n\r]+)",
        'minors_allowed': r"Allow Minors\?\s*([^<>\n\r]+)",
        'takeout_allowed': r"Allow Takeout\?\s*([^<>\n\r]+)",
        'delivery_allowed': r"Allow Delivery\?\s*([^<>\n\r]+)",
        'purchase_limit': r"Purchase Limit per Customer\?\s*([^<>\n\r]+)",
        'daily_availability_limit': r"Daily Availability Limit\?\s*([^<>\n\r]+)",
    }

    info = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, description, re.DOTALL | re.IGNORECASE)
        info[key] = match.group(1).strip() if match else 'Unknown'

    # Determine availability
    meat_options = ['Meat' in info['meat_or_veggie']]
    veg_options = ['Vegetarian' in info['meat_or_veggie'], info['vegetarian_substitute'] == 'Yes']
    vegan_options = ['Vegan' in info['meat_or_veggie'], info['vegan_substitute'] == 'Yes']

    # Set indicators
    info['meat_option'] = 'Yes' if any(meat_options) else 'No'
    info['vegetarian_option'] = 'Yes' if any(veg_options) else 'No'
    info['vegan_option'] = 'Yes' if any(vegan_options) else 'No'

    # Compile Type field
    types = []
    if 'Yes' in info['vegan_option']:
        types.append('Vegan')
    if 'Yes' in info['vegetarian_option']:
        types.append('Vegetarian')
    if 'Yes' in info['meat_option']:
        types.append('Meat')
    info['Type'] = ', '.join(types) if types else 'Unknown'

    # Additional extraction for vendor name
    vendor_name = extract_vendor(description)
    info['vendor'] = vendor_name if vendor_name != 'Unknown' else 'Unknown Vendor'

    info['link'] = url

    return info
 

def rename_and_reorder_columns(df):
    # Renaming the columns
    new_column_names = {
        'name': 'Pizza Name',
        'location_time': 'Locations and Times',
        'Type': 'Type',
        'vegetarian_option': 'Vegetarian Option',
        'vegan_option': 'Vegan Option',
        'meat_option': 'Meat Option',
        'gluten_free': 'Gluten-Free',
        'gluten_free_substitute': 'Gluten-Free Substitute Available',
        'whole_pie_or_slice': 'Serving Style',
        'minors_allowed': 'Minors Allowed',
        'takeout_allowed': 'Takeout Available',
        'delivery_allowed': 'Delivery Available',
        'purchase_limit': 'Purchase Limit',
        'daily_availability_limit': 'Daily Availability',
        'vendor': 'Vendor Name',
        'link': 'More Info Link',
    }
    df = df.rename(columns=new_column_names)

    # Specifying the desired order of the columns
    desired_order = [
        'Pizza Name',
        'Vendor Name',
        'Serving Style',
        'Type',
        'Vegan Option',
        'Vegetarian Option',
        'Meat Option',
        'Gluten-Free',
        'Gluten-Free Substitute Available',
        'Minors Allowed',
        'Takeout Available',
        'Delivery Available',
        'Purchase Limit',
        'Locations and Times',
        'More Info Link',
        # Add any additional columns in the order you want them to appear
    ]
    
    # Reordering the columns in the DataFrame
    df = df[desired_order]

    return df

def main():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()

    df_raw = pd.read_excel(file_path)

    processed_data = [extract_info_from_description(row['Meta Description'], row['URL']) for index, row in df_raw.iterrows()]

    df_extracted = pd.DataFrame(processed_data)
    df_cleaned = rename_and_reorder_columns(df_extracted)

    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files (*.xlsx)", "*.xlsx"), ("All files", "*.*")])
    if output_file_path:
        # Explicitly using openpyxl engine for better handling of Unicode.
        df_cleaned.to_excel(output_file_path, index=False, engine='openpyxl')
        print(f"Processed file saved to {output_file_path}")

if __name__ == "__main__":
    main()