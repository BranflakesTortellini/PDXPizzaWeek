import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openai import OpenAI
import re

# Initialize the OpenAI client for the local LLM
client = OpenAI(base_url="http://localhost:1234/v1", api_key="lm-studio")

def query_llm_for_addresses(entry):
    completion = client.chat.completions.create(
        model="delli/mistral-7b-address-validator-merged-gguf/mistral-7b-address-validator-merged.Q4_K_M.gguf",
        #model="bartowski/Starling-LM-7B-beta-GGUF/Starling-LM-7B-beta-IQ4_XS.gguf",
        messages=[
            {
                "role": "system",
                "content": (
                    "Your task is to analyze text entries detailing one or more locations. Each entry may include the location's address, "
    "operational hours, days open, and possibly other notes. It is crucial to enumerate each location distinctly and provide "
    "all related details in an organized manner. Begin with 'Address 1:', 'Address 2:', etc., for each location, followed by "
    "'Hours:', and 'Days:'. If there are additional notes, list them under 'Notes:'. Use a new line for each piece of information, "
    "and separate locations with a blank line. Ensure the format is consistent for ease of reading and direct use in database entries. "
    "Note that sometimes the text entries will list a location such as a district, neighborhood, town, etc. before the address using a format like Location: or Location -."
    "For example something that says PSU: 540 SW College St is the full address you should share even though it includes a location: with the location being PSU."
    "Some entries will simply have a street address at the beginning. Something like 410 NW 21st Ave has an address of 410 NW 21st Ave."
    "Please include this location, if provided, in the Address fields that you populate."
    "There will always be an Address 1, there will sometimes be an Address 2."
    "The format should be as follows:\n\n"
    "- Address 1: [First address details]\n"
    "- Hours: [Operational hours for Address 1], including variations by day if applicable\n"
    "- Days: [Days of operation for Address 1]\n"
    "- Notes: [Any additional information for Address 1]\n\n"
    "- Address 2: [Second address details]\n"
    "- Hours: [Operational hours for Address 2], including variations by day if applicable\n"
    "- Days: [Days of operation for Address 2]\n"
    "- Notes: [Any additional information for Address 2]\n\n"
    "Continue this pattern for all locations mentioned within the entry. This structured approach will help in accurately capturing and listing the details for each location, especially when the operating hours vary or additional notes are provided."
                )
            },
            {"role": "user", "content": entry}
        ],
        temperature=0.3,
        stream=False 
    )
    response = completion.choices[0].message if completion.choices else None
    return response.content if hasattr(response, 'content') else response




def parse_llm_response(content):
    # Replace <br> tags with \n to standardize newlines for splitting
    content = re.sub(r'<br>', '\n', content)
    parsed_locations = []
    # Split based on 'Address ' to find each location info
    locations = content.split('Address ')

    address1 = None  # Variable to hold Address 1

    for i, loc in enumerate(locations[1:], start=1):  # Skip the first split as it's empty or not an address
        # Check if this block is 'Address 2: None'
        if i == 2 and 'None' in loc:
            continue  # Skip this block
        
        # Parse address, hours, and days
        address_match = re.search(r'\d+: (.*?)\nHours:', loc)
        hours_match = re.search(r'Hours: (.*?)\nDays:', loc)
        days_match = re.search(r'Days: (.*)', loc)

        if address_match and hours_match and days_match:
            # Extract the matches and strip to remove any whitespace
            address = address_match.group(1).strip()
            hours = hours_match.group(1).strip()
            days = days_match.group(1).strip()

            if i == 1:  # If this is Address 1, store it in case Address 2 is 'None'
                address1 = address

            parsed_locations.append({
                'Address': address,
                'hours_and_days': f"{hours} / {days}"
            })
        elif i == 1:  # If Address 1 parsing failed, mark it for a retry
            address1 = "RETRY"

    if address1 == "RETRY":
        # Return a special flag to trigger a retry for this entry
        return [{'Address': "RETRY", 'hours_and_days': "RETRY"}]

    # If Address 1 was found but Address 2 was 'None', ensure Address 1's info is outputted
    if address1 and not parsed_locations:
        parsed_locations.append({
            'Address': address1,
            'hours_and_days': "Hours/Days: Not specified"
        })

    return parsed_locations


def process_dataframe(df):
    new_data = []
    for index, row in df.iterrows():
        retry_count = 0
        while retry_count < 3:  # Allow up to two retries
            response = query_llm_for_addresses(row['Locations and Times'])
            if response:
                address_entries = parse_llm_response(response)
                valid_entries = [entry for entry in address_entries if entry['Address'] != "RETRY"]
                if valid_entries:
                    break
                retry_count += 1
            else:
                break  # Break if there is no response from LLM

        if not valid_entries and retry_count == 3:
            # If after two retries the correct info is still not obtained, flag as unavailable
            valid_entries = [{'Address': "LLM response unavailable", 'hours_and_days': "LLM response unavailable"}]

        for entry in valid_entries:
            new_row = row.copy()
            new_row['Address'] = entry['Address']
            new_row['Hours and Days of operation'] = entry['hours_and_days']
            new_data.append(new_row)

    new_columns = df.columns.tolist() + ['Address', 'Hours and Days of operation']
    return pd.DataFrame(new_data, columns=new_columns)


def reorder_and_drop_columns(df, drop_LandT=True):
    # Define the new order of the columns
    new_order = [
        'Pizza Name', 'Vendor Name', 'Serving Style', 'Type',
        'Vegan Option','Vegetarian Option', 'Meat Option', 'Gluten-Free',
        'Gluten-Free Substitute Available', 'Address',
        'Hours and Days of operation', 'Minors Allowed',
        'Takeout Available', 'Delivery Available',
        'Purchase Limit', 'More Info Link'
    ]

    # Check if 'Locations and Times' should be dropped
    if drop_LandT and 'Locations and Times' in df.columns:
        df = df.drop(columns=['Locations and Times'])

    # Reorder the DataFrame according to 'new_order'
    df = df[new_order]

    return df



def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window

    file_path = filedialog.askopenfilename(
        title="Select the file with locations and times",
        filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
    )

    if file_path:
        df = pd.read_csv(file_path) if file_path.endswith('.csv') else pd.read_excel(file_path)
        processed_df = process_dataframe(df)

        # Reorder and drop columns as needed
        final_df = reorder_and_drop_columns(processed_df)

        output_file_path = filedialog.asksaveasfilename(
            title="Save the processed file",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            defaultextension=".xlsx" if file_path.endswith('.xlsx') else ".csv"
        )

        if output_file_path:
            final_df.to_csv(output_file_path, index=False) if output_file_path.endswith('.csv') else final_df.to_excel(output_file_path, index=False, engine='openpyxl')
            print(f"Processed file saved to {output_file_path}")
    else:
        print("No file selected.")

if __name__ == "__main__":
    main()
