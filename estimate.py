import pandas as pd
import requests
from bs4 import BeautifulSoup

# Read the Excel file from desktop
file_path = r"C:\Users\serendipaty\OneDrive\Desktop\year.xlsx"
df = pd.read_excel(file_path)

# Function to fetch 52-week range from Yahoo Finance
def fetch_52_week_range(yahoo_code):
    try:
        url = f"https://finance.yahoo.com/quote/{yahoo_code}?p={yahoo_code}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Look for 52-week range in the page content
        range_element = soup.find("td", {"data-test": "FIFTY_TWO_WK_RANGE-value"})
        if range_element:
            return range_element.text
        return None
    except Exception as e:
        print(f"Error fetching for {yahoo_code}: {e}")
        return None

# Update the dataframe with 52-week ranges
df['52-Week Range'] = df['Yahoo Code'].apply(fetch_52_week_range)

# Save the updated dataframe back to Excel
df.to_excel(file_path, index=False)
