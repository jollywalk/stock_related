import pandas as pd
import requests
from bs4 import BeautifulSoup

# Read the Excel file from desktop
file_path = r"C:\Users\serendipaty\OneDrive\Desktop\year.xlsx"
df = pd.read_excel(file_path)

# Function to fetch 52-week range and calculate percentage return
def fetch_52_week_data(yahoo_code):
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
            range_text = range_element.text
            low, high = range_text.split(' - ')
            low = float(low.replace(',', ''))
            high = float(high.replace(',', ''))
            percentage_return = ((high - low) / low) * 100
            return range_text, percentage_return
        return None, None
    except Exception as e:
        print(f"Error fetching for {yahoo_code}: {e}")
        return None, None

# Update the dataframe with 52-week ranges and percentage returns
df[['52-Week Range', 'Percentage Return']] = df['Yahoo Code'].apply(lambda x: pd.Series(fetch_52_week_data(x)))

# Save the updated dataframe back to Excel
df.to_excel(file_path, index=False)
