import requests
import pandas as pd
from openpyxl.utils import get_column_letter
import time

# Function to fetch live cryptocurrency data using CoinGecko API
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        'vs_currency': 'usd',                    # Currency in USD
        'order': 'market_cap_desc',               # Order by market cap
        'per_page': 50,                           # Top 50 cryptocurrencies
        'page': 1,                                # Page 1
        'sparkline': 'false'                      # No sparkline data
    }
    response = requests.get(url, params=params)
    
    if response.status_code == 200:
        return response.json()                    # Return the data as JSON
    else:
        print(f"Error fetching data: {response.status_code}")
        return None

# Function to analyze the data and return analysis results
def analyze_data(data):
    # 1. Identify the top 5 cryptocurrencies by market cap
    top_5_cryptos = sorted(data, key=lambda x: x['market_cap'], reverse=True)[:5]

    # 2. Calculate the average price of the top 50 cryptocurrencies
    total_price = sum([crypto['current_price'] for crypto in data])
    avg_price = total_price / len(data) if data else 0

    # 3. Find the cryptocurrency with the highest and lowest 24-hour percentage price change
    highest_change = max(data, key=lambda x: x['price_change_percentage_24h'])
    lowest_change = min(data, key=lambda x: x['price_change_percentage_24h'])
    
    # Return all the analysis results
    return {
        'top_5_cryptos': top_5_cryptos,
        'avg_price': avg_price,
        'highest_change': highest_change,
        'lowest_change': lowest_change
    }

# Function to write data to Excel with formatted columns
def write_to_excel(data, analysis, filename="crypto_data.xlsx"):
    # Extract required fields for Excel sheet
    crypto_data = []
    for crypto in data:
        crypto_info = {
            'Cryptocurrency Name': crypto['name'],
            'Symbol': crypto['symbol'],
            'Current Price (USD)': crypto['current_price'],
            'Market Capitalization (USD)': crypto['market_cap'],
            '24h Trading Volume (USD)': crypto['total_volume'],
            'Price Change (24h %)': crypto['price_change_percentage_24h']
        }
        crypto_data.append(crypto_info)
    
    # Convert to DataFrame
    df = pd.DataFrame(crypto_data)
    
    # Write to Excel file
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Write live data
        df.to_excel(writer, sheet_name='Live Data', index=False)

        # Get the active sheet
        worksheet = writer.sheets['Live Data']

        # Adjust column width for 'Live Data' sheet
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name (A, B, C, etc.)
            for cell in col:
                try:
                    # Get the maximum length of any cell in the column
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)  # Add some padding (2 extra spaces)
            worksheet.column_dimensions[column].width = adjusted_width

        # Write analysis results to a new sheet
        analysis_sheet = writer.book.create_sheet("Analysis")
        analysis_sheet.append(["Top 5 Cryptos by Market Cap"])
        for crypto in analysis['top_5_cryptos']:
            analysis_sheet.append([crypto['name'], crypto['symbol'], crypto['market_cap']])
        
        analysis_sheet.append(["Average Price of Top 50 Cryptos", analysis['avg_price']])
        analysis_sheet.append(["Highest 24h % Change", analysis['highest_change']['name'], analysis['highest_change']['price_change_percentage_24h']])
        analysis_sheet.append(["Lowest 24h % Change", analysis['lowest_change']['name'], analysis['lowest_change']['price_change_percentage_24h']])

        # Adjust column width for the 'Analysis' sheet
        for col in analysis_sheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    # Get the maximum length of any cell in the column
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)  # Add some padding
            analysis_sheet.column_dimensions[column].width = adjusted_width

    print(f"Data saved to {filename}")

# Main function to run the script with continuous updates
def main():
    filename = "crypto_data.xlsx"
    
    while True:
        crypto_data = fetch_crypto_data()
        
        if crypto_data:
            # Perform analysis
            analysis_results = analyze_data(crypto_data)
            
            # Write to Excel
            write_to_excel(crypto_data, analysis_results, filename)
            print("Excel updated with live data and analysis")
        else:
            print("Error in fetching data")
        
        # Wait for 5 minutes before next update (300 seconds)
        time.sleep(300)

if __name__ == "__main__":
    main()
