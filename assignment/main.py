import requests
import pandas as pd
from openpyxl.utils import get_column_letter
import time

def fetchCrypto():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "inr",
        "order": "market_cap_desc",
        "per_page": "50",
        "page": "1",
        "sparkline": "false" #Sparkline means the graph of the coin
    }

    res = requests.get(url, params=params)

    if res.status_code == 200 : 
        return res.json()
    else:
        print(f"Error {res.status_code} while fetching data")
        return None
    
def analyze_data(data):
    #Top 5 Crypto by market Cap
    top5Crypto = sorted(data, key=lambda x: x['market_cap'], reverse=True)[:5]

    #Calculate The average price of Top 50 Crypto
    totalPrice = sum([crypto['current_price'] for crypto in data])
    avgPrice = totalPrice / len(data) if data else 0

    #find the crypto with the highest & Lowest 24 hr change
    highestChange = max(data, key=lambda x: x['price_change_percentage_24h'])
    lowestChange = min(data, key=lambda x: x['price_change_percentage_24h'])

    return {
        "top5Crypto": top5Crypto,
        "avgPrice": avgPrice,
        "highestChange": highestChange,
        "lowestChange": lowestChange
    }

def writeToExcel(data, analysis, filename="crypto_data.xlsx"):

    cryptoData = []

    for x in data:
        cryptoData.append({
            "Cryptocurrency Name" : x['name'],
            "Symbol" : x['symbol'],
            "Current Price (INR)" : x['current_price'],
            "Market Cap (INR)" : x['market_cap'],
            "24hr Trading Volume (INR)" : x['total_volume'],
            "Price Change (24hr)" : x['price_change_percentage_24h']
        })

    #Convertinf the data into DataFrame
    df = pd.DataFrame(cryptoData)

    #Writing in excel with all the analysis needed to be Done
    with pd.ExcelWriter(filename, engine='openpyxl') as writer :  

        #writing Live Data
        df.to_excel(writer, sheet_name="Live Crypto Data", index=False)

        worksheet = writer.sheets["Live Crypto Data"]

        #Adjusting Column Width
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
        
        #Writing Analysis to new Sheet
        analysis_sheet = writer.book.create_sheet("Analysis")
        analysis_sheet.append(["Top 5 Cryptos by Market Cap"])
        for crypto in analysis['top5Crypto']:
            analysis_sheet.append([crypto['name'], crypto['symbol'], crypto['market_cap']])
            analysis_sheet.append(["Average Price of Top 50 Cryptos", analysis['avgPrice']])
        analysis_sheet.append(["Highest 24h % Change", analysis['highestChange']['name'], analysis['highestChange']['price_change_percentage_24h']])
        analysis_sheet.append(["Lowest 24h % Change", analysis['lowestChange']['name'], analysis['lowestChange']['price_change_percentage_24h']])

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

def main():

    filename = "CryptoAnalysis.xlsx"

    while True : 
        data = fetchCrypto()

        if(data):
            #First we get the analysis of the data that is needed
            analysis = analyze_data(data)
            writeToExcel(data, analysis, filename)

            print("Excel updated with live data and analysis")
        else:
            print("Error in fetching data")
        
        # Wait for 5 minutes before next update (300 seconds)
        time.sleep(300)

if __name__ == "__main__":
    main()
            


