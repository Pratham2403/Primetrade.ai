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
        "sparkline": "false" 
    }

    res = requests.get(url, params=params)

    if res.status_code == 200 : 
        return res.json()
    else:
        print(f"Error {res.status_code} while fetching data")
        return None
    
def analyze_data(data):
    top5Crypto = sorted(data, key=lambda x: x['market_cap'], reverse=True)[:5]
    totalPrice = sum([crypto['current_price'] for crypto in data])
    avgPrice = totalPrice / len(data) if data else 0
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

    
    df = pd.DataFrame(cryptoData)

    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer :  

        
        df.to_excel(writer, sheet_name="Live Crypto Data", index=False)

        worksheet = writer.sheets["Live Crypto Data"]

        
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter  
            for cell in col:
                try:
                    
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)  
            worksheet.column_dimensions[column].width = adjusted_width
        
        
        analysis_sheet = writer.book.create_sheet("Analysis")
        analysis_sheet.append(["Top 5 Cryptos by Market Cap"])
        for crypto in analysis['top5Crypto']:
            analysis_sheet.append([crypto['name'], crypto['symbol'], crypto['market_cap']])
            analysis_sheet.append(["Average Price of Top 50 Cryptos", analysis['avgPrice']])
        analysis_sheet.append(["Highest 24h % Change", analysis['highestChange']['name'], analysis['highestChange']['price_change_percentage_24h']])
        analysis_sheet.append(["Lowest 24h % Change", analysis['lowestChange']['name'], analysis['lowestChange']['price_change_percentage_24h']])

        
        for col in analysis_sheet.columns:
            max_length = 0
            column = col[0].column_letter  
            for cell in col:
                try:
                    
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)  
            analysis_sheet.column_dimensions[column].width = adjusted_width

    print(f"Data saved to {filename}")

def main():

    filename = "CryptoAnalysis.xlsx"

    while True : 
        data = fetchCrypto()

        if(data):
            
            analysis = analyze_data(data)
            writeToExcel(data, analysis, filename)

            print("Excel updated with live data and analysis")
        else:
            print("Error in fetching data")
        
        
        time.sleep(300)

if __name__ == "__main__":
    main()
            


