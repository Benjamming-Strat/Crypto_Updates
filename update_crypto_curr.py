#Update daily cryptocurrencies
from requests import Request, Session
from requests.exceptions import ConnectionError, Timeout, TooManyRedirects
import json
import pickle
import pandas as pd
import numpy as np
import openpyxl
import os
import xlrd 


url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest'
parameters = {
  'start':'1',
  'limit':'5000',
  'convert':'USD'
}
headers = {
  'Accepts': 'application/json',
  'X-CMC_PRO_API_KEY': 'INSERT_UR_KEY',
}

session = Session()
session.headers.update(headers)

try:
  response = session.get(url, params=parameters)
  data = json.loads(response.text)
  print(data)
except (ConnectionError, Timeout, TooManyRedirects) as e:
  print(e)

#For testing i saved the Dictionary into an pickle object to save my API-Key-Credits
# with open('crypto_data.pkl', 'wb') as crypto_dict:
#     pickle.dump(data, crypto_dict)
# with open('crypto_data.pkl', 'rb') as crypto_dict:
#     data = pickle.load(crypto_dict)

coin_list_name = []
coin_list_price = []

for i in range(0,25):
    crypto_name = data["data"][i]["name"]
    crypto_price = data["data"][i]["quote"]["USD"]["price"]

    coin_list_name.append(crypto_name)
    coin_list_price.append(crypto_price)


#building the appropriate Dictioanry to read into Dataframe
coin_dict = {
    "Cryptocurrency": coin_list_name,
    "Price in USD": coin_list_price
            }

#building Dataframe            
coin_df = pd.DataFrame.from_dict(coin_dict)

#writing df to Excel
dashboard_file_exists = os.path.exists("Dashboard_Crypto.xlsx") #make sure xl_file and python file are in same directory



if  dashboard_file_exists==True:
    coin_excel_df = pd.read_excel("Dashboard_Crypto.xlsx", engine="openpyxl", skiprows=3, usecols=[4,5,6,7], date_parser="Einkaufsdatum")
    coin_excel_df["Price in USD"] = coin_df["Price in USD"]
    print(coin_excel_df)


else:
    coin_df.to_excel("Dashboard_Crypto.xlsx", startcol=3, startrow=3)

#Saving new file with upgoing counter
file = os.listdir(r"C:\Users\bennk\Documents\Programmierung\Crypto_pro")
counter = 0
for i in file:
    if str(i).startswith("Dash"):
        counter += 1

filename = "Dashboard_Crypto_"+str(counter)+".xlsx"
coin_excel_df.to_excel(filename)
print(filename+"Counter: "+str(counter))
