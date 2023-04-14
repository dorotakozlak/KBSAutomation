# -*- coding: utf-8 -*-
"""
Created on Thu Apr 13 22:03:43 2023


@author: 80300698
"""

import pandas as pd 
import xlsxwriter 
import openpyxl 

url = "D:\\Codes\\Cardex\\"
source_file = "Cardex_details_DTS.xlsm"
tab = "data"
split_data_file = "Cardex_details_DTS2.xlsx"
url_split_data = "D:\\Codes\\Cardex\\Customers\\"

df = pd.read_excel(url + source_file, tab)

#modification to the file 
df = df.iloc[2: ,1:]
df.columns = df.iloc[0]
df = df.iloc[1:]

writer = pd.ExcelWriter(url + source_file, engine = 'openpyxl', mode = 'a', if_sheet_exists= 'replace')

for customer in df['Customer'].unique():
    newDf = df[df['Customer'] == customer]
    newDf.to_excel(writer, sheet_name = customer, index = False)
writer.save()


customers = {
    'LIDL': ['Customer', 'Technology', 'Flavour'],
    'BIM': ['Customer', 'Brand']}
#iterate for every key and every value in dictionary 
for customer, columns in customers.items(): 
    df_client = df[columns]
    df_client.to_excel(url_split_data + f'{customer}.xlsx', sheet_name = customer, index = False)

