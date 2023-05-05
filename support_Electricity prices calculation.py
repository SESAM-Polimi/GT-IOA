# -*- coding: utf-8 -*-
"""
Created on Thu May  4 09:44:07 2023

@author: loren
"""

import mario
import pandas as pd

user = "LR"
sN = slice(None)

paths = 'Paths.xlsx'

years = range(2011,2020)
ref = 'EXIOHSUT'

ref_and_unitconv = {
    'IEA': 1/3600*1e9,
    'EXIOHSUT': 1/3600*1e9,
    }

world = {}

#%% Parse aggregated  with new sectors database from excel
for year in years:
    if ref == 'EXIOHSUT' and year!=2011:
        print("year must be 2011")
    else:
        world[year] = mario.parse_from_txt(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\c. Baseline\\{year}\\coefficients", table='SUT', mode="coefficients")
    
#%% Extract electricity production in EUR
electricity_comms = world[years[0]].search("Commodity", "Electricity")[1:-1]
ee_prod_EUR = pd.DataFrame()

for year, db in world.items():
    X_ee = db.X.loc[(slice(None),slice(None),electricity_comms)].groupby(level=[0], axis=0).sum() 
    X_ee = X_ee.sort_index()
    X_ee.columns = [year]
    ee_prod_EUR = pd.concat([ee_prod_EUR, X_ee], axis=1)

#%% Read electricity production data
ee_prod_kWh = pd.read_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Electricity',user]}", sheet_name=f"{ref}_Electricity production", index_col=[0])*ref_and_unitconv[ref]
try:
    ee_prod_kWh = ee_prod_kWh.to_frame()
except:
    pass
ee_prod_kWh = ee_prod_kWh.sort_index()
ee_prod_kWh.index.names = ee_prod_EUR.index.names

#%% Calculate ee_prices
ee_prod_EUR.columns = ee_prod_kWh.columns
ee_prices = ee_prod_EUR/ee_prod_kWh
ee_prices.index.names = ['€/kWh']

#%% Save electricity prices
writer = pd.ExcelWriter(f"{pd.read_excel(paths, index_col=[0]).loc['Electricity',user]}", engine='openpyxl', mode='a', if_sheet_exists='replace')
ee_prices.to_excel(writer, sheet_name=f"{ref}_Electricity prices")
writer.close()

