#%% -*- coding: utf-8 -*-
"""
Created on Tue Apr 18 2023

@authors: 
    Lorenzo Rinaldi, Department of Energy, Politecnico di Milano
    Nicolò Golinucci, PhD, Department of Energy, Politecnico di Milano
    Emanuele Mainardi, Department of Energy, Politecnico di Milano
    Prof. Matteo Vincenzo Rocco, PhD, Department of Energy, Politecnico di Milano
    Prof. Emanuela Colombo, PhD, Department of Energy, Politecnico di Milano
"""


import mario
import pandas as pd

user = "LR"
sN = slice(None)
years = range(2011,2020)

paths = 'Paths.xlsx'

#%% Parse aggregated database from excel
world = {}
for year in years:
    world[year] = mario.parse_from_txt(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\a. Aggregated_SUT\\{year}\\flows", table='SUT', mode="flows")

#%% Define new commodities
new_sectors = {
    'commodities': [
        'Photovoltaic plants',
        'Photovoltaic modules',
        'Mono-Si and poli-Si cells',
        'Raw silicon',
        'Onshore wind plants',
        'DFIG generators',
        'Offshore wind plants',
        'PMG generators',
        'Neodymium',
        'Dysprosium',
        ],
    'activities': [
        'Production of photovoltaic plants',
        'Production of photovoltaic modules',
        'Production of mono-Si and poli-Si cells',
        'Production of onshore wind plants',
        'Production of DFIG generators',
        'Production of offshore wind plants',
        'Production of PMG generators'
        ]
    }

#%% Getting excel templates to add new commodities
path_commodities = f"{pd.read_excel(paths, index_col=[0]).loc['Add Sectors',user]}\\new_commodities.xlsx"
path_activities  = f"{pd.read_excel(paths, index_col=[0]).loc['Add Sectors',user]}\\new_activities.xlsx"
# world[year].get_add_sectors_excel(new_sectors = new_sectors['commodities'],regions= world.get_index('Region'),path=path_commodities, item='Commodity')
# world[year].get_add_sectors_excel(new_sectors = new_sectors['activities'],regions= world.get_index('Region'),path=path_activities, item='Activity')

#%% Adding new commodities and activities
for year in years:
    world[year].add_sectors(io=path_commodities, new_sectors= new_sectors['commodities'], regions= world[year].get_index('Region'), item= 'Commodity', inplace=True)
    world[year].add_sectors(io=path_activities,  new_sectors= new_sectors['activities'],  regions= world[year].get_index('Region'), item= 'Activity',  inplace=True)

#%% Aggregated database with new sectors to excel
for year in years:
    world[year].to_txt(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\b. Aggregated & new sectors SUT\\{year}", flows=False, coefficients=True)
