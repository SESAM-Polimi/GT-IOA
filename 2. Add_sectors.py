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

paths = 'Paths.xlsx'

#%% Parse aggregated database from excel
world = mario.parse_from_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\Aggregated_SUT.xlsx", table='SUT', mode="flows")

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
        'Dysprosium'
        ],
    'activities': [
        'Production of photovoltaic plants',
        'Production of photovoltaic modules',
        'Production of mono-Si and poli-Si cells',
        'Extraction of raw silicon',
        'Production of onshore wind plants',
        'Production of DFIG generators',
        'Production of offshore wind plants',
        'Production of PMG generators',
        'Extraction of neodymium',
        'Extraction of dysprosium'
        ]
    }

#%% Getting excel templates to add new commodities
path_commodities = f"{pd.read_excel(paths, index_col=[0]).loc['Add Sectors',user]}\\new_commodities.xlsx"
path_activities  = f"{pd.read_excel(paths, index_col=[0]).loc['Add Sectors',user]}\\new_activities.xlsx"
# world.get_add_sectors_excel(new_sectors = new_sectors['commodities'],regions= world.get_index('Region'),path=path_commodities, item='Commodity')
# world.get_add_sectors_excel(new_sectors = new_sectors['activities'],regions= world.get_index('Region'),path=path_activities, item='Activity')

#%% Adding new commodities and activities
world.add_sectors(io=path_commodities, new_sectors= new_sectors['commodities'], regions= world.get_index('Region'), item= 'Commodity', inplace=True)
world.add_sectors(io=path_activities,  new_sectors= new_sectors['activities'],  regions= world.get_index('Region'), item= 'Activity',  inplace=True)

#%% Aggregated database with new sectors to excel
world.to_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\Aggregated & new sectors SUT.xlsx")
