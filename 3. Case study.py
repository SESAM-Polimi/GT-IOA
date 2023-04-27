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
import numpy as np

user = "LR"
sN = slice(None)

paths = 'Paths.xlsx'

#%% Parse aggregated  with new sectors database from excel
# world = mario.parse_from_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\b. Aggregated & new sectors SUT.xlsx", table='SUT', mode="coefficients")

#%% Getting shock templates
# world.get_shock_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\_template.xlsx")

#%% Implementing e fixing
# world.shock_calc(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\\Baseline - Carbon emissions fix.xlsx", v=True, e=True, z=True, scenario="shock 0")

# %% Baseline database to excel
# world.to_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\c. Baseline.xlsx", flows=False, coefficients=True)

#%% Parse baseline from excel
world = mario.parse_from_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\c. Baseline.xlsx", table='SUT', mode="coefficients")

#%% Linkages for baseline
world_iot = world.to_iot(method='B', inplace=False)
linkages = {
    'Baseline': world_iot.calc_linkages(multi_mode=False),
    }

#%% Implementing endogenization of capital
world.shock_calc(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\\Shock - Endogenization of capital.xlsx", z=True, v=True, scenario="Endogenization of capital")

#%%
path_aggr  = r"Aggregations\Aggregation_postprocess.xlsx"
world_aggr = world.aggregate(path_aggr, inplace=False, levels=["Activity","Commodity"])

f = {}
sat_accounts = [
    'Energy Carrier Supply: Total', 
    'CO2 - combustion - air', 
    'CH4 - combustion - air', 
    'N2O - combustion - air',
    ]

commodities = [
        'PV plants',
        'PV modules',
        'Si cells',
        'Onshore wind plants',
        'DFIG generators',
        'Offshore wind plants',
        'PMG generators',
        'Electricity by wind',
        'Electricity by solar photovoltaic'        
   ]


for a in sat_accounts:
    f[a] = {}
    if ":" in a:
        name = a.replace(':'," -")
    else:
        name = a
    writer = pd.ExcelWriter(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\{name}.xlsx", engine='openpyxl', mode='w')
    for s in world_aggr.scenarios:
        e = world_aggr.get_data(matrices=['e'], scenarios=[s])[s][0].loc[a]
        w = world_aggr.get_data(matrices=['w'], scenarios=[s])[s][0]
        f[a][s] = np.diag(e) @ w
        f[a][s].index = f[a][s].columns
        
    f[a]['delta'] = f[a]['Endogenization of capital'] - f[a]['baseline']
    for k,v in f[a].items():
        df_to_export = v.loc[(slice(None),"Commodity",slice(None)),(slice(None),'Commodity',commodities)]
        df_to_export.columns = pd.MultiIndex.from_tuples(df_to_export.columns) 
        df_to_export.to_excel(writer, sheet_name=k)
    writer.close()

#%% Endogenization of capital database to excel
world.to_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\d. Shock - Endogenization of capital.xlsx", flows=False, coefficients=True)

#%% Parse endogenization of capital from excel
world = mario.parse_from_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\d. Shock - Endogenization of capital.xlsx", table='SUT', mode="coefficients")

#%% Linkages for shock
world_iot = world.to_iot(method='B', inplace=False)
linkages['Endogenization of capital'] = world_iot.calc_linkages(multi_mode=False)

#%%
writer = pd.ExcelWriter(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\Linkages.xlsx", engine='openpyxl', mode='w')
for kind,link in linkages.items():   
    link.to_excel(writer, sheet_name=kind, merge_cells=False)
writer.close()



