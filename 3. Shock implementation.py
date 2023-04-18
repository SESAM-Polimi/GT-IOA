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

#%% Parse aggregated  with new sectors database from excel
# world = mario.parse_from_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\Aggregated & new sectors SUT.xlsx", table='SUT', mode="flows")

#%% Getting shock templates
# world.get_shock_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\_template.xlsx")

#%% Implementing shock 0
# world.shock_calc(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\Shock 0 - Endogenization & carbon intensity fix.xlsx", v=True, e=True, z=True, scenario="shock 0")

#%% Shock 0 database to excel
# world.to_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\Database_final.xlsx")

#%% Parse shock 0 database from excel
world = mario.parse_from_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\Database_final.xlsx", table='SUT', mode="flows")

#%% Implementing shock 1.1 and 1.2
world.shock_calc(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\Shock 1.1 - Single MWh production ante endogenization.xlsx", Y=True, scenario="shock 1.1")
world.shock_calc(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\Shock 1.2 - Single MWh production post endogenization.xlsx", Y=True, v=True, z=True, scenario="shock 1.2")

#%% Implementing shock 2
world.shock_calc(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\Shock 2 - Single MW installed.xlsx",z=True,v=True,Y=True, scenario="shock 2")

#%% Implementing shock 3
world.shock_calc(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\Shock 3.1 - Avoided carbon intensive electricicty production.xlsx",z=True,v=True,Y=True,e=True,scenario="shock 3.1")
world.shock_calc(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\Shock 3.2 - Avoided carbon intensive electricicty production.xlsx",z=True,v=True,Y=True,e=True,scenario="shock 3.2")

#%% Aggregate for postprocess and visuals
path_aggr = r"Aggregations\Aggregation_postprocess.xlsx"
world.aggregate(io=path_aggr,levels=['Activity','Commodity','Factor of production','Satellite account','Consumption category'])

#%% Exporting results
world.matrices_export(
    export_folder = f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}",
    )
