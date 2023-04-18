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
scenario = 'shock 1.1'

#%% Calculate and plot linkages
def linkages_calc_plot(world,method,scenario,path,auto_open):
    world_IOT = world.to_iot(method=method, inplace=False)
    linkages  = world_IOT.calc_linkages(multi_mode=False)
    world_IOT.plot_linkages(path=f"{path}\{scenario}_linkages.html",normalized=False,auto_open=auto_open)
    world_IOT.plot_linkages(path=f"{path}\{scenario}linkages_norm.html",auto_open=auto_open)
    world_IOT.plot_linkages(path=f"{path}\{scenario}linkages_multi.html",normalized=False, multi_mode=True,auto_open=auto_open)
    return linkages

#%% Parse database from excel
world = mario.parse_from_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\Tables\{scenario}.xlsx", table='SUT', mode="flows")

#%%
path_aggr = r"Aggregations\Aggregation_linkages.xlsx"
world.aggregate(path_aggr, levels=['Activity','Commodity'])

#%%
linkages = linkages_calc_plot(
    world = world,
    method = 'B',
    scenario = scenario,
    path = pd.read_excel(paths, index_col=[0]).loc['Plots',user],
    auto_open = False,
    )
    