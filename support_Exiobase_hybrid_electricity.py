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

#%% Parse hybrid exiobase (2011)
world = mario.parse_exiobase_hsut(f"{pd.read_excel(paths, index_col=[0]).loc['EXIOBASE Hybrid',user]}")

#%% Aggregating 
path_aggr  = r"Aggregations\Aggregation_hybrid.xlsx"
world.aggregate(path_aggr, levels="Region")

#%%
electricity_comms = world.search("Commodity", "Electricity")[:-2]
ee_prod_TJ = world.X.loc[(slice(None),slice(None),electricity_comms)].groupby(level=[0], axis=0).sum() 
ee_prod_TJ = ee_prod_TJ.sort_index()
