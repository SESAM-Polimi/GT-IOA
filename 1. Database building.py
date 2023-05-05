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

#%% Parsing raw Exiobase
exio_sut = {}
exio_iot = {}

for year in years:
    exio_sut_path  = pd.read_excel(paths, index_col=[0]).loc['EXIOBASE SUT',user]+f"/MRSUT_{year}.zip"
    exio_iot_path  = pd.read_excel(paths, index_col=[0]).loc['EXIOBASE IOT',user]+f"/IOT_{year}_ixi.zip"

    exio_sut[year]  = mario.parse_exiobase(path=exio_sut_path, table='SUT', unit='Monetary')
    exio_iot[year]  = mario.parse_exiobase(path=exio_iot_path, table='IOT', unit='Monetary')

#%% Getting excel templates to aggregate raw Exiobase
path_aggr_IOT  = r"Aggregations\Aggregation_raw_IOT.xlsx"
path_aggr_SUT  = r"Aggregations\Aggregation_raw_SUT.xlsx"
# exio_iot[year].get_aggregation_excel(path_aggr_IOT)
# exio_sut[year].get_aggregation_excel(path_aggr_SUT)

#%% Aggregating exiobase as other models
for year in years:
    exio_iot[year].aggregate(path_aggr_IOT, levels=["Region"])
    exio_sut[year].aggregate(path_aggr_SUT, levels=["Region"])

#%% Adding environmental extensions from miot to msut
for year in years:
    sat_IOT = exio_iot[year].E
    sat_SUT = exio_sut[year].E
    
    new_sat_SUT = pd.DataFrame(0, index=sat_IOT.index, columns=sat_SUT.columns)
    new_sat_SUT.loc[:,(slice(None),'Activity')] = sat_IOT.values
    new_units = exio_iot[year].units['Satellite account'][9:]
    
    exio_sut[year].add_extensions(io= new_sat_SUT, matrix= 'E', units= new_units, inplace=True)

#%% Aggregated database to excel
for year in years:
    exio_sut[year].to_txt(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\a. Aggregated_SUT\\{year}", flows=False, coefficients=True)
