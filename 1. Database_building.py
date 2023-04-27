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

#%% Parsing raw Exiobase
exio_sut_path  = pd.read_excel(paths, index_col=[0]).loc['EXIOBASE SUT',user]
exio_iot_path  = pd.read_excel(paths, index_col=[0]).loc['EXIOBASE IOT',user]

exio_sut  = mario.parse_exiobase_msut(exio_sut_path)
exio_iot = mario.parse_exiobase_miot(exio_iot_path)

#%% Adding environmental extensions from miot to msut
sat_IOT = exio_iot.E
sat_SUT = exio_sut.E

new_sat_SUT = pd.DataFrame(0, index=sat_IOT.index, columns=sat_SUT.columns)
new_sat_SUT.loc[:,(slice(None),'Activity')] = sat_IOT.values
new_units = exio_iot.units['Satellite account'][9:]

exio_sut.add_extensions(io= new_sat_SUT, matrix= 'E', units= new_units, inplace=True)
    
#%% Getting excel templates to aggregate raw Exiobase
path_aggr  = r"Aggregations\Aggregation_baseline.xlsx"
# exio_sut.get_aggregation_excel(path_aggr)

#%% Aggregating exiobase as other models
exio_sut.aggregate(path_aggr, levels="Region")

#%% Aggregated database to excel
exio_sut.to_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\a. Aggregated_SUT.xlsx", flows=False, coefficients=True)
