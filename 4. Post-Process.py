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


import pandas as pd
import plotly.express as px

user = "LR"
sN = slice(None)

paths = 'Paths.xlsx'

sat_accounts = [
    'Energy Carrier Supply - Total', 
    'CO2 - combustion - air', 
    'CH4 - combustion - air', 
    'N2O - combustion - air',
    ]

units = {
    'Satellite account': {
        'Energy Carrier Supply - Total': {
            "raw": 'TJ',
            "new": 'GWh',
            "conv": 1/3600,
            },
        'CO2 - combustion - air': {
            "raw": 'kg',
            "new": 'tonnes',
            "conv": 0.001,
            }, 
        'CH4 - combustion - air': {
            "raw": 'kg',
            "new": 'tonnes',
            "conv": 0.001,
            }, 
        'N2O - combustion - air': {
            "raw": 'kg',
            "new": 'tonnes',
            "conv": 0.001,
            }, 
        },
    'Commodity': {
        "Offshore wind plants": {
            "raw": 'MEUR',
            "new": 'MW',
            "conv": 3.19,
            }, 
        "Onshore wind plants": {
            "raw": 'MEUR',
            "new": 'MW',
            "conv": 1.44,
            }, 
        "PV plants": {
            "raw": 'MEUR',
            "new": 'MW',
            "conv": 1.81,
            }, 
        "Electricity by wind": {
            "raw": 'MEUR',
            "new": 'GWh',
            "conv": 0.217,
            }, 
        "Electricity by solar photovoltaic": {
            "raw": 'MEUR',
            "new": 'GWh',
            "conv": 0.217,
            }, 
        },
    }

GWP = {
       "CO2 - combustion - air": 1,
       "CH4 - combustion - air": 26,
       "N2O - combustion - air": 298,
    }

scenarios = ['baseline', 'Endogenization of capital', 'delta']

regions_to = ['EU27+UK']
commodities_to = [
    "Offshore wind plants",
    "Onshore wind plants",
    "PV plants",
    'Electricity by wind',
    "Electricity by solar photovoltaic",
    ]

not_commodities_from = [
    "DFIG generators",
    "PMG generators",
    "Offshore wind plants",
    "Onshore wind plants",
    "PV plants",
    'Electricity by wind',
    "Electricity by solar photovoltaic",
    ]

capacity_figure = ["Offshore wind plants","Onshore wind plants","PV plants"]
energy_figure = ["Electricity by wind","Electricity by solar photovoltaic"]

#%% Reading and rearranging footprints results
f = {}
for sa in sat_accounts:
    f[sa] = pd.DataFrame()
    for scen in scenarios:
        f_sa_scen = pd.read_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\{sa}.xlsx", sheet_name=scen, index_col=[0,1,2], header=[0,1,2])
        f_sa_scen = f_sa_scen.stack(level=[0,1,2])
        f_sa_scen = f_sa_scen.to_frame()
        f_sa_scen.columns = ['Value']
        f_sa_scen["Account"] = sa
        f_sa_scen["Scenario"] = scen
        f_sa_scen = f_sa_scen.droplevel(level=[1,4], axis=0)
        commodities_from = []
        for act in set(f_sa_scen.index.get_level_values(1)):
            if act not in not_commodities_from:
                commodities_from += [act]
        f_sa_scen.index.names = ["Region from", "Commodity from", "Region to", "Commodity to"]
        f_sa_scen = f_sa_scen.loc[(sN,commodities_from,regions_to,commodities_to),:]
        f_sa_scen.reset_index(inplace=True)
        f[sa] = pd.concat([f[sa], f_sa_scen], axis=0)
    f[sa].set_index(["Region from", "Commodity from", "Region to", "Commodity to","Scenario","Account"], inplace=True)
    f[sa] = f[sa].groupby(level=f[sa].index.names).mean()

#%% Conversions to physical units
for sa,footprint in f.items():
    for i in footprint.index:
        footprint.loc[i,"Unit"] = f"{units['Satellite account'][sa]['new']}/{units['Commodity'][i[3]]['new']}"
        footprint.loc[i,"Value"] *= units['Satellite account'][sa]['conv']/units['Commodity'][i[3]]['conv']
    footprint.set_index(['Unit'], append=True, inplace=True)

#%% Calculation of total GHG emissions
f['GHGs'] = pd.DataFrame()
for sa,gwp in GWP.items():
    f['GHGs'] = pd.concat([f['GHGs'], f[sa]*gwp], axis=0)
f['GHGs'] = f['GHGs'].groupby(level=["Region from","Commodity from","Region to","Commodity to","Scenario"]).sum()
for i in f['GHGs'].index:
    f['GHGs'].loc[i,"Account"] = "GHG emmissions"
    f['GHGs'].loc[i,"Unit"] = f"{units['Satellite account'][sa]['new']}/{units['Commodity'][i[3]]['new']}"
f['GHGs'].set_index(['Account','Unit'], append=True, inplace=True)
        
#%%
auto = False
for sa in f.keys():
    # footprint_by_reg = f[sa].groupby(level=["Region from","Activity to", "Scenario"]).sum()
    # footprint_by_act = f[sa].groupby(level=["Activity from","Activity to", "Scenario"]).sum()
    footprint_by_reg = f[sa].groupby(level=["Region from","Commodity to", "Scenario","Unit"]).sum()
    footprint_by_act = f[sa].groupby(level=["Commodity from","Commodity to", "Scenario","Unit"]).sum()
    
    capacity_footprint_by_reg = footprint_by_reg.loc[(sN,capacity_figure,sN,sN),:]
    capacity_footprint_by_act = footprint_by_act.loc[(sN,capacity_figure,sN,sN),:]    
    energy_footprint_by_reg = footprint_by_reg.loc[(sN,energy_figure,sN,sN),:]
    energy_footprint_by_act = footprint_by_act.loc[(sN,energy_figure,sN,sN),:]
    
    capacity_footprint_by_reg.reset_index(inplace=True)
    capacity_footprint_by_act.reset_index(inplace=True)
    energy_footprint_by_reg.reset_index(inplace=True)
    energy_footprint_by_act.reset_index(inplace=True)

    fig_cap_by_reg = px.bar(capacity_footprint_by_reg, x="Commodity to", y="Value", color="Region from", facet_col="Scenario")
    fig_cap_by_act = px.bar(capacity_footprint_by_act, x="Commodity to", y="Value", color="Commodity from", facet_col="Scenario")
    fig_ene_by_reg = px.bar(energy_footprint_by_reg, x="Commodity to", y="Value", color="Region from", facet_col="Scenario")
    fig_ene_by_act = px.bar(energy_footprint_by_act, x="Commodity to", y="Value", color="Commodity from", facet_col="Scenario")

    fig_cap_by_reg.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{sa} - Cap,reg.html", auto_open=auto)
    fig_cap_by_act.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{sa} - Cap,act.html", auto_open=auto)
    fig_ene_by_reg.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{sa} - Ene,reg.html", auto_open=auto)
    fig_ene_by_act.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{sa} - Ene,act.html", auto_open=auto)



