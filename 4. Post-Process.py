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
            "conv": 3600,
            },
        'CO2 - combustion - air': {
            "raw": 'kg',
            "new": 'ton',
            "conv": 1/1000,
            }, 
        'CH4 - combustion - air': {
            "raw": 'kg',
            "new": 'ton',
            "conv": 1/1000,
            }, 
        'N2O - combustion - air': {
            "raw": 'kg',
            "new": 'ton',
            "conv": 1/1000,
            }, 
        'GHGs': {
            "raw": 'kg',
            "new": 'tonCO2eq',
            "conv": 1/1000,
            }, 
        },
    'Activity': {
        "Offshore wind plants": {
            "raw": 'EUR',
            "new": 'MW',
            "conv": 3.19e6,
            }, 
        "Onshore wind plants": {
            "raw": 'EUR',
            "new": 'MW',
            "conv": 1.44e6,
            }, 
        "PV plants": {
            "raw": 'EUR',
            "new": 'MW',
            "conv": 1.81e6,
            }, 
        "Electricity by wind": {
            "raw": 'EUR',
            "new": 'GWh',
            "conv": 0.217e6,
            }, 
        "Electricity by PV": {
            "raw": 'EUR',
            "new": 'GWh',
            "conv": 0.217e6,
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
activities_to = [
    "Offshore wind plants",
    "Onshore wind plants",
    "PV plants",
    'Electricity by wind',
    "Electricity by PV",
    ]

not_activities_from = [
    "DFIG generators",
    "PMG generators",
    "Offshore wind plants",
    "Onshore wind plants",
    "PV plants",
    'Electricity by wind',
    "Electricity by PV",
    ]

capacity_figure = ["Offshore wind plants","Onshore wind plants","PV plants"]
energy_figure = ["Electricity by wind","Electricity by PV"]

#%% Reading and rearranging footprints results
f = {}
for sa in sat_accounts:
    f[sa] = pd.DataFrame()
    for scen in scenarios:
        f_sa_scen = pd.read_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\{sa}.xlsx", sheet_name=scen, index_col=[0,1,2], header=[0,1,2]).loc[(sN,"Activity",sN),(sN,"Activity",sN)]
        f_sa_scen = f_sa_scen.stack(level=[0,1,2])
        f_sa_scen = f_sa_scen.to_frame()
        f_sa_scen.columns = ['Value']
        f_sa_scen["Account"] = sa
        f_sa_scen["Scenario"] = scen
        f_sa_scen = f_sa_scen.droplevel(level=[1,4], axis=0)
        activities_from = []
        for act in set(f_sa_scen.index.get_level_values(1)):
            if act not in not_activities_from:
                activities_from += [act]
        f_sa_scen.index.names = ["Region from", "Activity from", "Region to", "Activity to"]
        f_sa_scen = f_sa_scen.loc[(sN,activities_from,regions_to,activities_to),:]
        f_sa_scen.reset_index(inplace=True)
        f[sa] = pd.concat([f[sa], f_sa_scen], axis=0)
    f[sa].replace("baseline", "Baseline", inplace=True)
    f[sa].replace("Endogenization of capital", "Endogenous capital", inplace=True)
    f[sa].replace("delta", "Variation", inplace=True)
    f[sa].set_index(["Region from", "Activity from", "Region to", "Activity to","Scenario","Account"], inplace=True)
    f[sa] = f[sa].groupby(level=f[sa].index.names).mean()
    
#%% Conversions to physical units
for sa,footprint in f.items():
    for i in footprint.index:
        footprint.loc[i,"Unit"] = f"{units['Satellite account'][sa]['new']}/{units['Activity'][i[3]]['new']}"
        footprint.loc[i,"Value"] *= units['Satellite account'][sa]['conv']*units['Activity'][i[3]]['conv']
    footprint.set_index(['Unit'], append=True, inplace=True)

#%% Calculation of total GHG emissions
f['GHGs'] = pd.DataFrame()
for sa,gwp in GWP.items():
    f['GHGs'] = pd.concat([f['GHGs'], f[sa]*gwp], axis=0)
f['GHGs'] = f['GHGs'].groupby(level=["Region from","Activity from","Region to","Activity to","Scenario"]).sum()
for i in f['GHGs'].index:
    f['GHGs'].loc[i,"Account"] = "GHG emmissions"
    f['GHGs'].loc[i,"Unit"] = f"{units['Satellite account']['GHGs']['new']}/{units['Activity'][i[3]]['new']}"
f['GHGs'].set_index(['Account','Unit'], append=True, inplace=True)
        
#%% Aggregating
new_activities = pd.read_excel(r"Aggregations\Aggregation_plots.xlsx", index_col=[0])
for sa,v in f.items():
    index_names = list(v.index.names)
    for i in v.index:
        v.loc[i,"Activity from"] = new_activities.loc[i[1],"New"]
    v = v.droplevel("Activity from", axis=0)
    v.reset_index(inplace=True)
    v.set_index(index_names, inplace=True)
    v = v.groupby(level=index_names, axis=0).sum()
    f[sa] = v

#%%
auto = True
to_plot = {
    'GHGs': 'GHG emissions',
    'Energy Carrier Supply - Total': 'Primary energy'
    }
colors = px.colors.qualitative.Pastel
template = "seaborn"
font = "HelveticaNeue Light"
size = 15

for sa in to_plot.keys():
    footprint_by_reg = f[sa].groupby(level=["Region from","Activity to", "Scenario","Unit"]).sum()
    footprint_by_act = f[sa].groupby(level=["Activity from","Activity to", "Scenario","Unit"]).sum()
    
    capacity_footprint_by_reg = footprint_by_reg.loc[(sN,capacity_figure,'Endogenous capital',sN),:].sort_values(['Region from','Scenario'], ascending=[True,True])
    capacity_footprint_by_act = footprint_by_act.loc[(sN,capacity_figure,'Endogenous capital',sN),:].sort_values(['Activity from','Scenario'], ascending=[True,True])  
    energy_footprint_by_reg = footprint_by_reg.loc[(sN,energy_figure,sN,sN),:].sort_values(['Region from','Scenario'], ascending=[True,True])
    energy_footprint_by_act = footprint_by_act.loc[(sN,energy_figure,sN,sN),:].sort_values(['Activity from', 'Scenario'], ascending=[True,True])
    
    capacity_footprint_by_reg.reset_index(inplace=True)
    capacity_footprint_by_act.reset_index(inplace=True)
    energy_footprint_by_reg.reset_index(inplace=True)
    energy_footprint_by_act.reset_index(inplace=True)

    fig_cap_by_reg = px.bar(capacity_footprint_by_reg, x="Activity to", y="Value", color="Region from", color_discrete_sequence=colors, title=f"{to_plot[sa]} footprint per unit of installed capacity, allocated by region", template=template)
    fig_cap_by_act = px.bar(capacity_footprint_by_act, x="Activity to", y="Value", color="Activity from", color_discrete_sequence=colors, title=f"{to_plot[sa]} footprint per unit of installed capacity, allocated by sector", template=template)
    fig_ene_by_reg = px.bar(energy_footprint_by_reg, x="Activity to", y="Value", color="Region from", facet_col="Scenario", color_discrete_sequence=colors, title=f"{to_plot[sa]} footprint per unit of electricity produced, allocated by region", template=template)
    fig_ene_by_act = px.bar(energy_footprint_by_act, x="Activity to", y="Value", color="Activity from", facet_col="Scenario", color_discrete_sequence=colors, title=f"{to_plot[sa]} footprint per unit of electricity produced, allocated by sector", template=template)

    fig_cap_by_reg.update_layout(legend=dict(title=None), xaxis=dict(title=None), yaxis=dict(title=f"{list(set(capacity_footprint_by_reg['Unit']))[0]}"), font_family=font, font_size=size)    
    fig_cap_by_act.update_layout(legend=dict(title=None), xaxis=dict(title=None), yaxis=dict(title=f"{list(set(capacity_footprint_by_act['Unit']))[0]}"), font_family=font, font_size=size)    
    fig_ene_by_reg.update_layout(legend=dict(title=None), yaxis=dict(title=f"{list(set(energy_footprint_by_reg['Unit']))[0]}"), font_family=font, font_size=size)    
    fig_ene_by_act.update_layout(legend=dict(title=None), yaxis=dict(title=f"{list(set(energy_footprint_by_act['Unit']))[0]}"), font_family=font, font_size=size)    

    fig_cap_by_reg.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
    fig_cap_by_act.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
    fig_ene_by_reg.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
    fig_ene_by_act.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
    
    fig_ene_by_reg.for_each_xaxis(lambda axis: axis.update(title=None))
    fig_ene_by_act.for_each_xaxis(lambda axis: axis.update(title=None))

    fig_cap_by_reg.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{to_plot[sa]} - Cap,reg.html", auto_open=auto)
    fig_cap_by_act.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{to_plot[sa]} - Cap,act.html", auto_open=auto)
    fig_ene_by_reg.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{to_plot[sa]} - Ene,reg.html", auto_open=auto)
    fig_ene_by_act.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{to_plot[sa]} - Ene,act.html", auto_open=auto)



