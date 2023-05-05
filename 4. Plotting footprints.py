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

price_logics = ['Constant', 'IEA', 'EXIOHSUT']
years = range(2011,2020)
tech_performances = ['Worst','Average','Best']

#%%
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
            "conv": 1/3.6,
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
    'Commodity': {
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
            "conv": 'price',
            }, 
        "Electricity by PV": {
            "raw": 'EUR',
            "new": 'GWh',
            "conv": 'price',
            }, 
        },
    }

GWP = {
       "CO2 - combustion - air": 1,
       "CH4 - combustion - air": 26,
       "N2O - combustion - air": 298,
    }

regions_to = ['EU27+UK']
activities_to = [
    "Offshore wind plants",
    "Onshore wind plants",
    "PV plants",
    'Electricity by wind',
    "Electricity by PV",
    ]

scemarios = []
for y in years:
    for s in price_logics:
        for t in tech_performances:
            scemarios += [f"{s} - {y} - {t}"]

#%% Reading and rearranging footprints results
# f = {}
# for sa in sat_accounts:
#     f[sa] = pd.DataFrame()
#     for scem in scemarios:
#         scen = scem.split(' - ')[0]
#         year = scem.split(' - ')[1]
#         tech = scem.split(' - ')[2]
#         f_sa_scen = pd.read_csv(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\Footprints - Monetary units\\{sa}\\{scen} - {year} - {tech}.csv", index_col=[0,1,2], header=[0,1,2], sep=',').loc[(sN,"Activity",sN),(sN,"Commodity",sN)]
#         f_sa_scen = f_sa_scen.stack(level=[0,1,2])
#         f_sa_scen = f_sa_scen.to_frame()
#         f_sa_scen.columns = ['Value']
#         f_sa_scen["Account"] = sa
#         f_sa_scen["Scenario"] = f"{scen} - {year} - {tech}"
#         f_sa_scen = f_sa_scen.droplevel(level=[1,4], axis=0)
#         f_sa_scen.index.names = ["Region from", "Commodity", "Region to", "Activity to"]
#         f_sa_scen = f_sa_scen.loc[(sN,sN,regions_to,activities_to),:]
#         f_sa_scen.reset_index(inplace=True)
#         f[sa] = pd.concat([f[sa], f_sa_scen], axis=0)
#     f[sa].replace("baseline", "Baseline", inplace=True)
#     f[sa].set_index(["Region from", "Commodity", "Region to", "Activity to","Scenario","Account"], inplace=True)
#     f[sa] = f[sa].groupby(level=f[sa].index.names).mean()
    
#%% Conversions to physical units
# for sa,footprint in f.items():
#     for i in footprint.index:
#         footprint.loc[i,"Unit"] = f"{units['Satellite account'][sa]['new']}/{units['Commodity'][i[3]]['new']}"
#         if units['Commodity'][i[3]]['conv'] == 'price':
#             ee_price = pd.read_excel(f"{pd.read_excel(paths, index_col=[0]).loc['ShockMaster',user]}", sheet_name=f"{i[4].split(' - ')[0]}_Electricity prices", index_col=[0]).loc[i[2],int(i[4].split(' - ')[1])]
#             footprint.loc[i,"Value"] *= units['Satellite account'][sa]['conv']*ee_price*1e6
#         else:
#             footprint.loc[i,"Value"] *= units['Satellite account'][sa]['conv']*units['Commodity'][i[3]]['conv']
#     footprint.set_index(['Unit'], append=True, inplace=True)
    
    
#%% Saving converted footprints
# writer = pd.ExcelWriter(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\Footprints - Physical units.xlsx", engine='openpyxl', mode='w')
# for sa,footprint in f.items():
#     footprint.to_excel(writer, sheet_name=sa)
# writer.close()

#%% Read saved footprints in physical units
f = {}
for sa in sat_accounts:
    f[sa] = pd.read_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\Footprints - Physical units.xlsx", sheet_name=sa, index_col=[0,1,2,3,4,5,6])
    
#%% Calculation of total GHG emissions
f['GHGs'] = pd.DataFrame()
for sa,gwp in GWP.items():
    f['GHGs'] = pd.concat([f['GHGs'], f[sa]*gwp], axis=0)
f['GHGs'] = f['GHGs'].groupby(level=["Region from","Commodity","Region to","Activity to","Scenario"]).sum()
for i in f['GHGs'].index:
    f['GHGs'].loc[i,"Account"] = "GHG emmissions"
    f['GHGs'].loc[i,"Unit"] = f"{units['Satellite account']['GHGs']['new']}/{units['Commodity'][i[3]]['new']}"
f['GHGs'].set_index(['Account','Unit'], append=True, inplace=True)

#%% Split scemarios columns
for sa,footprint in f.items():
    footprint.loc[:,'Scenario'] = [i.split(' - ')[0] for i in footprint.index.get_level_values("Scenario")]
    footprint.loc[:,'Year'] = [i.split(' - ')[1] for i in footprint.index.get_level_values("Scenario")]
    footprint.loc[:,'Performance'] = [i.split(' - ')[2] for i in footprint.index.get_level_values("Scenario")]
    footprint = footprint.droplevel("Scenario")
    footprint.reset_index(inplace=True)
    footprint.set_index(['Region from', 'Commodity', 'Region to', 'Activity to', 'Scenario', 'Year', 'Performance', 'Account', 'Unit'], inplace=True)
    f[sa] = footprint
        
#%% Aggregating
new_commodities = pd.read_excel(r"Aggregations\Aggregation_plots.xlsx", index_col=[0])
for sa,v in f.items():
    index_names = list(v.index.names)
    for i in v.index:
        v.loc[i,"Commodity"] = new_commodities.loc[i[1],"New"]
    v = v.droplevel("Commodity", axis=0)
    v.reset_index(inplace=True)
    v.set_index(index_names, inplace=True)
    v = v.groupby(level=index_names, axis=0).sum()
    f[sa] = v

#%% Plotting
auto = True
to_plot = {
    'Capacity': {
        'x': 'Year',
        'y': 'Value',
        # 'facet_row': 'Region from',
        'facet_col': 'Region from',
        'color': 'Commodity',
        'title': 'installed capacity',
        # 'animation_frame': 'Scenario',
        'activities': ["Offshore wind plants","Onshore wind plants","PV plants"],
        'satellite accounts': {
            'GHGs': {
                'name': 'GHG emissions',
                },
            'Energy Carrier Supply - Total': {
                'name': 'Primary energy',
                },
            },
        },
    'Electricity': {
        'x': 'Performance',
        'y': 'Value',
        'facet_row': 'Region from',
        'facet_col': 'Year',
        'color': 'Commodity',
        'title': 'electricity produced',
        'animation_frame': 'Scenario',
        'activities': ["Electricity by wind","Electricity by PV"],
        'satellite accounts': {
            'GHGs': {
                'name': 'GHG emissions',
                },
            'Energy Carrier Supply - Total': {
                'name': 'Primary energy',
                },
            },
        },
    }


colors = px.colors.qualitative.Pastel
template = "seaborn"
font = "HelveticaNeue Light"
size = 16
labels = {
    'Activity to': False,   
    # 'Scenario': False,
    'Value': ':.2f',
    }

for plot,properties in to_plot.items():
    for sa,name in properties['satellite accounts'].items():
        for act in properties['activities']:

            plot_df = f[sa].groupby(level=["Region from","Commodity","Activity to", "Scenario", "Year","Performance","Unit"]).sum() 

            if plot == 'Capacity':
                plot_df = f[sa].groupby(level=["Region from","Commodity","Activity to", "Year","Performance","Unit"]).mean() 
                plot_df = f[sa].groupby(level=["Region from","Commodity","Activity to", "Year","Unit"]).mean() 
                plot_df = plot_df.loc[(sN,sN,act,sN,sN,sN),:].sort_values(['Region from','Commodity','Year'], ascending=[False,False,True])  
            else:
                plot_df = plot_df.loc[(sN,sN,act,sN,sN,sN),:].sort_values(['Region from','Commodity','Scenario','Year'], ascending=[False,False,True,True])  
            
            plot_df.reset_index(inplace=True)
            
            if plot == 'Capacity':
                fig = px.bar(
                    plot_df, 
                    x=properties['x'],
                    y=properties['y'],
                    color=properties['color'],
                    facet_col=properties['facet_col'],
                    # facet_row=properties['facet_row'],
                    # animation_frame=properties['animation_frame'],
                    color_discrete_sequence=colors,
                    title=f"{sa} footprint per unit of {properties['title']} in EU27+UK, allocated by region and commodity [{list(set(plot_df['Unit']))[0]}]",
                    template=template, 
                    hover_data=labels,
                    )
            else:
                fig = px.bar(
                    plot_df, 
                    x=properties['x'],
                    y=properties['y'],
                    color=properties['color'],
                    facet_col=properties['facet_col'],
                    facet_row=properties['facet_row'],
                    animation_frame=properties['animation_frame'],
                    color_discrete_sequence=colors,
                    title=f"{sa} footprint per unit of {properties['title']} in EU27+UK, allocated by region and commodity [{list(set(plot_df['Unit']))[0]}]",
                    template=template, 
                    hover_data=labels,
                    )
    
            fig.update_layout(
                legend=dict(
                    title=None, 
                    traceorder='reversed'
                    ), 
                xaxis=dict(
                    title=None
                    ), 
                font_family=font, 
                font_size=size
                )    
    
            fig.update_traces(marker_line_width=0)
            fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
            fig.for_each_yaxis(lambda axis: axis.update(title=None))
            fig.for_each_xaxis(lambda axis: axis.update(title=None))
            fig.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{act}_{name['name']}.html", auto_open=auto)
    
    

#%%





    
    
    # footprint_by_reg = f[sa].groupby(level=["Region from","Activity to", "Scenario","Unit"]).sum()
    # footprint_by_act = f[sa].groupby(level=["Activity from","Activity to", "Scenario","Unit"]).sum()
    # footprint_by_reg_act = f[sa].groupby(level=["Region from","Activity from","Activity to", "Scenario","Unit"]).sum()
    
    # capacity_footprint_by_reg = footprint_by_reg.loc[(sN,capacity_figure,'Endogenous capital',sN),:].sort_values(['Region from','Scenario'], ascending=[False,True])
    # capacity_footprint_by_act = footprint_by_act.loc[(sN,capacity_figure,'Endogenous capital',sN),:].sort_values(['Activity from','Scenario'], ascending=[False,True])  
    # capacity_footprint_by_reg_act = footprint_by_reg_act.loc[(sN,sN,capacity_figure,'Endogenous capital',sN),:].sort_values(['Region from','Activity from','Scenario'], ascending=[False,False,True])  
    # energy_footprint_by_reg = footprint_by_reg.loc[(sN,energy_figure,['Baseline','Endogenous capital'],sN),:].sort_values(['Region from','Scenario'], ascending=[False,True])
    # energy_footprint_by_act = footprint_by_act.loc[(sN,energy_figure,['Baseline','Endogenous capital'],sN),:].sort_values(['Activity from', 'Scenario'], ascending=[False,True])
    # energy_footprint_by_reg_act = footprint_by_reg_act.loc[(sN,sN,energy_figure,['Baseline','Endogenous capital'],sN),:].sort_values(['Region from','Activity from','Scenario'], ascending=[False,False,True])  
    
    # capacity_footprint_by_reg.reset_index(inplace=True)
    # capacity_footprint_by_act.reset_index(inplace=True)
    # capacity_footprint_by_reg_act.reset_index(inplace=True)
    # energy_footprint_by_reg.reset_index(inplace=True)
    # energy_footprint_by_act.reset_index(inplace=True)
    # energy_footprint_by_reg_act.reset_index(inplace=True)
    
    # capacity_footprint_by_reg.rename(columns={'Activity from':'Commodity', 'Region from': 'Region'}, inplace=True)
    # capacity_footprint_by_act.rename(columns={'Activity from':'Commodity', 'Region from': 'Region'}, inplace=True)
    # capacity_footprint_by_reg_act.rename(columns={'Activity from':'Commodity', 'Region from': 'Region'}, inplace=True)
    # energy_footprint_by_reg.rename(columns={'Activity from':'Commodity', 'Region from': 'Region'}, inplace=True)
    # energy_footprint_by_act.rename(columns={'Activity from':'Commodity', 'Region from': 'Region'}, inplace=True)
    # energy_footprint_by_reg_act.rename(columns={'Activity from':'Commodity', 'Region from': 'Region'}, inplace=True)

    # fig_cap_by_reg = px.bar(capacity_footprint_by_reg, x="Activity to", y="Value", color="Region", color_discrete_sequence=colors, title=f"{to_plot[sa]} footprint per unit of installed capacity in EU27+UK, allocated by region", template=template, hover_data=labels)
    # fig_cap_by_act = px.bar(capacity_footprint_by_act, x="Activity to", y="Value", color="Commodity", color_discrete_sequence=colors, title=f"{to_plot[sa]} footprint per unit of installed capacity in EU27+UK, allocated by commodity", template=template, hover_data=labels)
    # fig_cap_by_reg_act = px.bar(capacity_footprint_by_reg_act, x="Activity to", y="Value", color="Commodity", facet_col="Region", color_discrete_sequence=colors, title=f"{to_plot[sa]} footprint per unit of installed capacity in EU27+UK, allocated by region and commodity [{list(set(capacity_footprint_by_act['Unit']))[0]}]", template=template, hover_data=labels)
    # fig_ene_by_reg = px.bar(energy_footprint_by_reg, x="Scenario", y="Value", color="Region", facet_col="Activity to", color_discrete_sequence=colors, title=f"{to_plot[sa]} footprint per unit of electricity produced in EU27+UK, allocated by region", template=template, hover_data=labels)
    # fig_ene_by_act = px.bar(energy_footprint_by_act, x="Scenario", y="Value", color="Commodity", facet_col="Activity to", color_discrete_sequence=colors, title=f"{to_plot[sa]} footprint per unit of electricity produced in EU27+UK, allocated by commodity", template=template, hover_data=labels)
    # fig_ene_by_reg_act = px.bar(energy_footprint_by_reg_act, x="Scenario", y="Value", color="Commodity", facet_col="Region", facet_row="Activity to", color_discrete_sequence=colors, title=f"{to_plot[sa]} footprint per unit of electricity produced in EU27+UK, allocated by region and commodity [{list(set(energy_footprint_by_act['Unit']))[0]}]", template=template, hover_data=labels)

    # fig_cap_by_reg.update_layout(legend=dict(title=None, traceorder='reversed'), xaxis=dict(title=None), yaxis=dict(title=f"{list(set(capacity_footprint_by_reg['Unit']))[0]}"), font_family=font, font_size=size)    
    # fig_cap_by_act.update_layout(legend=dict(title=None, traceorder='reversed'), xaxis=dict(title=None), yaxis=dict(title=f"{list(set(capacity_footprint_by_act['Unit']))[0]}"), font_family=font, font_size=size)    
    # fig_cap_by_reg_act.update_layout(legend=dict(title=None, traceorder='reversed'), xaxis=dict(title=None), font_family=font, font_size=size)    
    # fig_ene_by_reg.update_layout(legend=dict(title=None, traceorder='reversed'), yaxis=dict(title=f"{list(set(energy_footprint_by_reg['Unit']))[0]}"), font_family=font, font_size=size)    
    # fig_ene_by_act.update_layout(legend=dict(title=None, traceorder='reversed'), yaxis=dict(title=f"{list(set(energy_footprint_by_act['Unit']))[0]}"), font_family=font, font_size=size)    
    # fig_ene_by_reg_act.update_layout(legend=dict(title=None, traceorder='reversed'), font_family=font, font_size=size)    

    # fig_cap_by_reg.update_traces(marker_line_width=0)
    # fig_cap_by_act.update_traces(marker_line_width=0)
    # fig_cap_by_reg_act.update_traces(marker_line_width=0)
    # fig_ene_by_reg.update_traces(marker_line_width=0)
    # fig_ene_by_act.update_traces(marker_line_width=0)
    # fig_ene_by_reg_act.update_traces(marker_line_width=0)

    # fig_cap_by_reg.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
    # fig_cap_by_act.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
    # fig_cap_by_reg_act.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
    # fig_ene_by_reg.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
    # fig_ene_by_act.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
    # fig_ene_by_reg_act.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
    
    # fig_ene_by_reg.for_each_xaxis(lambda axis: axis.update(title=None))
    # fig_ene_by_act.for_each_xaxis(lambda axis: axis.update(title=None))
    # fig_cap_by_reg_act.for_each_xaxis(lambda axis: axis.update(title=None))
    # fig_ene_by_reg_act.for_each_xaxis(lambda axis: axis.update(title=None))

    # fig_cap_by_reg_act.for_each_yaxis(lambda axis: axis.update(title=None))
    # fig_ene_by_reg_act.for_each_yaxis(lambda axis: axis.update(title=None))

    # fig_cap_by_reg.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{to_plot[sa]} - Cap,reg.html", auto_open=auto)
    # fig_cap_by_act.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{to_plot[sa]} - Cap,act.html", auto_open=auto)
    # fig_cap_by_reg_act.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{to_plot[sa]} - Cap,reg-act.html", auto_open=auto)
    # fig_ene_by_reg.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{to_plot[sa]} - Ene,reg.html", auto_open=auto)
    # fig_ene_by_act.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{to_plot[sa]} - Ene,act.html", auto_open=auto)
    # fig_ene_by_reg_act.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\{to_plot[sa]} - Ene,reg-act.html", auto_open=auto)


