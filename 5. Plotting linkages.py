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
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

user = "LR"
sN = slice(None)

paths = 'Paths.xlsx'

scenarios = ['baseline', 'Endogenization of capital']
region = "EU27+UK"
commodities = [
    "Electricity by PV",
    "Electricity by coal",
    "Electricity by gas",
    "Electricity by nuclear",
    "Electricity by oil",
    "Electricity by other RES",
    "Electricity by wind",
    ]

# #%% Parse databases from excel
# world_sut_base  = mario.parse_from_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\c. Baseline.xlsx", table='SUT', mode="coefficients")
# world_sut_encap = mario.parse_from_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\d. Shock - Endogenization of capital.xlsx", table='SUT', mode="coefficients")

# #%% Transforming the two databases into one 2-scenarios iot table
# world_iot  = world_sut_base.to_iot(method="B", inplace=False)
# world_iot_encap = world_sut_encap.to_iot(method="B", inplace=False)

# world_iot.clone_scenario('baseline','Endogenization of capital')
# world_iot.matrices['Endogenization of capital'] = world_iot_encap.matrices['baseline']   

# #%% Aggregating
# path_aggr = r"Aggregations\Aggregation_linkages.xlsx"
# world_iot.aggregate(path_aggr, levels=["Sector"])

# #%% Calculating and saving linkages
# linkages = {}
# writer = pd.ExcelWriter(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\Linkages.xlsx", engine='openpyxl', mode='w')
# for s in world_iot.scenarios:
#     linkages[s] = world_iot.calc_linkages(
#         scenario=s,
#         multi_mode=False,
#         )
#     columns = pd.MultiIndex.from_arrays([
#         [i.split(" ")[0] for i in linkages[s].columns],
#         [i.split(" ")[1] for i in linkages[s].columns],
#         ])
#     linkages[s].columns = columns
#     linkages[s] = linkages[s].droplevel(1)
#     linkages[s] = linkages[s].stack(level=[0])
#     linkages[s].index.names = ['Region','Commodity','Scope']
#     linkages[s].reset_index(inplace=True)
#     linkages[s].to_excel(writer, sheet_name=s)
# writer.close()

#%% Importing saved linkages
linkages = pd.DataFrame()
for s in scenarios:
    links = pd.read_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\Linkages.xlsx", sheet_name=s, index_col=[0])
    links.loc[:,"Scenario"] = s
    linkages = pd.concat([linkages, links])

linkages.replace("baseline", "Baseline", inplace=True)
linkages.replace("Endogenization of capital", "Endogenous capital", inplace=True)

#%% Plotting Linkages analyzable by commodity
linkages_plot = linkages.query("Scope=='Total' & Commodity in @commodities")
linkages_plot = linkages_plot.sort_values(['Region','Commodity','Scenario'], ascending=[True,True,True])  

auto=True
colors = px.colors.qualitative.Pastel
template = "seaborn"
font = "HelveticaNeue Light"
size = 16
labels = {
    'Forward': ':.3f',
    }
cols = sorted(list((set(linkages_plot['Scenario']))))

regions_markers = {
    'China': 'circle',
    'EU27+UK': 'square',
    'RoW': 'diamond',
    'USA': 'x',
    }
commodities_colors = {
    'Electricity by PV': 'rgb(246, 207, 113)', 
    'Electricity by coal': 'rgb(179, 179, 179)', 
    'Electricity by gas': 'rgb(220, 176, 242)', 
    'Electricity by nuclear': 'rgb(248, 156, 116)', 
    'Electricity by oil': 'rgb(254, 136, 177)', 
    'Electricity by other RES': 'rgb(102, 197, 204)', 
    'Electricity by wind': 'rgb(135, 197, 95)', 
    }

commodities_legend = sorted(list(set(linkages_plot['Commodity'])))

fig_linkages = make_subplots(rows=1,cols=len(cols), subplot_titles=tuple([i for i in cols]), shared_yaxes=True)
linkages_plot.index = [i for i in range(linkages_plot.shape[0])]

"First legend groups"
first_legend_labels = []
first_legend_title = "Commodities"
for trace in linkages_plot.index:
    for scenario in cols:

        if trace != 0 or scenario != cols[0]:
            first_legend_title = None
        
        first_label = linkages_plot.loc[trace,"Commodity"]    
        if first_label in first_legend_labels:
            show_legend = False
        else:
            first_legend_labels += [first_label]
            show_legend = True
        
        fig_linkages.add_trace(
            go.Scatter(
                x=[linkages_plot.loc[trace,"Forward"]],
                y=[linkages_plot.loc[trace,"Backward"]],
                name=first_label,
                showlegend=show_legend,
                legendgroup = linkages_plot.loc[trace,"Commodity"],
                legendgrouptitle_text = first_legend_title,
                mode = "markers",
                marker = dict(
                    color=commodities_colors[linkages_plot.loc[trace,"Commodity"]], 
                    size=10, 
                    symbol=regions_markers[linkages_plot.loc[trace,"Region"]]
                    ),
            ),
            row=1, 
            col=cols.index(linkages_plot.loc[trace,"Scenario"])+1
        )

"Fake trace to fake spacing in legend"
fig_linkages.add_trace(
    go.Scatter(
        x=[None],
        y=[None],
        name='',
        showlegend=True,
        mode = "markers",
        marker = dict(color='white'),
    ),
    row=1, 
    col=1,
)


"Second legend groups"
second_legend_labels = []
second_legend_title = "Regions"
for trace in linkages_plot.index:        
    for scenario in cols:

        if trace != 0 or scenario != cols[0]:
            second_legend_title = None
        
        second_label = linkages_plot.loc[trace,"Region"]    
        if second_label in second_legend_labels:
            show_legend = False
        else:
            second_legend_labels += [second_label]
            show_legend = True
        
        fig_linkages.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                name=second_label,
                showlegend=show_legend,
                legendgroup = linkages_plot.loc[trace,"Region"],
                legendgrouptitle_text = second_legend_title,
                mode = "markers",
                marker = dict(
                    color='black', 
                    size=10, 
                    symbol=regions_markers[linkages_plot.loc[trace,"Region"]]
                    ),
            ),
            row=1, 
            col=cols.index(linkages_plot.loc[trace,"Scenario"])+1
        )


fig_linkages.update_layout(title="Forward & Backward Linkages, explorable by Commodity", font_family=font, font_size=size, template=template)
fig_linkages.layout.legend.tracegroupgap = 1
fig_linkages.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\Linkages_com.html", auto_open=auto)


#%% Plotting Linkages analyzable by region
linkages_plot = linkages.query("Scope=='Total' & Commodity in @commodities")
linkages_plot = linkages_plot.sort_values(['Region','Commodity','Scenario'], ascending=[True,True,True])  

auto=True
colors = px.colors.qualitative.Pastel
template = "seaborn"
font = "HelveticaNeue Light"
size = 16
labels = {
    'Forward': ':.3f',
    }
cols = sorted(list((set(linkages_plot['Scenario']))))

regions_markers = {
    'China': 'circle',
    'EU27+UK': 'square',
    'RoW': 'diamond',
    'USA': 'x',
    }
commodities_colors = {
    'Electricity by PV': 'rgb(246, 207, 113)', 
    'Electricity by coal': 'rgb(179, 179, 179)', 
    'Electricity by gas': 'rgb(220, 176, 242)', 
    'Electricity by nuclear': 'rgb(248, 156, 116)', 
    'Electricity by oil': 'rgb(254, 136, 177)', 
    'Electricity by other RES': 'rgb(102, 197, 204)', 
    'Electricity by wind': 'rgb(135, 197, 95)', 
    }

commodities_legend = sorted(list(set(linkages_plot['Commodity'])))

fig_linkages = make_subplots(rows=1,cols=len(cols), subplot_titles=tuple([i for i in cols]), shared_yaxes=True)
linkages_plot.index = [i for i in range(linkages_plot.shape[0])]

"First legend groups"
first_legend_labels = []
first_legend_title = "Commodities"
for trace in linkages_plot.index:
    for scenario in cols:

        if trace != 0 or scenario != cols[0]:
            first_legend_title = None
        
        first_label = linkages_plot.loc[trace,"Commodity"]    
        if first_label in first_legend_labels:
            show_legend = False
        else:
            first_legend_labels += [first_label]
            show_legend = True
        
        fig_linkages.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                name=first_label,
                showlegend=show_legend,
                legendgroup = linkages_plot.loc[trace,"Commodity"],
                legendgrouptitle_text = first_legend_title,
                mode = "markers",
                marker = dict(
                    color=commodities_colors[linkages_plot.loc[trace,"Commodity"]], 
                    size=10, 
                    symbol=regions_markers[linkages_plot.loc[trace,"Region"]]
                    ),
            ),
            row=1, 
            col=cols.index(linkages_plot.loc[trace,"Scenario"])+1
        )

"Fake trace to fake spacing in legend"
fig_linkages.add_trace(
    go.Scatter(
        x=[None],
        y=[None],
        name='',
        showlegend=True,
        mode = "markers",
        marker = dict(color='white'),
    ),
    row=1, 
    col=1,
)


"Second legend groups"
second_legend_labels = []
second_legend_title = "Regions"
for trace in linkages_plot.index:        
    for scenario in cols:

        if trace != 0 or scenario != cols[0]:
            second_legend_title = None
        
        second_label = linkages_plot.loc[trace,"Region"]    
        if second_label in second_legend_labels:
            show_legend = False
        else:
            second_legend_labels += [second_label]
            show_legend = True
        
        fig_linkages.add_trace(
            go.Scatter(
                x=[linkages_plot.loc[trace,"Forward"]],
                y=[linkages_plot.loc[trace,"Backward"]],
                name=second_label,
                showlegend=show_legend,
                legendgroup = linkages_plot.loc[trace,"Region"],
                legendgrouptitle_text = second_legend_title,
                mode = "markers",
                marker = dict(
                    color='black', 
                    size=10, 
                    symbol=regions_markers[linkages_plot.loc[trace,"Region"]]
                    ),
            ),
            row=1, 
            col=cols.index(linkages_plot.loc[trace,"Scenario"])+1
        )

        fig_linkages.add_trace(
            go.Scatter(
                x=[linkages_plot.loc[trace,"Forward"]],
                y=[linkages_plot.loc[trace,"Backward"]],
                name=second_label,
                showlegend=False,
                legendgroup = linkages_plot.loc[trace,"Region"],
                legendgrouptitle_text = second_legend_title,
                mode = "markers",
                marker = dict(
                    color=commodities_colors[linkages_plot.loc[trace,"Commodity"]], 
                    size=10, 
                    symbol=regions_markers[linkages_plot.loc[trace,"Region"]]
                    ),
            ),
            row=1, 
            col=cols.index(linkages_plot.loc[trace,"Scenario"])+1
        )


fig_linkages.update_layout(title="Forward & Backward Linkages, explorable by Region", font_family=font, font_size=size, template=template)
fig_linkages.layout.legend.tracegroupgap = 1
fig_linkages.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\Linkages_reg.html", auto_open=auto)

