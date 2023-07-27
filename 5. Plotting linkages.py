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
import plotly.graph_objects as go
from plotly.subplots import make_subplots

user = "LR"
sN = slice(None)

paths = 'Paths.xlsx'

commodities = [
    "Electricity by coal",
    "Electricity by gas",
    "Electricity by hydro",
    "Electricity by nuclear",
    "Electricity by petroleum and other oil derivatives",
    "Electricity by biomass and waste",
    "Electricity by solar photovoltaic",
    "Electricity by wind",
    "Electricity by Geothermal",
    "Electricity by solar thermal",
    "Electricity by tide, wave, ocean",
    ]

scenarios = ['EXIOHSUT']
years = [2011,2015,2019]

#%% Importing saved linkages
linkages = pd.read_csv(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\Linkages.csv")#¶, index_col=[0])

#%% Plotting Linkages analyzable by commodity
linkages_plot = linkages.query("Scope=='Total' & Item in @commodities & Scenario in @scenarios & Year in @years")
linkages_plot = linkages_plot.sort_values(['Region','Item','Scenario','Year','Performance'], ascending=[True,True,True,True,True])  

auto=True
colors = px.colors.qualitative.Pastel
template = "seaborn"
font = "HelveticaNeue Light"
size = 16
labels = {
    'Forward': ':.3f',
    }

fig_linkages = px.scatter(
    linkages_plot,
    x='Forward',
    y='Backward',
    color = 'Item',
    animation_frame = 'Region',
    facet_col = "Performance",
    facet_row = "Year",
    color_discrete_sequence=colors,
    )

fig_linkages.update_layout(title="Forward & Backward Linkages", font_family=font, font_size=size, template=template)
fig_linkages.for_each_annotation(lambda a: a.update(text=a.text.split("=")[1]))
fig_linkages.write_html(f"{pd.read_excel(paths, index_col=[0]).loc['Plots',user]}\\Linkages.html", auto_open=auto)


