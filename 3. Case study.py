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
from mario.tools.constants import _MASTER_INDEX as MI
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import os

user = "LR"
sN = slice(None)
years = range(2011,2020)

paths = 'Paths.xlsx'

world = {}
price_logics = ['Constant', 'IEA', 'EXIOHSUT']
tech_performances = ['Worst','Average','Best']

#%% Parse aggregated  with new sectors database from excel
# for year in years:
#     world[year] = mario.parse_from_txt(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\b. Aggregated & new sectors SUT\\{year}\\coefficients", table='SUT', mode="coefficients")

#%% Getting shock templates
# world[year].get_shock_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\_template.xlsx")

#%% Implementing fixing of direct CO2 emissions
# for year in years:
#     world[year].shock_calc(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\\Raw dataset to Baseline model.xlsx", v=True, e=True, z=True, scenario="shock 0")

# %% Baseline database to excel
# for year in years:
#     world[year].to_txt(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\c. Baseline\\{year}", scenario="shock 0", flows=False, coefficients=True)

#%% Parse baseline from excel
for year in years:
    world[year] = mario.parse_from_txt(f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\c. Baseline\\{year}\\coefficients", table='SUT', mode="coefficients")
    
#%% Getting shock templates
for logic in price_logics:
    for year in years:
        for tech in tech_performances:
            world[year].get_shock_excel(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\\Shock_files\\{logic}_{year}_{tech}.xlsx")

#%% Filling shock templates
ShockMaster = pd.read_excel(f"{pd.read_excel(paths, index_col=[0]).loc['ShockMaster',user]}", sheet_name="ShockMaster", index_col=[0,1,2,3,4,5,6,7,8,9,10,11])
indexnames = list(ShockMaster.index.names)
ShockInput = pd.DataFrame(columns = ShockMaster.columns)

scemarios = sorted(list(set(ShockMaster.index.get_level_values('SceMARIO'))))

for scem in scemarios:
    scen = scem.split(' - ')[0]
    year = scem.split(' - ')[1]
    tech = scem.split(' - ')[2]
    ShockScemario = ShockMaster.loc[(sN,sN,sN,scem,sN,sN,sN,sN,sN,sN,sN),:]

    if scen == 'All' and year == 'All' and tech == 'All':
        for s in price_logics:
            for y in years:
                for t in tech_performances:
                    scenario_index = [s for i in ShockScemario.index]
                    year_index = [y for i in ShockScemario.index]
                    perf_index = [t for i in ShockScemario.index]
                    scemario_index = [f"{s} - {y} - {t}" for i in ShockScemario.index]
                    ShockScemario['Scenario'] = scenario_index
                    ShockScemario['Year'] = year_index
                    ShockScemario['Performance'] = perf_index
                    ShockScemario['SceMARIO'] = scemario_index
                    ShockScemario = ShockScemario.droplevel('Scenario')
                    ShockScemario = ShockScemario.droplevel('Year')
                    ShockScemario = ShockScemario.droplevel('Performance')
                    ShockScemario = ShockScemario.droplevel('SceMARIO')
                    ShockScemario.reset_index(inplace=True)
                    ShockScemario.set_index(['Scenario','Year','Performance','SceMARIO','Production region','Consumption region','Activity','Commodity','Factor of production','Matrix','Notes','Reference'], inplace=True)
                    ShockInput = pd.concat([ShockInput, ShockScemario])
        
    else:
        ShockInput = pd.concat([ShockInput, ShockScemario])
        

ShockInput.index = pd.MultiIndex.from_tuples(ShockInput.index)
ShockInput.index.names = ['Scenario','Year','Performance','SceMARIO','Production region','Consumption region','Activity','Commodity','Factor of production','Matrix','Notes','Reference']
ShockInput.reset_index(inplace=True)
ShockInput.drop_duplicates(inplace=True)
ShockInput.set_index(['Scenario','Year','Performance','SceMARIO','Production region','Consumption region','Activity','Commodity','Factor of production','Matrix','Notes','Reference'], inplace=True)


for scen in price_logics:
    for year in years:
        for tech in tech_performances:
            workbook = load_workbook(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\\Shock_files\\{logic}_{year}_{tech}.xlsx")
    
            add_data_s = ShockInput[ShockInput.index.get_level_values('Scenario') == scen].fillna(0)
            add_data_s = add_data_s[add_data_s.index.get_level_values('Year') == year].fillna(0)
            add_data_s = add_data_s[add_data_s.index.get_level_values('Performance') == tech].fillna(0)
            matrices = sorted(list(set(add_data_s.index.get_level_values('Matrix'))))
            
            s=0            
            for m in matrices:
                add_data_sm = add_data_s[add_data_s.index.get_level_values('Matrix') == m].fillna(0)
        
                if m == 's':
                    for k in range(len(add_data_sm)):
                        workbook['z']['A'+str(k+2)] = add_data_sm.index.get_level_values('Production region')[k]
                        workbook['z']['B'+str(k+2)] = MI['a']
                        workbook['z']['C'+str(k+2)] = add_data_sm.index.get_level_values('Activity')[k]
                        workbook['z']['D'+str(k+2)] = add_data_sm.index.get_level_values('Consumption region')[k]
                        workbook['z']['E'+str(k+2)] = MI['c']
                        workbook['z']['F'+str(k+2)] = add_data_sm.index.get_level_values('Commodity')[k]
                        workbook['z']['G'+str(k+2)] = 'Update'
                        workbook['z']['H'+str(k+2)] = add_data_sm.iloc[k,-1]
                    s+=1
        
                if m == 'u':
                    for u in range(len(add_data_sm)):
                        workbook['z']['A'+str(s+u+2)] = add_data_sm.index.get_level_values('Production region')[u]
                        workbook['z']['B'+str(s+u+2)] = MI['c']
                        workbook['z']['C'+str(s+u+2)] = add_data_sm.index.get_level_values('Commodity')[u]
                        workbook['z']['D'+str(s+u+2)] = add_data_sm.index.get_level_values('Consumption region')[s]
                        workbook['z']['E'+str(s+u+2)] = MI['a']
                        workbook['z']['F'+str(s+u+2)] = add_data_sm.index.get_level_values('Activity')[u]
                        workbook['z']['G'+str(s+u+2)] = 'Update'
                        workbook['z']['H'+str(s+u+2)] = add_data_sm.iloc[u,-1]
            
                elif m == 'Y':
                    for i in range(len(add_data_sm)):
                        workbook['Y']['A'+str(i+2)] = add_data_sm.index.get_level_values('Production region')[i]
                        workbook['Y']['B'+str(i+2)] = MI['c']
                        workbook['Y']['C'+str(i+2)] = add_data_sm.index.get_level_values('Commodity')[i]
                        workbook['Y']['D'+str(i+2)] = add_data_sm.index.get_level_values('Consumption region')[i]
                        workbook['Y']['E'+str(i+2)] = add_data_sm.index.get_level_values('Activity')[i]
                        workbook['Y']['F'+str(i+2)] = 'Update'
                        workbook['Y']['G'+str(i+2)] = add_data_sm.iloc[i,-1]
                    
                elif m=='v':
                    for i in range(len(add_data_sm)):
                        workbook['v']['A'+str(i+2)] = add_data_sm.index.get_level_values('Factor of production')[i]
                        workbook['v']['B'+str(i+2)] = add_data_sm.index.get_level_values('Consumption region')[i]
                        workbook['v']['C'+str(i+2)] = MI['a']
                        workbook['v']['D'+str(i+2)] = add_data_sm.index.get_level_values('Activity')[i]
                        workbook['v']['E'+str(i+2)] = 'Update'
                        workbook['v']['F'+str(i+2)] = add_data_sm.iloc[i,-1]
        
                elif m=='e':
                    for i in range(len(add_data_sm)):
                        workbook['e']['A'+str(i+2)] = add_data_sm.index.get_level_values('Factor of production')[i]
                        workbook['e']['B'+str(i+2)] = add_data_sm.index.get_level_values('Consumption region')[i]
                        workbook['e']['C'+str(i+2)] = MI['a']
                        workbook['e']['D'+str(i+2)] = add_data_sm.index.get_level_values('Activity')[i]
                        workbook['e']['E'+str(i+2)] = 'Update'
                        workbook['e']['F'+str(i+2)] = add_data_sm.iloc[i,-1]
                        
            workbook.save(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\\Shock_files\\{logic}_{year}_{tech}.xlsx")
            workbook.close()
        
#%% Implementing endogenization of capital
for scen in price_logics:
    for year in years:
        for tech in tech_performances:
            world[year].shock_calc(f"{pd.read_excel(paths, index_col=[0]).loc['Shocks',user]}\\Shock_files\\{logic}_{year}_{tech}.xlsx", z=True, v=True, scenario=f"{scen} - {year} - {tech}")

#%%
path_aggr  = r"Aggregations\Aggregation_postprocess.xlsx"
folder = f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\Footprints - Monetary units"

for year in years:
    world_aggr = world[year].aggregate(path_aggr, inplace=False, levels=["Activity","Commodity"])
    
    f = {}
    sat_accounts = [
        'Energy Carrier Supply: Total', 
        'CO2 - combustion - air', 
        'CH4 - combustion - air', 
        'N2O - combustion - air',
        ]
    
    commodities = [
            'PV plants',
            'PV modules',
            'Si cells',
            'Onshore wind plants',
            'DFIG generators',
            'Offshore wind plants',
            'PMG generators',
            'Electricity by wind',
            'Electricity by solar photovoltaic'        
       ]
    
    for a in sat_accounts:
        f[a] = {}
        if ":" in a:
            name = a.replace(':'," -")
        else:
            name = a
        for s in world_aggr.scenarios:
            e = world_aggr.get_data(matrices=['e'], scenarios=[s])[s][0].loc[a]
            w = world_aggr.get_data(matrices=['w'], scenarios=[s])[s][0]
            f[a][s] = np.diag(e) @ w
            f[a][s].index = f[a][s].columns
        
        for k,v in f[a].items():
            subfolder = f"{folder}\\{name}"
            if not os.path.exists(subfolder):
                os.mkdir(subfolder)
            v.to_csv(f"{subfolder}\\{k}.csv")

    f['GHGs'] = {}
    for s in world_aggr.scenarios:
        f['GHGs'][s] = f['CO2 - combustion - air'][s] + f['CH4 - combustion - air'][s]*26 + f['N2O - combustion - air'][s]*298
        subfolder = f"{folder}\\{'GHGs'}"
        if not os.path.exists(subfolder):
            os.mkdir(subfolder)
        f['GHGs'][s].to_csv(f"{subfolder}\\{k}.csv")

#%% Linkages
linkages = {}
linkages_df = pd.DataFrame()
for year in years:
    for scem in world[year].scenarios:
        db = mario.Database(
            Z = world[year].matrices[scem]['Z'],
            Y = world[year].matrices[scem]['Y'],
            E = world[year].matrices[scem]['E'],
            V = world[year].matrices[scem]['V'],
            EY =world[year].matrices[scem]['EY'],
            units = world[year].units,
            table='SUT',
            )
        db.to_iot(method='B')
        if scem == 'baseline':
            scen = 'Baseline'
            tech = 'Average'
        else:
            scen = scem.split(' - ')[0]
            tech = scem.split(' - ')[-1]

        linkages[f'{scen} - {year} - {tech}'] = db.calc_linkages(multi_mode=False)
        linkages[f'{scen} - {year} - {tech}'] = linkages[f'{scen} - {year} - {tech}'].droplevel(1)
        new_columns = pd.MultiIndex.from_arrays(
            [[i.split(" ")[0] for i in list(linkages[f'{scen} - {year} - {tech}'].columns)],
            [i.split(" ")[1] for i in list(linkages[f'{scen} - {year} - {tech}'].columns)]],
            )
        linkages[f'{scen} - {year} - {tech}'].columns = new_columns
        linkages[f'{scen} - {year} - {tech}'].columns.names = ['Scope',"Origin"]
        linkages[f'{scen} - {year} - {tech}'] = linkages[f'{scen} - {year} - {tech}'].stack(0)
        linkages[f'{scen} - {year} - {tech}']['Scenario'] = scen
        linkages[f'{scen} - {year} - {tech}']['Year'] = year
        linkages[f'{scen} - {year} - {tech}']['Performance'] = tech
        
        linkages_df = pd.concat([linkages_df, linkages[f'{scen} - {year} - {tech}']], axis=0)

linkages_df.reset_index(inplace=True)
linkages_df.to_csv(f"{pd.read_excel(paths, index_col=[0]).loc['Results',user]}\\Linkages.csv", index=False)
            
#%% Endogenization of capital database to excel
for year in years:
    for scen in price_logics:
        for tech in tech_performances:
            print(f"{scen} - {year} - {tech}")
            folder_name = f"{pd.read_excel(paths, index_col=[0]).loc['Database',user]}\\d. Shock - Endogenization of capital\\{scen} - {year} - {tech}"
            if not os.path.exists(folder_name):
                os.mkdir(folder_name)
            world[year].to_txt(folder_name, scenario=f"{scen} - {year} - {tech}", flows=False, coefficients=True)

#%%



