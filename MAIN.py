#%% -*- coding: utf-8 -*-
"""
Created on Tue May 04 2023

@authors: 
    Lorenzo Rinaldi, Department of Energy, Politecnico di Milano
    Nicolò Golinucci, PhD, Department of Energy, Politecnico di Milano
    Emanuele Mainardi, Department of Energy, Politecnico di Milano
    Prof. Matteo Vincenzo Rocco, PhD, Department of Energy, Politecnico di Milano
    Prof. Emanuela Colombo, PhD, Department of Energy, Politecnico di Milano
"""

"""
Run this script to run the whole case study from scratch to results. Be sure 'years' is always the same list in all the scripts
""" 

import subprocess

# List of scripts to run in order
scripts = ['1. Database building.py', '2. Add sectors.py', '3. Case study.py']

# Loop through each script and run it
for script in scripts:
    subprocess.run(['python', script])

