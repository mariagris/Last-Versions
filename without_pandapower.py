#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 19 10:27:10 2019

@author: mariagris
"""

"""
- Only demo grid, first 3 lines!!
- Import from the excel the data necessary for the line. 
- Create the line type. 
- Iterate over the 3 first lines in the excel to get the impedances (already multiplied by the length)
- Once we have the impedance (and R, Xc and Xl separately if needed) we can proceed to create
  the dataframe with all this information
- With the dataframe the matric can be done, both the impedance and admittance matrix
- From the admittance matrix we can obtain G and B

"""
import os
from pathlib import Path
import numpy as np
from numpy import matrix
import pandas as pd
import math
from numpy.linalg import inv

#dfload = pd.read_json(Path(str(os.getcwd()) +'/ofpfs_sent.json', orient='rows'))
file = 'Sitel_Invade_MV_Topology.xlsx'
xl = pd.ExcelFile(Path(str(os.getcwd()) + '/'+ file))

MVNetwork = xl.parse('Linies') 
x_ohm_per_km = MVNetwork['Reactancia_ohm_km'][0]
c_uf_per_km = MVNetwork['Capacitat_uF_km'][0]
r_ohm_per_km = MVNetwork['Resistencia_ohm_km'][0]

line_data = pd.DataFrame(np.array([[0.161, 0.113, 240]]),columns=['r_ohm_km', 'xl_ohm_km', 'c_uF_km'])

line = pd.DataFrame() 
for i in range(3):
    R=(MVNetwork['Length_km'][i])*(line_data['r_ohm_km'][0])
    X_l=(MVNetwork['Length_km'][i])*(line_data['xl_ohm_km'][0])
    X_c=1/(2*math.pi*50*(MVNetwork['Length_km'][i])*10e-6*(line_data['c_uF_km'][0]))
    X=X_l-X_c
    Z =R+(X_l-X_c)*1j
    line=line.append([[MVNetwork['Length_km'][i],R,X_l,X_c,X,Z]])

line.columns = ['l', "R_ohm", "Xl_ohm", "Xc_ohm", 'X_ohm', "Z_ohm"]
line = line.reset_index(drop=True)

### Inductance/Admittance matrix creation [Z], [Y], [G], [B]
########## Impedance matrix
Z = np.matrix( [ [line["Z_ohm"][0],(-line["Z_ohm"][0]),0,0],
                 [-line["Z_ohm"][0],line["Z_ohm"][0]+line["Z_ohm"][1],-line["Z_ohm"][1],0],
                 [0,-line["Z_ohm"][1],line["Z_ohm"][1]+line["Z_ohm"][2],-line["Z_ohm"][2]],
                 [0,0,-line["Z_ohm"][2],line["Z_ohm"][0]] ] )

########## Admittance matrix    
Y=inv(Z)  

########## [G], [B]
G=Y.real
B=(Y.imag)*1j




