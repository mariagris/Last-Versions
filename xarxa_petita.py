#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 12 09:44:07 2019

@author: mariagris
"""

"""
- Using pandapower to run de powerflow, a litle part of the grid has been created to observe its behaviour
- No iteration over the Json file of loads has been made
- Just first iteration and created by hand in order to see all steps
- 1st: create the empty network
- 2nd: import the data from the jason (loads) and the excel (grid)
- 3rd: create line std type
- 4th: crate MV busses and LV busses after trafos
- 5th: create tranformes between two already existing busses
- 6th: create lines between MV busses
- 7th: create the external grid (slack bus) in an already existing bus (as radial grid, the first one)
- 8th: create the loads for the LV busses (invented, as data from json doesn't have any load 
       different to zero in the LV busses taken for the demo in this first iteration)
- 9th: run the net
- 10th: plotting the results

"""

import pandas as pd 
#import matplotlib.pyplot as plt 
import os 
from pathlib import Path  
import pandapower as pp 
#import igraph
from pandapower import plotting
import numpy as np
from numpy import matrix

net = pp.create_empty_network() 

##############################################################################
############################ IMPORTAR JSON + EXCEL ###########################

dfload = pd.read_json(Path(str(os.getcwd()) +'/ofpfs_sent.json', orient='rows'))
file = 'Sitel_Invade_MV_Topology.xlsx'
xl = pd.ExcelFile(Path(str(os.getcwd()) + '/'+ file))

MVNetwork = xl.parse('Linies') 
x_ohm_per_km = MVNetwork['Reactancia_ohm_km'][0]
c_uf_per_km = MVNetwork['Capacitat_uF_km'][0]
r_ohm_per_km = MVNetwork['Resistencia_ohm_km'][0]

line_data = {"c_nf_per_km": c_uf_per_km*1000, "r_ohm_per_km": r_ohm_per_km, "x_ohm_per_km": x_ohm_per_km, "max_i_ka":0.415}
pp.create_std_type(net, line_data, "line_ESTABANELL", element='line')

MVNetworkbusses = xl.parse('Busses')
for i in range(4):
    pp.create_bus(net, vn_kv=20.5,name=MVNetworkbusses['name'][i], max_vm_pu=1.1, min_vm_pu=0.9)

MVNetworkbussesTrafos = xl.parse('Trafos')
for i in range(3):  #(len(MVNetworkbussesTrafos['name'])):
    pp.create_bus(net, vn_kv= MVNetworkbussesTrafos['vn_lv_kv'][i], name=MVNetworkbussesTrafos['name'][i], max_vm_pu=1.1, min_vm_pu=0.9)

pp.create_transformer_from_parameters(net, hv_bus=1, lv_bus=4, sn_mva=0.4, vn_hv_kv=20.5, vn_lv_kv=0.23, vk_percent=4., vkr_percent=0.04, pfe_kw=0.75, i0_percent=1.8, shift_degree=150, in_service=True, parallel=1, name='E.T. NODE 1', tap_side='hv', tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=False, df=1, index=75)    
pp.create_transformer_from_parameters(net, hv_bus=1, lv_bus=5, sn_mva=0.63, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=4.09, vkr_percent=0.0409, pfe_kw=0.548, i0_percent=0.9, shift_degree=150, in_service=True, parallel=1, name='E.T. NODE 2', tap_side='hv',tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True, df=1, index=76)
pp.create_transformer_from_parameters(net, hv_bus=3, lv_bus=6, sn_mva=1, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=6., vkr_percent=0.08, pfe_kw=1.4, i0_percent=1.3, shift_degree=150, in_service=True, parallel=1, name='E.T.CUARTEL', tap_side='hv', tap_pos=0,tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True,df=1, index=22)

pp.create_line(net, from_bus=0, to_bus=1, length_km=MVNetwork['Length_km'][0], std_type="line_ESTABANELL", name='line 1')
pp.create_line(net, from_bus=1, to_bus=2, length_km=MVNetwork['Length_km'][1], std_type="line_ESTABANELL", name='line 2')
pp.create_line(net, from_bus=2, to_bus=3, length_km=MVNetwork['Length_km'][2], std_type="line_ESTABANELL", name='line 3')
   
pp.create_ext_grid(net, bus=0, vm_pu=1.0, va_degree=0.0, in_service=True, name="Grid Connection")#, max_p_mw=2000000, min_p_mw=0, max_q_mvar=38000, min_q_mvar=-20000)
 
pp.create_load(net, bus=4, p_mw = (dfload['output'][0]['input'][0]['ActivePower'])/1000000, q_mvar = (dfload['output'][0]['input'][i]['ReactivePower'])/1000000)
pp.create_load(net, bus=5, p_mw = (dfload['output'][0]['input'][1]['ActivePower'])/1000000, q_mvar = (dfload['output'][1]['input'][i]['ReactivePower'])/1000000)
pp.create_load(net, bus=6, p_mw = (dfload['output'][0]['input'][2]['ActivePower'])/1000000, q_mvar = (dfload['output'][2]['input'][i]['ReactivePower'])/1000000)

pp.runpp(net, algorithm='bfsw')
pp.plotting.simple_plot(net)


