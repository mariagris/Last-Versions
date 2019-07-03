#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jun  5 12:44:24 2019

@author: mariagris
"""

"""
Complete grid!
Saves in a DataFrame all the overloaded lines IN COLUMNS:
    - Each timestamp is a new column
    - For each timestamp (column) the Id of the line and the overload percentage is saves
    - If there is no overload, "no congestion" message is saved
Saves in a DataFrame all the loads also IN COLUMNS for each timestamp:
    - In such way, is easier to compare both (overloaded and loads) dataframes
    - If there is one than one load referred to a bus, both are saved in the same bus but separately,
      so that the two numbers can be seen.
    - If no load in current bus, a 0 is saved    
Global variables are created for each iteration nets:
    - they are callable and manipulable.
    - Elements inside net elements are DataFrames!!! very useful
Also an excel file is created with the information of overloaded lines.
"""

import pandas as pd 
#import matplotlib.pyplot as plt 
import os 
from pathlib import Path  
import pandapower as pp 
from pandapower import plotting
#import igraph
dfload = pd.read_json(Path(str(os.getcwd()) +'/ofpfs_sent.json', orient='rows'))

def grid():
    net = pp.create_empty_network()  
    file = 'Sitel_Invade_MV_Topology.xlsx' 
    xl = pd.ExcelFile(Path(str(os.getcwd()) + '/'+ file))  
    MVNetwork = xl.parse('Linies') 
    x_ohm_per_km = MVNetwork['Reactancia_ohm_km'][0]
    c_uf_per_km = MVNetwork['Capacitat_uF_km'][0]
    r_ohm_per_km = MVNetwork['Resistencia_ohm_km'][0]

#IMPORTANT: no change in the capacitance units has been made. In pandapower documentation it is 
#a the magnitude around 240, what we already have. if multiplied by 1000 to change it from micro 
#to nano farads, overload is of 300%. Make sure with Estabanell they are giving us micro. Meanwhile,
#it is left as the value of micro as it is the most similar of a reasonable magnitude
    
    line_data = {"c_nf_per_km": c_uf_per_km , "r_ohm_per_km": r_ohm_per_km, "x_ohm_per_km": x_ohm_per_km, "max_i_ka":0.415}
    pp.create_std_type(net, line_data, "line_ESTABANELL", element='line')

    MVNetworkbusses = xl.parse('Busses')
    for i in MVNetworkbusses['name']:
        pp.create_bus(net, vn_kv=20.5,name=i, max_vm_pu=1.1, min_vm_pu=0.9)

    MVNetworkbussesTrafos = xl.parse('Trafos')
    for i in range(len(MVNetworkbussesTrafos['name'])):
        pp.create_bus(net, vn_kv= MVNetworkbussesTrafos['vn_lv_kv'][i], name=MVNetworkbussesTrafos['name'][i], max_vm_pu=1.1, min_vm_pu=0.9)

    ## per aconseguir els Id's tant dels empalmaments com els nusos a la banda LV dels trafos:
    l=MVNetworkbusses['Id Bus']                 #dataframe
    r=l.append(MVNetworkbussesTrafos['Id Bus']) #r still a DataFrame
    r=r.values.tolist()                         #r list
    net.bus.insert(5, 'Id Bus', r)  
    net.bus = net.bus.set_index("Id Bus", drop = False)


    pp.create_transformer_from_parameters(net, hv_bus=202, lv_bus=75, sn_mva=0.4, vn_hv_kv=20.5, vn_lv_kv=0.23, vk_percent=4., vkr_percent=0.04, pfe_kw=0.75, i0_percent=1.8, shift_degree=150, in_service=True, parallel=1, name='E.T. NODE 1', tap_side='hv', tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=False, df=1, index=75)    
    pp.create_transformer_from_parameters(net, hv_bus=202, lv_bus=76, sn_mva=0.63, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=4.09, vkr_percent=0.0409, pfe_kw=0.548, i0_percent=0.9, shift_degree=150, in_service=True, parallel=1, name='E.T. NODE 2', tap_side='hv',tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True, df=1, index=76)
    pp.create_transformer_from_parameters(net, hv_bus=204, lv_bus=22, sn_mva=1, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=6., vkr_percent=0.08, pfe_kw=1.4, i0_percent=1.3, shift_degree=150, in_service=True, parallel=1, name='E.T.CUARTEL', tap_side='hv', tap_pos=0,tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True,df=1, index=22)
    pp.create_transformer_from_parameters(net, hv_bus=209, lv_bus=26, sn_mva=0.63, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=4., vkr_percent=0.053, pfe_kw=1.03, i0_percent=1.6, shift_degree=150, in_service=True, parallel=1, name='E.T.PEDRALS',  tap_side='hv',tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True,df=1, index=26) 
    pp.create_transformer_from_parameters(net, hv_bus=213, lv_bus=30, sn_mva=0.8, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=6., vkr_percent=0.06, pfe_kw=1.4, i0_percent=1.3, shift_degree=150, in_service=True, parallel=1, name='E.T. LA MUTUA',tap_side='hv', tap_pos=0,tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True, df=1,index=30)
    pp.create_transformer_from_parameters(net, hv_bus=219, lv_bus=34, sn_mva=0.63, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=4.28, vkr_percent=0.0428, pfe_kw=1.3, i0_percent=1.05, shift_degree=150, in_service=True, parallel=1, name='E.T. LA LLEÓ', tap_side='hv',tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True,df=1, index=34)
    pp.create_transformer_from_parameters(net, hv_bus=220, lv_bus=64, sn_mva=0.63, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=4., vkr_percent=0.04, pfe_kw=1.3, i0_percent=1.6, shift_degree=150, in_service=True, parallel=1, name='E.T. PRADES 1',tap_side='hv',tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True,df=1, index=64)
    pp.create_transformer_from_parameters(net, hv_bus=228, lv_bus=42, sn_mva=0.4, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=4.05, vkr_percent=0.054, pfe_kw=0.768, i0_percent=1.72, shift_degree=150, in_service=True, parallel=1, name='E.T. ECUADOR', tap_side='hv', tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True,df=1, index=42)
    pp.create_transformer_from_parameters(net, hv_bus=230, lv_bus=46, sn_mva=0.4, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=5., vkr_percent=0.066, pfe_kw=1.193, i0_percent=0.45, shift_degree=150, in_service=True, parallel=1, name='E.T. URUGUAI', tap_side='hv',tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True, df=1, index=46)
    pp.create_transformer_from_parameters(net, hv_bus=232, lv_bus=50, sn_mva=0.63, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=4., vkr_percent=0.053, pfe_kw=1.3, i0_percent=1.6, shift_degree=150, in_service=True, parallel=1, name='E.T. LA TORRETA', tap_side='hv',tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2,tap_step_percent=2.5,tap_step_degree=0,tap_phase_shifter=True,df=1, index=50)
    pp.create_transformer_from_parameters(net, hv_bus=234, lv_bus=54, sn_mva=0.4, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=4., vkr_percent=0.04, pfe_kw=0.93, i0_percent=1.8, shift_degree=150, in_service=True, parallel=1,name='E.T. COSTA BRAVA 1', tap_side='hv',tap_pos=0, tap_neutral=0, tap_min=-2,tap_max=2,tap_step_percent=2.5,tap_step_degree=0,tap_phase_shifter=True, df=1, index=54)
    pp.create_transformer_from_parameters(net, hv_bus=236, lv_bus=58, sn_mva=0.4, vn_hv_kv=20.5, vn_lv_kv=0.23, vk_percent=4., vkr_percent=0.04, pfe_kw=0.93, i0_percent=1.8, shift_degree=150,in_service=True,parallel=1,name='E.T. NOVA VERDAGUER 1',tap_side='hv',tap_pos=0, tap_neutral=0,tap_min=-2,tap_max=2,tap_step_percent=2.5,tap_step_degree=0,tap_phase_shifter=True, df=1, index=58)
    pp.create_transformer_from_parameters(net, hv_bus=236, lv_bus=66, sn_mva=0.4, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=4., vkr_percent=0.04, pfe_kw=0.75, i0_percent=1.8, shift_degree=150, in_service=True, parallel=1,name='E.T. NOVA VERDAGUER 2',tap_side='hv',tap_pos=0, tap_neutral=0,tap_min=-2,tap_max=2,tap_step_percent=2.5,tap_step_degree=0,tap_phase_shifter=True, df=1,index=66)
    pp.create_transformer_from_parameters(net, hv_bus=238, lv_bus=61, sn_mva=0.25, vn_hv_kv=20.5, vn_lv_kv=0.4, vk_percent=4., vkr_percent=0.04, pfe_kw=0.65, i0_percent=2, shift_degree=150, in_service=True, parallel=1, name='E.T. GRANADA', tap_side='hv', tap_pos=0, tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tap_step_degree=0, tap_phase_shifter=True, df=1, index=61)

    for i in range(len(MVNetwork['Id TedisNet'])): 
        pp.create_line(net, from_bus=MVNetworkbusses['Id Bus'][i], to_bus=MVNetworkbusses['Id Bus'][i+1], length_km=MVNetwork['Length_km'][i], std_type="line_ESTABANELL", name=MVNetwork['origen'][i])
    net.line.insert(14, 'Id linia', MVNetwork['Id TedisNet'])

    #the slack bus will act as an external grid
    pp.create_ext_grid(net, bus=201, vm_pu=1.0, va_degree=0.0, in_service=True, name="Grid Connection")

    return(net)



dataframe_lines_overloaded=pd.DataFrame()
dataframe_loads=pd.DataFrame()

q=0
for i in dfload['output']:
    power=i['input']
    net=grid()
    l=[]
    for e in power:
        pp.create_load(net, bus=e['IdTedisNet'], p_mw = (e['ActivePower'])/1000000, q_mvar = (e['ReactivePower'])/1000000)
        l.append(e['IdTedisNet']) 
    for n in net.bus['Id Bus']:
        if n not in l:
            pp.create_load(net, bus=n, p_mw=0, q_mvar=0)
    
    pp.runpp(net, algorithm='bfsw')
    pp.plotting.simple_plot(net)

        
    ##########################################################################
    ######### GUARDEM EN UN DATAFRAME TOTES LES OVERLOADED ###################    
    
    lines_overloaded=[]
    for t in range(len(net.res_line['loading_percent'])):
        if net.res_line['loading_percent'][t]>3:  
            lines_overloaded.append((net.line['Id linia'][t],net.res_line['loading_percent'][t]))
        else:
            lines_overloaded.append('no congestio')
        
    dataframe_lines_overloaded.insert(q,i['timestamp'],lines_overloaded)
    
    ##########################################################################
    ######### GUARDEM EN UN DATAFRAME TOTES LES LOADS  #######################    
    net.load = net.load.set_index('bus', drop = False)
    results_loads=[]
    if q==0:
        dataframe_loads.insert(0,'Id Bus',net.bus['Id Bus'])
        dataframe_loads = dataframe_loads.set_index('Id Bus', drop = False)
    dataframe_loads.insert(q,i['timestamp'],'')
    for e in dataframe_loads['Id Bus']: #net.load['p_mw']:
        dataframe_loads[i['timestamp']][e]=net.load['p_mw'][e]
        
    ########################################################################## 
        
    q=q+1
    print('Data i hora actuals:',pd.datetime.now())
    print('Timestamp:',i['timestamp'])
    globals()['net{}'.format(q)] = net
    
#### FI DE LA ITERACIÓ
    
    
dataframe_lines_overloaded.to_excel(Path(str(os.getcwd()) +'/PowerFlow.xlsx'))    

#We want to try and call a global variable to see if it works.    
prova=net2.res_line



