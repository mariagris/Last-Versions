#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Mar  6 11:02:55 2019

@author: mariagris
"""
"""
- First attempt of pandapower net
- 4 buses, 1 external grid (slack bus, grid connection), meshed grid, 4 lines, as a matlab example we had
- 4 loads
- line type created for each line
- plot of the net
"""

import pandapower as pp
import pandapower.networks as pn


#create empty net
net = pp.create_empty_network() 
#f_hz=50.0, sn_kva=100000
#create buses
bus1 = pp.create_bus(net, vn_kv=230., name="Bus 1") # type = "SLACK")
bus2 = pp.create_bus(net, vn_kv=230., name="Bus 2") #type = "PQ")
bus3 = pp.create_bus(net, vn_kv=230., name="Bus 3") #type = "PQ")
bus4 = pp.create_bus(net, vn_kv=230., name="Bus 4") #type = "PV")

#create bus elements

#pp.create_gen(net, bus=bus1, vm_pu = 1, name="Generator_1", p_kw = 0, in_service=True)
pp.create_ext_grid(net, bus=bus1, vm_pu=1, va_degree=0, in_service=True, name="Grid Connection")

pp.create_load(net, bus=bus1, p_mw = 50000, q_mvar = 30987.216920155117, name="Load_1")

pp.create_load(net, bus=bus2, p_mw = 170000, q_mvar = 105356.53752852738, name="Load_2")

pp.create_load(net, bus=bus3, p_mw = 200000, q_mvar = 123948.86768062047, name="Load_3")

pp.create_gen(net, bus=bus4, vm_pu = 1.02, name="Generator_2", p_mw = -318000, in_service=True)
pp.create_load(net, bus=bus4, p_mw = 80000, q_mvar = 49579.54707224818, name="Load_4")



#create lines

line_data_12 = {"c_nf_per_km": 616.76, "r_ohm_per_km": 5.33232, "x_ohm_per_km": 26.6616, "max_i_ka":0.415}
pp.create_std_type(net, line_data_12, "line12", element='line')
pp.create_line(net, from_bus=bus1, to_bus=bus2, length_km=1, std_type="line12",  name="line_12")


line_data_13 = {"c_nf_per_km": 466.63, "r_ohm_per_km": 3.93576, "x_ohm_per_km": 19.6788, "max_i_ka":0.415}
pp.create_std_type(net, line_data_13, "line13", element='line')
pp.create_line(net, from_bus=bus1, to_bus=bus3, length_km=1, std_type="line13",  name="line_13")

line_data_24 = {"c_nf_per_km": 466.63, "r_ohm_per_km": 3.93576, "x_ohm_per_km": 19.6788, "max_i_ka":0.415}
pp.create_std_type(net, line_data_24, "line24", element='line')
pp.create_line(net, from_bus=bus2, to_bus=bus4, length_km=1, std_type="line24",  name="line_24")

line_data_34 = {"c_nf_per_km": 767.19, "r_ohm_per_km": 6.7288, "x_ohm_per_km": 33.6444, "max_i_ka":0.415}
pp.create_std_type(net, line_data_34, "line34", element='line')
pp.create_line(net, from_bus=bus3, to_bus=bus4, length_km=1, std_type="line34",  name="line_34")

 
pp.plotting.simple_plot(net)

#create branch elements
#no tenim trafos en aquest exemple



