# -*- coding: utf-8 -*-
"""
Created on Wed Apr  3 10:01:42 2019

@author: Ingrid - CITCEA
"""

###############################################################################
##########################INFORMATION#########################################

#OPF Model, formulated as a Relaxed Branch Flow Model (MISOCP)
#ingrid.munne@citcea.upc.edu
###############################################################################

import pyomo.opt
from pyomo.environ import *
import pandas as pd
import xlwings as xw
from sys import argv

###############################################################################
########################## OPTIMIZATION MODEL #################################
###############################################################################

##READ THE INPUTS FROM EXCEL##

# Edit this line to select input file
file = 'OPF_Basic_IMC.xlsx'

## Parsing input file
## Llegeixo l'excel sencer 
xl = pd.ExcelFile(file)
    
##Parse specified sheet(s) into a DataFrame. Equivalent to read_excel
## Llegeixo cada pestanya, dins de xl PANDAS 
param_data = xl.parse('parameter')
bus_data = xl.parse('bus')
line_data = xl.parse('line')
demand_data = xl.parse('demand')


generation_data = xl.parse('generation')

##indexes
init_l = line_data.index #lines index
init_bn = bus_data.index  #Total index nodes 
init_bfrom = line_data.index #Total number of from buses
init_bto = line_data.index #Total number of TO buses 
init_gen = generation_data.index #total number of generation

#tots els params a excel els posem sense espai entre nom i unitats. 
#Pg [MW] NO!! Pg[MW] SI! 

### param initialization vectors 
init_Sbase = param_data.iloc[0]['Sbase_kVA'] #Sbase value for pu calculation
init_Vbase = param_data.iloc[0]['Vbase_kV'] # Voltage base value for pu calculation
init_Vslack = param_data.iloc[0]['Vslack_pu'] #pu voltage for slack node
init_Zbase = param_data.iloc[0]['Zbase_ohms'] #admitance base value in OHMS 

#### params of bus_data 
#Per determinar un dictionary quan tinc més d'una key ho faig fent dos fors per crear un dictionary de dictionary amb les mateixes keys.
#init_bus_data = {(bus,char): bus_data[char][bus] for bus in bus_data.index for char in bus_data.columns}
init_bus_data = { i: {'Type': bus_data['Type'][i], 'Vnom_pu': bus_data['Vnom_pu'][i], 'Vangle': bus_data['Vangle'][i],
                           'Vmax_pu': bus_data['Vmax_pu'][i], 'Vmin_pu': bus_data['Vmin_pu'][i]} for i in bus_data['BusID']}

init_Btype = bus_data['Type'].to_dict()  #(SLACK=1, PV=2(not implemented), PQ=3)

#### params of line_data
#init_line_data = {(line,char): line_data[char][line] for line in line_data.index for char in line_data.columns}
init_line_data = {i:{'Nodefrom': line_data['Nodefrom'][i], 'Nodeto': line_data['Nodeto'][i], 'r_pu': line_data['r_pu'][i],
                     'x_pu': line_data['x_pu'][i], 'b_pu': line_data['b_pu'][i], 'length_km': line_data['length_km'][i],
                     'Imax_pu': line_data['Imax_pu'][i]} for i in line_data['LineID']}

init_Pg = generation_data['Pg_MW'].to_dict() #active power value generation at node i 
init_Qg = generation_data['Qg_Mvar'].to_dict() #reactive power generation at node i 

init_Pd = demand_data['Pd_MW'].to_dict()  #Active power forecast demand at node i 
init_Qd = demand_data['Qd_Mvar'].to_dict() #Reactive power forecast demand at node i 

                                 

########################################################## MODEL DEFINITION 

#defineixo el ConcreteModel o AbstractModel aquí. Si vull hi puc posar un nom dins 
model = ConcreteModel()

#########  SETS  ##########

model.Bn = Set(initialize=init_bn) #buses in the network   
model.Bf = Set(initialize=init_bfrom) # FROM buses 
model.Bt = Set(initialize=init_bto)   #TO buses  
model.L = Set(initialize=init_l)   #network lines
model.G = Set(initialize= init_gen) #generators  

### Define parameters 

#Network characteristics
model.S_base = Param(initialize = init_Sbase, within=NonNegativeReals)
model.V_base = Param(initialize = init_Vbase, within=NonNegativeReals)
model.V_slack = Param(initialize = init_Vslack, within=NonNegativeReals)
model.Z_base = Param(initialize = init_Zbase, within=NonNegativeReals)

#Demand values
model.Pd = Param(model.Bn, initialize = init_Pd, within=NonNegativeReals)
model.Qd = Param(model.Bn, initialize = init_Qd, within=NonNegativeReals)

#Generation values

model.Pg = Param(model.Bn, initialize = init_Pg, within=NonNegativeReals)
model.Qg = Param(model.Bn, initialize = init_Qg, within=NonNegativeReals)

#Bus Type
model.Btype= Param(model.Bn, initialize = init_Btype, within = NonNegativeReals) 

#Line parameters 
model.r = Param(model.L, initialize = init_line_data[]['r_pu'], within=NonNegativeReals)
model.x = Param(model.L, initialize = XXXXXXXXXX, within=NonNegativeReals)
model.b = Param(model.L, initialize = XXXXXXXXXX, within=NonNegativeReals)

"""
# Flexibility limits parameters
model.Pgmax = Param(model.Bn, )
model.Pgmin =
model.Qgmax =
model.Qgmin = 

"""



###decision variables
model.fi_act = Var(model.Bn, within = NonNegativeReals) #amount of electricity produced by the generator g


## OBJECTIVE FUNCTION

## We define the OF to minimize the total costs of flexibility activation 
def obj_rule(model): 
 return sum(sum(model.c_f[t] * model.fi_act[i,t] for i in model.Bn) for t in model.T)

model.Objective = Objective(rule=obj_rule, sense = minimize)

   
### CONSTRAINTS 

"""Important constraints: 
    Quan defineixo la FUNCIÓ constraint importo tot el model sencer i l'index sobre el qual tinc la restricció 
    Després he de definir la constraint com a tal i allà hi importo els SETS i la funció que he definit com a rule
    """
    
### Active power bounds by the flexibility source 
    
def P_bounds_Flex_Source(model, i, t): 
    return model.fi_act[i,t] <= model.Pflex




#
##PV can not reduce more energy than Forecast prediction
#def limit_generation_forecast_red(model,t,g):
#    return model.psi[t,g] <= model.w_gen_r[t,g]
#model.constraint1=Constraint(model.t,model.g_r,rule=limit_generation_forecast_red)


#def limit_generator(model, i): 
#    return model.gi[i] <= model.g_max[i] and model.gi[i] >= model.g_min[i] 
#
#model.limit_generatorc1 = Constraint(model.G, rule = limit_generator)
#
##windfarm bounds 
#def limit_windfarm(model,w):
#    return model.wi[w] <= model.w_f[w] and model.wi[w]>=0.0
#
#model.limit_windfarmc2 = Constraint(model.W, rule= limit_windfarm)
#
## power balance constraint
#
#def power_balance(model): 
#   return sum((model.gi[i] + model.wi[f]) for i in model.G for f in model.W) == sum((model.d[b]) for b in model.B)
#model.power_balancec3 = Constraint(model.G, model.W, model.B, rule = power_balance)

######### RUNNING MODEL #########
 
""" 
Quan vull córrer el model necessito primer de tot crear una instancia. 
Després he de cridar el solver, en aquest cas glpk
després marco el gap entre iteracions. Això em pot anar bé més endavant per problemes complexos. 
Després escric a l'instance el fitxer junk.txt o junk.lp que es un fitxer amb totes les equacions que fa servir pyomo per
resoldre el model. Això és interessant quan em surti un infeasible per veure un llistat amb totes les equacions. 
Després genero el fitxer de resultats. PERÒ aquest fitxer no es mostra enlloc, les variables de resultats es desen en variables
internes que no podem mostrar tal qual. 

Aquesta és l'estructura general de com resoldre un problema d'optimització, de com cridar el solver. 
"""

instance = model.create_instance()
optimizer = pyomo.opt.SolverFactory('glpk') # Call the Solver (tested): glpk, gurobi
optimizer.options['mipgap'] = 0.005
instance.preprocess()
#junk es el .txt o .lp amb totes les equacions que fa servir per fer servir el model. 
# es interessant per quan ens surti un infeasible. 
instance.write('junk.lp', io_options={'symbolic_solver_labels': True})
results = optimizer.solve(instance, tee=True, report_timing=True) # , tee=True or False




##### CREATING OUTPUTS ###


"""
Per poder veure els valors que tenen les variables en el punt òptim he de seguir els següents passos
"""


#gi = pd.DataFrame(index= generator_data.index, columns = ['Generation']) #Creo un dataframe per cadascuna de les variables de decisió que té el meu model. 
#
#wi = pd.DataFrame(index = windfarm_data.index, columns = ['Generation']) #Creo un dataframe per emmagatzemar els resultats de la variable de decisió wi. 
#
#for gen in generator_data.index: # Per cadascun dels generadors (dels seus índex g1 i g2) itero per llegir el valor
#    gi['Generation'][gen] = instance.gi[gen].value  # llegeixo el valor de l'instance. és important aquí recordar que emmagatzemo el valor de la seguent manera 
##   nomvariabledataframe[COLUMNA][FILA] = instance.VARIABLEDECISIO[INDEX].value
#    
#for wind in windfarm_data.index: 
#    wi['Generation'][wind] = instance.wi[wind].value
#
#
#print(gi) 
#print(wi)

#psi = pd.DataFrame(index=dso_request.index, columns=p_gen_r.index)
#
#for gen in p_gen_r.index:
#    for period in dso_request.index:
#        psi[gen][period] = instance.psi[period, gen].value
#
#print(psi)
