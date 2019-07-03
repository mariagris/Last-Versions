# -*- coding: utf-8 -*-
"""
Created on Mon Mar 11 12:08:57 2019

@author: Ingrid - CITCEA
"""

import pyomo.opt
from pyomo.environ import *
import pandas as pd
import xlwings as xw
from sys import argv


"""  
# Edit this line to select input file 
file = 'Sitel_Invade_MV_Topology.xlsx' 
 
##READ THE INPUTS FROM EXCEL## 
 
## Parsing input file 
  
xlsFile = pd.ExcelFile(Path(str(os.getcwd()) + '/'+ file)) 
xl = pd.ExcelFile(xlsFile) 
     
##Parse specified sheet(s) into a DataFrame. Equivalent to read_excel 
## Llegeixo cada pestanya, dins de xl PANDAS  
MVNetwork = xl.parse('Linies') 
 
## index linies 
#init_MVlines = MVNetwork.index 
 
#Canvio l'index inicial de 0,1,2... per l'index de cada línia Id TedisNet  
MVNetwork2 = MVNetwork.set_index("Id TedisNet") 

"""

###############################################################################
########################## OPTIMIZATION MODEL #################################
###############################################################################

##READ THE INPUTS FROM EXCEL##

# Edit this line to select input file
file = 'ED_Data.xlsx'

## Parsing input file
## Llegeixo l'excel sencer 
xl = pd.ExcelFile(file)
    
##Parse specified sheet(s) into a DataFrame. Equivalent to read_excel
## Llegeixo cada pestanya, dins de xl PANDAS 
generator_data = xl.parse('generator_data')
windfarm_data=xl.parse('windfarm_data')
line_data=xl.parse('line_data')
bus_data= xl.parse('bus_data')


##indexes
init_g = generator_data.index
init_w = windfarm_data.index
init_l = line_data.index
init_b = bus_data.index


### Transformation of each column into dict.

#init_dso_down = dso_request['down'].to_dict()
init_gmin = generator_data['gmin'].to_dict() #minimum power output of generators
init_gmax = generator_data['gmax'].to_dict() #maximum power output of generators
init_cg = generator_data['cg'].to_dict() #incremental costs of generators
init_cg0 = generator_data['cg0'].to_dict() #fixed costs of generators

init_wforecast = windfarm_data['wforecast'].to_dict() #wind forecast
init_cw = windfarm_data['cw'].to_dict() #incremental cost of wind generators

init_fmax = line_data['fmax'].to_dict()
init_x = line_data['x'].to_dict()

init_d = bus_data['d'].to_dict()


### MODEL DEFINITION 

#defineixo el ConcreteModel o AbstractModel aquí. Si vull hi puc posar un nom dins 
model = ConcreteModel(name="(ED)")


 
#########  SETS  ##########

#model.t=Set(initialize=init_t)

model.G = Set(initialize=init_g)   #generators index set
model.W = Set(initialize=init_w)   #Wind farms placed 
model.L = Set(initialize=init_l)   #network lines 
model.B = Set(initialize=init_b)   #network buses  

### Define parameters 


# Generator parameters
#model.p_gen_r = Param(model.g_r, initialize=init_p_gen_r, within=NonNegativeReals)
#model.w_gen_r = Param(model.t*model.g_r, initialize=init_w_gen_r, within=NonNegativeReals)

# Regulation DOWN request
#model.dso_down = Param(model.t, initialize=init_dso_down, within=NonNegativeReals)

model.g_max = Param(model.G, initialize= init_gmax, within=NonNegativeReals) #maximum power output of generators
model.g_min = Param(model.G, initialize= init_gmin, within=NonNegativeReals) #minimum power output of generators
model.c_g = Param(model.G, initialize= init_cg, within=NonNegativeReals) #incremental cost of generator
model.c_g0 = Param(model.G, initialize= init_cg0, within=NonNegativeReals) #fixed cost of generator
model.c_w = Param(model.W, initialize= init_cw, within=NonNegativeReals) #incremental cost of wind generator
model.d = Param(model.B, initialize= init_d, within=NonNegativeReals) #total demand 
model.w_f = Param(model.W, initialize=init_wforecast, within=NonNegativeReals) #wind forecast 


###decision variables
model.gi = Var(model.G, within = NonNegativeReals) #amount of electricity produced by the generator g
model.wi = Var(model.W, within = NonNegativeReals)    #amount of electricity produced by the windfarm


## OBJECTIVE FUNCTION

def obj_rule(model): 
 return sum(model.c_g[i]*model.gi[i]+model.c_w[w]*model.wi[w] for i in model.G for w in model.W)

model.Objective = Objective(rule=obj_rule, sense = minimize)


"""def function_objetive(model):
    return sum(sum(model.p_gen_r[g]*(model.w_gen_r[t,g]-model.psi[t,g]) for t in model.t) for g in model.g_r)
model.objetivo=Objective(rule=function_objetive, sense=minimize)"""
   
### CONSTRAINTS 

"""Important constraints: 
    Quan defineixo la FUNCIÓ constraint importo tot el model sencer i l'index sobre el qual tinc la restricció 
    Després he de definir la constraint com a tal i allà hi importo els SETS i la funció que he definit com a rule.
       
    """

#Generator i should be between the minimum and the maximum 
def limit_generator(model, i): 
    return model.gi[i] <= model.g_max[i] and model.gi[i] >= model.g_min[i] 

model.limit_generatorc1 = Constraint(model.G, rule = limit_generator)

#windfarm bounds 
def limit_windfarm(model,w):
    return model.wi[w] <= model.w_f[w] and model.wi[w]>=0.0

model.limit_windfarmc2 = Constraint(model.W, rule= limit_windfarm)

# power balance constraint

def power_balance(model): 
   return sum((model.gi[i] + model.wi[f]) for i in model.G for f in model.W) == sum((model.d[b]) for b in model.B)
model.power_balancec3 = Constraint(model.G, model.W, model.B, rule = power_balance)

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


gi = pd.DataFrame(index= generator_data.index, columns = ['Generation']) #Creo un dataframe per cadascuna de les variables de decisió que té el meu model. 

wi = pd.DataFrame(index = windfarm_data.index, columns = ['Generation']) #Creo un dataframe per emmagatzemar els resultats de la variable de decisió wi. 

for gen in generator_data.index: # Per cadascun dels generadors (dels seus índex g1 i g2) itero per llegir el valor
    gi['Generation'][gen] = instance.gi[gen].value  # llegeixo el valor de l'instance. és important aquí recordar que emmagatzemo el valor de la seguent manera 
#   nomvariabledataframe[COLUMNA][FILA] = instance.VARIABLEDECISIO[INDEX].value
    
for wind in windfarm_data.index: 
    wi['Generation'][wind] = instance.wi[wind].value


print(gi) 
print(wi)

#psi = pd.DataFrame(index=dso_request.index, columns=p_gen_r.index)
#
#for gen in p_gen_r.index:
#    for period in dso_request.index:
#        psi[gen][period] = instance.psi[period, gen].value
#
#print(psi)
