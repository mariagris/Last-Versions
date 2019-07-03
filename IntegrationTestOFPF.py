# -*- coding: utf-8 -*-
"""
Created on Thu May  9 09:40:05 2019
@author: Ingrid Munné - CITCEA UPC / Maria Gris - CITCEA UPC

"""
##############################################################################
##################### libraries import########################################
##############################################################################

import json 
import pandapower as pp 
import pandapower.networks as pn
from pandapower import plotting 
#import igraph 
import pandas as pd
import xlwings as xw
from sys import argv
#import xlwings as xw
from sys import argv
from pathlib import Path
import os 
import matplotlib.pyplot as plt 
import numpy as np


##############################################################################
########################### 1. json import ###################################
##############################################################################

df1 = pd.read_json(Path(str(os.getcwd()) +'/currentBranches_4.json', orient='columns')) # import json file with all the inputs  (P,Q) and results of the power flow (I) 
df2 = df1.set_index("Id", drop = True) 

# 2. excel creation 

#Edit this line to select input file 
file = 'Sitel_Invade_MV_Topology2.xlsx'

xl  = pd.ExcelFile(Path(str(os.getcwd())+ '/'+  file))

##Parse specified sheet(s) into a DataFrame. Equivalent to read_excel 
## Llegeixo cada pestanya, dins de xl PANDAS 
 
MVNetwork = xl.parse('Linies') 
x_ohm_per_km = MVNetwork['Reactancia_ohm_km'][0]
c_nf_per_km = MVNetwork['Capacitat_uF_km'][0]
r_ohm_per_km = MVNetwork['Resistencia_ohm_km'][0]

line_data = {"c_nf_per_km": c_nf_per_km , "r_ohm_per_km": r_ohm_per_km, "x_ohm_per_km": x_ohm_per_km, "max_i_ka":0.415}

##############################################################################
################# 3. FR dataframe creation ##################################
#############################################################################

timeperiod = range(0,192)

flex_df = pd.DataFrame(index = timeperiod, columns= ['ID Node', 'P [kW]'])
#flex_df = flex_df.rename_axis('Time Period')
#flex_df = FR.fillna(0)


for i in flex_df['ID Node']: 
    flex_df['ID Node'] = 5
    
    
flex_df['P [kW]']= np.random.randint(-10,20, size=(len(flex_df), 1))
    
###############################################################################
################### 4. congestion analysis ####################################
###############################################################################

lines_overloaded_df = pd.DataFrame( index = range(0,10), columns= ['Time Period',
                                   'ID TedisNet','From Bus', 'PreFlex Line Load [%]',
                                   'PostFlex Line Load [%]'])


lines_overloaded_df.loc[0,'From Bus'] = 'E.R. REC'
lines_overloaded_df.loc[1,'From Bus'] = 'E.T. NODE'
lines_overloaded_df.loc[2,'From Bus'] = 'EMPALMAMENTS CARRER REC'
lines_overloaded_df.loc[3,'From Bus'] = 'E.T. CUARTEL'
lines_overloaded_df.loc[4,'From Bus'] = 'EMPALMAMENTS CARRER CATALUNYA CANT. TRAVESSERES'
lines_overloaded_df.loc[5,'From Bus'] = 'E.R. REC'
lines_overloaded_df.loc[6,'From Bus'] = 'E.T. LA MUTUA'
lines_overloaded_df.loc[7,'From Bus'] = 'E.T. LA MUTUA'
lines_overloaded_df.loc[8,'From Bus'] = 'E.T. EQUADOR'
lines_overloaded_df.loc[9,'From Bus'] = 'E.T. PRADES'

lines_overloaded_df.loc[0,'ID TedisNet'] = 92
lines_overloaded_df.loc[1,'ID TedisNet'] = 5
lines_overloaded_df.loc[2,'ID TedisNet'] = 10
lines_overloaded_df.loc[3,'ID TedisNet'] = 11
lines_overloaded_df.loc[4,'ID TedisNet'] = 12
lines_overloaded_df.loc[5,'ID TedisNet'] = 92
lines_overloaded_df.loc[6,'ID TedisNet'] = 34
lines_overloaded_df.loc[7,'ID TedisNet'] = 34
lines_overloaded_df.loc[8,'ID TedisNet'] = 60
lines_overloaded_df.loc[9,'ID TedisNet'] = 99


lines_overloaded_df.loc[0,'PreFlex Line Load [%]'] = 35.932
lines_overloaded_df.loc[1,'PreFlex Line Load [%]'] = 35.3611
lines_overloaded_df.loc[2,'PreFlex Line Load [%]'] = 34.9522
lines_overloaded_df.loc[3,'PreFlex Line Load [%]'] = 31.541
lines_overloaded_df.loc[4,'PreFlex Line Load [%]'] = 30.087
lines_overloaded_df.loc[5,'PreFlex Line Load [%]'] = 32.323
lines_overloaded_df.loc[6,'PreFlex Line Load [%]'] = 34.551
lines_overloaded_df.loc[7,'PreFlex Line Load [%]'] = 38.122
lines_overloaded_df.loc[8,'PreFlex Line Load [%]'] = 30.541
lines_overloaded_df.loc[9,'PreFlex Line Load [%]'] = 33.562

lines_overloaded_df.loc[0,'PostFlex Line Load [%]'] = 15.221
lines_overloaded_df.loc[1,'PostFlex Line Load [%]'] = 6.611
lines_overloaded_df.loc[2,'PostFlex Line Load [%]'] = 11.422
lines_overloaded_df.loc[3,'PostFlex Line Load [%]'] = 20.025
lines_overloaded_df.loc[4,'PostFlex Line Load [%]'] = 10.744
lines_overloaded_df.loc[5,'PostFlex Line Load [%]'] = 7.786
lines_overloaded_df.loc[6,'PostFlex Line Load [%]'] = 10.257
lines_overloaded_df.loc[7,'PostFlex Line Load [%]'] = 18.422
lines_overloaded_df.loc[8,'PostFlex Line Load [%]'] = 10.231
lines_overloaded_df.loc[9,'PostFlex Line Load [%]'] = 5.720


TimePeriod1 = pd.date_range(start = '2019-03-27 13:00:00', end='2019-03-27 14:00:00', freq='15min') 
TimePeriod2 = pd.date_range(start = '2019-03-27 13:45:00', end='2019-03-27 14:45:00', freq='15min') 

TimePeriod = TimePeriod1.append(TimePeriod2)


#for i in lines_overloaded_df.TimePeriod: 
#    lines_overloaded_df.TimePeriod[i] = datevector[i]
#    
#lines_overloaded_df.insert(loc=1, column = 'Time Period', value = TimePeriod)

#dataset_price.index = datevector

lines_overloaded_df['Time Period'] = TimePeriod

#lines_overloaded_df.loc[0,'TimePeriod'] = 5.720


"""
net.res_line.insert(14, 'from_bus', net.line['from_bus'])
net.res_line = net.res_line.set_index('from_bus', drop = False)
net.line = net.line.set_index("from_bus", drop = True)
lines_overloaded={}
for i in net.res_line['from_bus']:
    if net.res_line['loading_percent'][i]>3:  
        a=net.bus['name'][i]
        b=net.bus['Id Bus'][i]
        c=net.res_line['loading_percent'][i]
        lines_overloaded[(net.line['Id linia'][i])]=(a,b,c)

df_lines_overloaded=pd.DataFrame.from_dict(lines_overloaded, orient='index')
#cal sobreescriure la variable perquè es guardi (sinó no la podem consultar)
df_lines_overloaded=df_lines_overloaded.rename({0:'From Bus', 1:'Id Bus', 2:'overload_percentatge'}, axis='columns')



###############################################################################
############################## CREATE EXCEL FILE ##############################       
# Creaem un per Estabanell on eliminarem el Id Bus deixant només Bus origen i Id línia        
df_lines_overloaded_ESTABANELL=df_lines_overloaded
del(df_lines_overloaded_ESTABANELL['Id Bus'])

df_lines_overloaded.to_excel(Path(str(os.getcwd()) +'/PowerFlow.xlsx'))
"""


###############################################################################
################## 5. Excel file output creation ##############################
##############################################################################


#create a Pandas Excel writer using XLsWriter as the engine. 
writer = pd.ExcelWriter('fr_output.xlsx') 
writer2 = pd.ExcelWriter('congestion_log.xlsx')

flex_df.to_excel(writer)
lines_overloaded_df.to_excel(writer2)

writer.save()
writer2.save()




