# -*- coding: utf-8 -*-
"""
Created on Tue Feb 26 17:17:53 2019

@author: Ingrid - CITCEA
"""

#### import libraries 
import numpy as np 
import pandas as pd 


## import excel file grid topology 




###############################################################################
###############################################################################
############################### PF equations ##################################

## Initial data

baseMVA = 100 
basekV = 230 

### PU calculation 



## [Ybus] admittance matrix 

"""Builds the bus admittance matrix.Returns the full bus admittance matrix (i.e. for all buses) and the
matrices C{Yf} and C{Yt} which, when multiplied by a complex voltage
vector, yield the vector currents injected into each line from the
"from" and "to" buses respectively of each line. Does appropriate
conversions to p.u. """

#Lines = np.matrix([1, 2], [3,4])

#Lines = pd.DataFrame(np.array([[1,2,0.01008,0.05040,0.1025], [1,3,0.007446,0.03720,0.007755], 
#                               [2,4,0.00744,0.03720,0.07755], [3,4,0.01272, 0.06360, 0.1275]]), 
#                                columns=['FromBus', 'ToBus', 'r', 'x', 'b'])   


Lines = np.matrix([[0,1,0.01008,0.05040j,0.1025j], [0,2,0.007446,0.03720j,0.007755j], 
                              [1,3,0.00744,0.03720j,0.07755j], [2,3,0.01272, 0.06360j, 0.1275j]])


Linesdf = pd.DataFrame(Lines, columns= ['FromBus', 'ToBus', 'r', 'x', 'b'])

#Linesrows= Branch.index
#Linescolumns= Branch.columns

nb= Lines.shape[0]  ##number of buses 

Ybus = np.zeros((nb,nb))  #Ybus matrix of zeros creation. nb x nb 
Ybus= Ybus + 1j #transform a real matrix into a complex matrix 


## diagonal elements 
Ybus[0,0] = (1/(Lines[0,2] + Lines[0,3]) + Lines[0,4]/2 + 1/(Lines[1,2]+Lines[1,3]) + Lines[1,4]/2) 
Ybus[1,1] = (1/(Lines[0,2] + Lines[0,3]) + Lines[0,4]/2 + 1/(Lines[2,2]+Lines[2,3]) + Lines[2,4]/2) 
Ybus[2,2] = (1/(Lines[1,2] + Lines[1,3]) + Lines[1,4]/2 + 1/(Lines[3,2]+Lines[3,3]) + Lines[3,4]/2) 
Ybus[3,3] = (1/(Lines[2,2] + Lines[2,3]) + Lines[2,4]/2 + 1/(Lines[3,2]+Lines[3,3]) + Lines[3,4]/2) 

## non-diagonal elements 
Ybus[0,1] = - (1/(Lines[0,2] + Lines[0,3]))   
Ybus[1,0] = Ybus[0,1] 

Ybus[0,2] = - (1/(Lines[1,2] + Lines[1,3]))
Ybus[2,0] = Ybus[0,2]

Ybus[0,3] = 0 + 0j
Ybus[3,0] = Ybus[0,3]

Ybus[1,2] = 0
Ybus[2,1] = Ybus[1,2] 

Ybus[1,3] = - (1/(Lines[2,2] + Lines[2,3]))
Ybus[3,1] = Ybus[1,3]

Ybus[2,3] = - (1/(Lines[3,2] + Lines[3,3]))
Ybus[3,2] = Ybus[2,3]

Zbus = 1/Ybus 

"""
i=0
while i < nb: 
    Ybus[i,i] = sum(Lines[])  
    
    j=0
        
    while j < nb: 
        Ybus[i,j] = 
"""

## Altres coses connectades


## Transformer model (simplified)


