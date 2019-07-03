# -*- coding: utf-8 -*-
"""
Editor de Spyder

Este es un archivo temporal
"""

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


Lines = np.matrix([[0,1,0.01008+0.05040j,0.1025j], [0,2,0.00744+0.03720j,0.007755j], 
                              [1,3,0.00744+0.03720j,0.07755j], [2,3,0.01272+0.06360j, 0.1275j]])

Linesdf = pd.DataFrame(Lines, columns= ['FromBus', 'ToBus', 'r+jx','b'])

fb = Lines[:,0] #vector of from node 
fb2 = np.array([0.2,1.2,2.6,3.3,4.5,5.6,6.5])

for i in range(0, np.size(Lines, 1)-1):
    a = np.real(Lines[i,1])
    fb[i]= np.int(a)

#
tb = np.array(np.size(Lines,1))
for i in range(0, np.size(Lines, 1)-1):
    r = np.real(Lines[i,2])
    fb[i]= np.int(r)




y = 1./Lines[:,2] #creation of the admittance vector of each line (inverse of the impedance)

yLineTotal = y + Lines[:,3] #total admittance of line i considering impedance and tranversal capacitance 

nb = np.int(np.amax(np.maximum(fb,tb)) +1) #number of buses we have in the network 


Ybus = np.zeros((nb,nb))  #Ybus matrix of zeros creation. nb x nb 
Ybus= Ybus + 1j #transform a real matrix into a complex matrix 

Ybus[nb*(fb-1)+tb] = -y ### falta acabar aquesta comanda. 


