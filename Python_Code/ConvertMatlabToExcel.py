# -*- coding: utf-8 -*-
"""
Created on Fri May 10 10:53:07 2019

@author: wandji
"""

import scipy.io
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd


path = 'C:/Users/wandji/Downloads/Sample1_PN_Spur_78600.mat'

MatlabFile = scipy.io.loadmat(path)
print(MatlabFile)

print("###########################")
con_list = [[element for element in upperElement] for upperElement in MatlabFile['res']]
print(con_list)

# PN = [[element for element in upperElement] for upperElement in con_list['PN']]
# print(PN)



#PTX = mat['PTX']
#VTX = mat['VTX']
#
#PTX2 = mat2.ix[:, :]
#
#print(PTX[:, :])
#print(PTX2[:, :])
#
#plt.figure()
#plt.grid()
#plt.legend()
#plt.plot(PTX, VTX, 'b')
#plt.xlabel("PTX (dBm)")
#plt.ylabel("VmuxA - VmuxB (mV)")


