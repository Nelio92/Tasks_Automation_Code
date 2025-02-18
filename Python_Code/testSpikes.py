#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Sun Jan  1 14:51:18 2017

@author: pi
"""

""" HR Parts WITHOUT OFFSET and not normalized to 1 """

""" Import of all relevant packages for the job"""
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import time as tm
import datetime as dt
from statsmodels import robust
from scipy import signal

"""     1. ERGOMETRY RAW DATA to analyse from all patients """
path = '/home/pi/workspace/ergodata_Patient1.xls'

einzelschlaege = pd.read_excel(path, sheetname=2)      # Raw Data
alle_Messwerte = pd.read_excel(path, sheetname=1)      # Raw Data
#print (einzelschlaege.head(), alle_Messwerte.head())
RR = einzelschlaege.ix[: , 2]    # RR data (Heart beat period)
hour = einzelschlaege.ix[: , 1]  # Time
index = einzelschlaege.ix[:, 0]
HR = 60e3/RR    # HR data (Heart beat rate)
HR[0] = 0
  
"""     Detect SPIKES from raw data to evaluate the goodness of samples """
HRmed = np.array(HR)
RRmed = np.array(RR)
dataSet10 = {}
set10 = []
for hr, count in zip(HRmed, enumerate(HRmed, 1)):
    if (count[0] % 10) != 0:
        set10.append(hr)
    else:
        set10.append(count[1])
        dataSet10["Set{}".format(count[0])] = set10
        set10 = []
spikes = []        
for dsK, dsV in dataSet10.items():                
    MAD = robust.scale.mad(dsV)
    MED = np.median(dsV)
    for j in dsV:
        thresU = MED+3*MAD
        thresD = MED-3*MAD
        if (j > thresU) or (j < thresD):
            spikes.append(j)
        else: continue
print("")
print("MAD = {}, Length(Spikes) = {}".format(MAD, len(spikes)))

"""     2. Build Median filtering to remove all spikes from raw data """
HRmed = signal.medfilt(HRmed, kernel_size=9)
RRmed = signal.medfilt(RRmed, kernel_size=9)

"""     3. Time strings in seconds """
t_StringSeconds = []
for i in hour:
    x = tm.strptime(i.split('.')[0],'%H:%M:%S')
    xt = dt.timedelta(hours=x.tm_hour,minutes=x.tm_min,seconds=x.tm_sec).total_seconds()
    t_StringSeconds.append(xt)
t_StringSeconds = pd.DataFrame(t_StringSeconds, columns={"Time(s)"})

"""     4. Plotting of the HR and RR Raw Data to get some insight of the median-filtering quality """ 
fig = plt.figure()
plt.grid()
plt.tight_layout()
plt.title("Heart Rate Ergometry Raw Data")
plt.ylabel("HR (bpm)")
plt.plot(t_StringSeconds, HR, 'b', label='HR Raw')
plt.scatter(range(len(spikes)), spikes, marker='o')
plt.plot(t_StringSeconds, HRmed, 'r-', label='HR median filtered')
plt.legend(loc=4)
plt.show()
     