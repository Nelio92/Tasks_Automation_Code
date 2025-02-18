""" HR Parts WITH NORMALIZATION """

""" Import of all relevant packages for the job"""
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import lmfit as lm
import glob as gb
import time as tm
import datetime as dt
#import scipy as sci

"""     ERGOMETRY RAW DATA to analyse from all patients """
#path = '/home/pi/workspace/ergodata_Patient1.xls'
#path = 'H:\WinPython-64bit-3.4.4.5Qt5/ergodata_Patient1.xls'
path = 'H:\WinPython-64bit-3.4.4.5Qt5/*.xls'
#path = 'W:\Documents\Forschungspraktikum_WS_16_17/ergodata_Patient1.xls'
#path = 'W:\Documents\Forschungspraktikum_WS_16_17/*.xls'
files = gb.glob(path)
bestFit_Decays = {}
decaysMeanDict = {}
patientNo = []
for file, patientNr in zip(files, range(1, len(files)+1)):
    print("")
    print("Patient Nr.{}".format(patientNr))
    einzelschlaege = pd.read_excel(file, sheetname=2)      # Raw Data
    alle_Messwerte = pd.read_excel(file, sheetname=1)      # Raw Data
    print (einzelschlaege.head(), alle_Messwerte.head())
    RR = einzelschlaege.ix[: , 2]    # RR data (Heart beat period)
    hour = einzelschlaege.ix[: , 1]  # Time
    index = einzelschlaege.ix[:, 0]
    HR = 60e3/RR    # HR data (Heart beat rate)
    HR[0] = 0
    RRmd = RR # RR data to normalize
    HRmd = HR   # HR data to normalize
    
    """     Build Median filtering to remove all spikes from raw data """
    RRmed  = pd.Series.rolling(RRmd, window=5, min_periods=1).median()
    HRmed  = pd.Series.rolling(HRmd, window=5, min_periods=1).median()    
    #   Time strings in seconds
    t_StringSeconds = []
    for i in hour:
        x = tm.strptime(i.split('.')[0],'%H:%M:%S')
        xt = dt.timedelta(hours=x.tm_hour,minutes=x.tm_min,seconds=x.tm_sec).total_seconds()
        t_StringSeconds.append(xt)
    t_StringSeconds = pd.DataFrame(t_StringSeconds, columns={"Time(s)"})
    
    """     Plotting of the HR and RR Raw Data to get some insight of the median-filtering quality""" 
    fig = plt.figure(figsize=(12,8))
    ax1 = plt.subplot2grid((8,8), (0,0), colspan=8, rowspan=2)
    ax2 = plt.subplot2grid((8,8), (2,4), colspan=4, rowspan=2)
    ax4 = plt.subplot2grid((8,8), (2,0), colspan=4, rowspan=4)
    ax5 = plt.subplot2grid((8,8), (4,4), colspan=4, rowspan=4)
    ax6 = plt.subplot2grid((8,8), (6,0), colspan=4, rowspan=2)
    ax1.grid()
    plt.tight_layout()
    ax3 = ax1.twinx()
    ax3.set_title("Heart Rate Ergometry Raw Data")
    #ax3.set_xlabel("t (s)")
    ax3.set_ylabel("RR (ms)")
    ax3.plot(t_StringSeconds, RR, 'b',label='RR Raw')
    ax3.legend(loc=4)
    ax3.plot(t_StringSeconds, RRmed, 'g-', label='RR median filtered')
    ax3.legend(loc=4)
    ax1.set_title("Heart Rate Ergometry Raw Data")
    #ax1.set_xlabel("t (s)", labelpad=-10)
    ax1.set_ylabel("HR (bpm)")
    ax1.plot(t_StringSeconds, HR, 'b',label='HR Raw')
    ax1.legend(loc=3)
    ax1.plot(t_StringSeconds, HRmed, 'r-', label='HR median filtered')
    ax1.legend(loc=3)
    
    """     Get the duration of the corresponding load during the training"""
    wert = alle_Messwerte.ix[:, 2]
    einheit = alle_Messwerte.ix[:, 3]
    #print (wert, einheit)
    durationArray = []
    loadValue = []
    duration = 0
    for I in range(len(wert)):
            if einheit[I] == "Watt":
                curWert = wert[I]
                nextWert = wert[I+1]
                duration += 1
                if (curWert != nextWert) or (einheit[I] != "Watt"):
                    durationArray.append(duration)
                    loadValue.append(int(curWert))   #   Load Values in Watt
                    duration = 0
                else:
                    continue
            else:
                continue
    durationArray.insert(0, 0)
    print("")
    print("Duration Array = ")
    print(durationArray)
   
    """     Decomposition of the filtered raw data corresponding to their time slots """
    RRmedd = {}
    HRmedd = {}
    timeSlot = {}
    durArray = durationArray.copy()
    for i in range(len(durationArray)-1):
        a = durArray[i]
        b = sum(durationArray[0:i+2])
        hrmed = []
        rrmed = []
        for it, iHR, iRR in zip(t_StringSeconds.values, HRmed[:], RRmed[:]):
            if (it >= a) and (it <= b-1):
                hrmed.append(iHR)
                rrmed.append(iRR)
        HRmedd["HRmed{}".format(i)]  = hrmed.copy()
        RRmedd["RRmed{}".format(i)]  = rrmed.copy()
        timeSlot["timeSlot_{}".format(i)] = np.linspace(0, durationArray[i+1], len(hrmed))
        print("")
        print("i = {} and a = {} and b = {}".format(i, a, b))
        durArray[i+1] += a
    RRmedd = sorted(RRmedd.items())
    HRmedd = sorted(HRmedd.items())
    timeSlot = sorted(timeSlot.items())     #   Time Intervalls corresponding to the load changes
    
    """     Norming of the filtered raw data parts : 0 to 1 """    
    HRmedNorm = {}
    for key, val in HRmedd:
        HRmedMin = min(val)
        HRmedMax = max(val)
        HRmedNorm["{}_Norm".format(key)] = (val-HRmedMin)/(HRmedMax-HRmedMin)
    HRmedNorm = sorted(HRmedNorm.items())   #   Normed raw data corresponding to their time slots
    #print (HRmedNorm)
    for hr, slot, ld, k in zip(HRmedNorm, timeSlot, loadValue, range(len(loadValue))):
        if ((loadValue[k] == 0) and (k == 0)):
            continue
        elif (loadValue[k] > loadValue[k-1]):
            ax2.grid()
            ax2.set_title("HR Parts depending on the Load")
            #ax2.set_xlabel("t (s)")
            ax2.set_ylabel("HR Normalized (bpm)")
            ax2.plot(slot[1], hr[1], label='{}Watt'.format(ld))
            ax2.legend(loc=4)
        elif (loadValue[k] <= loadValue[k-1]):
            break
    
    """     Functions to use for the curve-fitting with lmfit """
    def Decay1(x, tau1):
        return (1-np.exp(-x/tau1))
    def Decay2(x, tau2):
        return (np.exp(-x/tau2))
    
    """     Curve fitting of the extracted normed data by INCREASING load changes """
    tau1 = []
    tau2 = []
    dataLoadFall = []
    maxTimeLoadFall = []
    Loads = []
    bestFit = {}
    for i, j, k, l in zip(timeSlot, HRmedNorm, range(len(loadValue)), HRmedd):
        #   Curve Fitting of Nominal Values
        if ((loadValue[k] == 0) and (k == 0)):
            continue
        elif ((loadValue[k] > loadValue[k-1]) and (loadValue[k] < loadValue[k+1])):
            data = np.array(j[1])
            mod = lm.Model(Decay1, independant_vars=['x'])
            result = mod.fit(data, x=i[1], tau1=1)
            tau = result.best_values['tau1']
            tau1.append(result.best_values['tau1'])
            Loads.append(loadValue[k])
            bestFit["Best_Fit_Tau_{}Watt".format(loadValue[k])] = sorted(result.best_values.items())
            #   Plots and Figures
            ax4.grid()
            #ax4.set_xlabel("t (s)")
            ax4.set_ylabel("HR (bpm)")
            ax4.set_title("HR corresponding to load changes")
            ax4.plot(i[1], j[1], label="Load = {}W".format(loadValue[k]))
            ax4.legend(loc=2)
            ax4.plot(i[1], result.best_fit, 'r-', label="yfit_{}W = 1-exp(-t/{})".format(loadValue[k], round(tau)))        
            ax4.legend(loc=2)            
        elif (loadValue[k] <= loadValue[k-1]):
            dataLoadFall.extend(l[1])
            maxTimeLoadFall.append(max(i[1]))
    """     Curve fitting of the extracted normed data by DECREASING load changes """
    timeLoadFall = np.linspace(0, sum(maxTimeLoadFall), len(dataLoadFall))
    dataLoadFall_Min = min(dataLoadFall)
    dataLoadFall_Max = max(dataLoadFall)
    dataLoadFall2 = []
    for lf in dataLoadFall:
        dataLoadFall2.append((lf-dataLoadFall_Min)/(dataLoadFall_Max-dataLoadFall_Min))     # Norming : 0 to 1
    mod = lm.Model(Decay2, independant_vars=['x'])
    resultLoadFall = mod.fit(dataLoadFall2, x=timeLoadFall, tau2=1)
    tau = resultLoadFall.best_values['tau2']
    tau2 = resultLoadFall.best_values['tau2']
    #   Plots and Figures
    ax5.grid()
    #ax5.set_xlabel("t (s)")
    ax5.set_ylabel("HR (bpm)")
    ax5.set_title("HR corresponding to load relaxation")
    ax5.plot(timeLoadFall, dataLoadFall2, 'b', label="Load Relaxation")
    ax5.legend()
    ax5.plot(timeLoadFall, resultLoadFall.best_fit, 'r-', label="yfit = exp(-t/{})".format(round(tau)))        
    ax5.legend()
    
    bestFit["Best_Fit_Load_Falling"] = sorted(resultLoadFall.best_values.items())
    bestFit_Decays["Tau_Best_Fit-Patient{}".format(patientNr)] = bestFit
    
    Loads.append(125)
    Decays = tau1.copy()
    Decays.append(tau2) 
    
    ax6.grid()           
    ax6.scatter(np.array(Loads), np.array(Decays)) 
    ax6.set_xticks(np.array(Loads))
    ax6.set_xticklabels(Loads)
    ax6.set_ylabel('Decays (s)')
    ax6.set_xlabel('Load (W)')
    ax6.set_title('Decays in relation to the Training Loads')      
    #fig.savefig("W:\Documents\Forschungspraktikum_WS_16_17\Raw_Data_Patient.pdf")
    fig.savefig("H:\WinPython-64bit-3.4.4.5Qt5\HR_Normalized_Patient{}.pdf".format(patientNr))
    plt.show() 
    
    patientNo.append(patientNr)
    decaysMean = np.mean(Decays)
    decaysMeanDict["Patient{}_Decays_Mean".format(patientNr)] = decaysMean
    
"""     Statistical Data Analysis of the extrated Time Constants   
for keys, values in bestFit_Decays.items():
    print("")
    print("{} = ".format(keys))
    for i, j in values.items():
        print("{} = ".format(i))
        for k in j:
            print(k[:][1])
            
"""