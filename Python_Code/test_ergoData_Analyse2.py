""" HR Parts WITHOUT OFFSET and not normalized to 1 """

""" Import of all relevant packages for the job"""
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import lmfit as lm
import glob as gb
import time as tm
import datetime as dt
#from mpl_toolkits.mplot3d import Axes3D
from statsmodels import robust
from scipy import stats
from scipy import signal
#import corner

"""     1. ERGOMETRY RAW DATA to analyse from all patients """
#path = '/home/pi/workspace/ergodata_Patient1.xls'
#path = 'H:\WinPython-64bit-3.5.3.0Qt5/ergodata_Patient1.xls'
#path = 'H:\WinPython-64bit-3.4.4.5Qt5/ergoData_Patient_mit_Herzinsuffizienz9.xls'
#path = 'W:\Documents\Forschungspraktikum_WS_16_17/ergodata_Patient1.xls'
path = 'W:\Documents\Forschungspraktikum_WS_16_17/*.xls'
#path = 'H:\WinPython-64bit-3.4.4.5Qt5/*.xls'
files = sorted(gb.glob(path))

bestFit_Decays = {}
decaysDict = {}
decaysPatNormal = {}
decaysPatSick = {}
decaysMeanPatNormal = []
decaysMeanPatSick = []
decaysMeanDict = {}
patientNo = []
patID = []
fileID = []
patternPatientNormal = "ergodata_Patient"
patternPatientSick   = "ergodata_Patient_mit_Herzinsuffizienz"

for file, patientNr in zip(files, range(1, len(files)+1)):
    
    print("")
    print("Patient Nr.{}".format(patientNr))
    print(file)
    print("")
    fileID.append(file)
    einzelschlaege = pd.read_excel(file, sheetname=2)      # Raw Data
    alle_Messwerte = pd.read_excel(file, sheetname=1)      # Raw Data
    #print (einzelschlaege.head(), alle_Messwerte.head())
    RR = einzelschlaege.ix[: , 2]    # RR data (Heart beat period)
    hour = einzelschlaege.ix[: , 1]  # Time
    index = einzelschlaege.ix[:, 0]
    HR = 60e3/RR    # HR data (Heart beat rate)
    HR[0] = 0
      
    # Check if ENOUGH RAW DATA to analyze
    if (len(HR) < 500):
        print("")
        print("Not enough raw data for this patient...")
        continue
    
    """     2. Time strings in seconds """
    t_StringSeconds = []
    for i in hour:
        x = tm.strptime(i.split('.')[0],'%H:%M:%S')
        xt = dt.timedelta(hours=x.tm_hour,minutes=x.tm_min,seconds=x.tm_sec).total_seconds()
        t_StringSeconds.append(xt)
    
    """     3. Detect SPIKES from raw data to evaluate the goodness of samples """
    HRmed = np.array(HR)
    RRmed = np.array(RR)
    dataSet10 = {}
    dataSet10T = {}
    set10 = []
    set10T = []
    for hr, ts, countHR, countT in zip(HRmed, t_StringSeconds, enumerate(HRmed, 1), enumerate(t_StringSeconds, 1)):
        if (countHR[0] % 10) != 0:
            set10.append(hr)
            set10T.append(ts)
        else:
            set10.append(countHR[1])
            set10T.append(countT[1])
            dataSet10["Set{}".format(countHR[0])] = set10
            dataSet10T["SetT{}".format(countT[0])] = set10T
            set10 = []
            set10T = []
    dataSet10 = sorted(dataSet10.items())
    dataSet10T = sorted(dataSet10T.items())
    spikes = [] 
    spikesT = []       
    for sHR, sT in zip(dataSet10, dataSet10T):                
        MAD = robust.scale.mad(sHR[1])
        MED = np.median(sHR[1])
        thresU = MED+4*MAD
        thresD = MED-4*MAD
        for i, j in zip(sHR[1], sT[1]):
            if (i > thresU) or (i < thresD):
                spikes.append(i)
                spikesT.append(j)
            else: continue
    print("")
    qualityData = 100*(len(spikes)/len(HRmed))
    # Check if ENOUGH RAW DATA QUALITY to analyze
    if (qualityData > 20):
        print("")
        print("Too much OUTLIERS in raw data for this patient...")
        continue
    print("Outliers = {}% of Raw Data".format(round(qualityData)))
    
    """     4. Build Median filtering to remove all spikes from raw data """
#    RRmed = pd.Series.rolling(RR, window=5, min_periods=1).median()
#    HRmed = pd.Series.rolling(HR, window=5, min_periods=1).median()
    HRmed = signal.medfilt(HRmed, kernel_size=9)
    RRmed = signal.medfilt(RRmed, kernel_size=9)
    
    """     5. Plotting of the HR and RR Raw Data to get some insight of the median-filtering quality """ 
    fig = plt.figure(figsize=(12,8))
    ax1 = plt.subplot2grid((8,8), (0,0), colspan=8, rowspan=2)  # Axis for  Heart Rate Ergometry Raw Data
    ax2 = plt.subplot2grid((8,8), (2,4), colspan=4, rowspan=2)  # Axis for  HR Parts depending on the Load
    ax3 = plt.subplot2grid((8,8), (2,0), colspan=4, rowspan=4)  # Axis for  Yfit = A*[1-exp(-t/Tau)]
    ax4 = plt.subplot2grid((8,8), (4,4), colspan=4, rowspan=4)  # Axis for  Yfit = A*[exp(-t/Tau)]
    ax5 = plt.subplot2grid((8,8), (6,0), colspan=4, rowspan=2)  # Axis for  Decays in relation to the Training Loads
    #ax7 = plt.subplot2grid((8,8), (6,4), colspan=4, rowspan=2)
    ax1.grid()
    plt.tight_layout()
    ax1.set_title("Heart Rate Ergometry Raw Data")
    ax1.set_ylabel("HR (bpm)")
    ax1.plot(t_StringSeconds, HR, 'b', label='HR Raw')
    ax1.scatter(spikesT, spikes, marker='o', label='Outliers')
    ax1.plot(t_StringSeconds, HRmed, 'r-', label='HR median filtered')
    ax1.legend(loc=4)
    
    """     6. Get the duration of the corresponding load during the training"""
    wert = alle_Messwerte.ix[:, 2]
    einheit = alle_Messwerte.ix[:, 3]
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
    print("Duration Array = ".format(durationArray))
   
    """     7. Decomposition of the filtered raw data corresponding to their time slots """
    RRmedd = {}
    HRmedd = {}
    timeSlot = {}
    durArray = durationArray.copy()
    for i in range(len(durationArray)-1):
        a = durArray[i]
        b = sum(durationArray[0:i+2])
        hrmed = []
        rrmed = []
        for it, iHR, iRR in zip(t_StringSeconds, HRmed[:], RRmed[:]):
            if (it >= a) and (it <= b-1):
                hrmed.append(iHR)
                rrmed.append(iRR)
        HRmedd["{}".format(i)]  = hrmed.copy()
        RRmedd["{}".format(i)]  = rrmed.copy()
        timeSlot["{}".format(i)] = np.linspace(0, durationArray[i+1], len(hrmed))
        print("i = {} and a = {} and b = {}".format(i, a, b))
        durArray[i+1] += a
    RRmedd = sorted(RRmedd.items())
    HRmedd = sorted(HRmedd.items())
    timeSlot = sorted(timeSlot.items())     #   Time Intervalls corresponding to the load changes
    
    """     8. Offset Cancellation of the filtered raw data parts """    
    HRmedNorm = {}
    for key, val in HRmedd:
        HRmedMin = np.percentile(val, q=5, overwrite_input=True, interpolation='nearest')
        HRmedNorm["{}".format(key)] = (val-HRmedMin)
    HRmedNorm = sorted(HRmedNorm.items())   #   Normed raw data corresponding to their time slots
    
    print("")
    for hr, slot, k in zip(HRmedNorm, timeSlot, range(len(loadValue))):
        if ((loadValue[k] == 0) and (k == 0)):
            continue
        elif (loadValue[k] > loadValue[k-1]):
            ax2.grid()
            ax2.set_title("HR Parts depending on the Load")
            ax2.set_ylabel("HR Normalized (bpm)")
            print("slot= {}, hr= {}, k= {}, load= {}".format(len(slot[1]), len(hr[1]), k, loadValue[k]))
            ax2.plot(slot[1], hr[1])
        elif (loadValue[k] <= loadValue[k-1]):
            break
    
    """     9. Functions to use for the curve-fitting with lmfit """
    def Decay1(x, Amp1, tau1):
        return (Amp1*(1-np.exp(-x/tau1)))
    def Decay2(x, Amp2, tau2):
        return (Amp2*np.exp(-x/tau2))
    
    """     10. Curve fitting of the extracted normed data by INCREASING load changes """
    tau1 = []
    ampli1 = []
    dataLoadFall = []
    maxTimeLoadFall = []
    Loads = []
    loadSticks = []
    bestFit = {}
    for i, j, k, l in zip(timeSlot, HRmedNorm, range(len(loadValue)), HRmedd):
        #   Curve Fitting of Nominal Values
        if ((loadValue[k] == 0) and (k == 0)):  # NO ACTIVITY Part is not analysed
            continue
        elif ((loadValue[k] > loadValue[k-1]) and (len(j[1]) >= 100)):   # Check if increasing current load and enough data points
            data = np.array(j[1])
            data2 = -data
            offset = min(data2)
            data2 -= offset
            mod = lm.models.ExponentialModel()#lm.Model(Decay2, independant_vars=['x'])
            result = mod.fit(data2, x=i[1], amplitude=1, decay=1)
            # Check if GOODNESS of CURVE-FITTING of HR PARTS
            fitGoodnessTestI = stats.ks_2samp(result.best_fit, data2)
            print("")
            print("Fit Goodness Test result = {}".format(fitGoodnessTestI))
            if (fitGoodnessTestI[1] < 1e-2):
                continue
            tau = result.best_values['decay']
            ampli = result.best_values['amplitude']
            tau1.append(result.best_values['decay'])
            ampli1.append(result.best_values['amplitude'])
            Loads.append(loadValue[k])
            loadSticks.append(str(loadValue[k]))
            bestFit["Best_Fit_Tau_{}Watt".format(loadValue[k])] = (result.best_values.items())
            # Likelihood of Curve-Fitting
#            p = lm.Parameters()
#            p.add_many(('amplitude', 1), ('decay', 1))
#            def residual(p):
#                v = p.valuesdict()
#                return v['amplitude']*np.exp(-i[1]/v['decay']) - data2
#            mi = lm.minimize(residual, p, method='Nelder')
#            lm.printfuncs.report_fit(mi.params, min_correl=0.5)
#            # add a noise parameter
#            mi.params.add('f', value=1, min=0.001, max=2)
#            # This is the log-likelihood probability for the sampling. We're going to estimate the
#            # size of the uncertainties on the data as well.
#            def lnprob(p):
#                resid = residual(p)
#                s = p['f']
#                resid *= 1 / s
#                resid *= resid
#                resid += np.log(2 * np.pi * s**2)
#                return -0.5 * np.sum(resid)
#            mini = lm.Minimizer(lnprob, mi.params)
#            res = mini.emcee(burn=300, steps=600, thin=10, params=mi.params)  
#            corner.corner(res.flatchain, labels=res.var_names, truths=list(res.params.valuesdict().values()))
            #   Plots and Figures
            ax3.grid()
            ax3.set_ylabel("HR (bpm)")
            ax3.set_title("Yfit = A*[exp(-t/Tau)]")
            ax3.plot(i[1], data2, 'r-')
            ax3.plot(i[1], result.best_fit, label="{}W, A = {}bpm, Tau = {}s".format(loadValue[k], round(ampli), round(tau)))        
            ax3.legend()            
        elif (loadValue[k] <= loadValue[k-1]): # Relaxation Part when load decreases
            dataLoadFall.extend(l[1])
            maxTimeLoadFall.append(max(i[1]))
    
    """     11. Curve fitting of the extracted normed data by DECREASING load changes """
    timeLoadFall = np.linspace(0, sum(maxTimeLoadFall), len(dataLoadFall))
    dataLoadFall_Min = np.percentile(dataLoadFall, q=5, overwrite_input=True, interpolation='nearest')
    dataLoadFall2 = []
    for lf in dataLoadFall:
        dataLoadFall2.append((lf-dataLoadFall_Min))     # Offset Cancellation
    mod = lm.models.ExponentialModel()#lm.Model(Decay2, independant_vars=['x'])
    resultLoadFall = mod.fit(dataLoadFall2, x=timeLoadFall, amplitude=1, decay=1)
    tau = resultLoadFall.best_values['decay']
    tau2 = resultLoadFall.best_values['decay']
    ampli = resultLoadFall.best_values['amplitude']
    ampli2 = resultLoadFall.best_values['amplitude']
    fitGoodnessTestD = stats.ks_2samp(resultLoadFall.best_fit, dataLoadFall2)
    print("")
    print("Fit Goodness Test result = {}".format(fitGoodnessTestD))
    #   Plots and Figures
    ax4.grid()
    ax4.set_ylabel("HR (bpm)")
    ax4.set_title("Yfit = A*[exp(-t/Tau)]")
    ax4.plot(timeLoadFall, dataLoadFall2, 'b')
    ax4.plot(timeLoadFall, resultLoadFall.best_fit, 'r-', label="Relaxation, A = {}bpm, Tau = {}s".format(round(ampli), round(tau)))        
    ax4.legend()
    bestFit["Best_Fit_Load_Falling"] = sorted(resultLoadFall.best_values.items())
    bestFit_Decays["Tau_Best_Fit-Patient{}".format(patientNr)] = bestFit
    
    """     12. Plotting of the time constants for EACH Patient """
    Decays = tau1.copy()
    HR_Amps = ampli1.copy()
    if (fitGoodnessTestD[1] >= 1e-2):
        if not(Loads):
            Loads.append(25)
            loadSticks.append("Relaxation")
            Decays.append(tau2)
            HR_Amps.append(ampli2)
        else:
            Loads.append(Loads[-1]+25)
            loadSticks.append("Relaxation")
            Decays.append(tau2)
            HR_Amps.append(ampli2)
    ax5.grid()
    colors = np.random.rand(len(Loads))
    ax5.scatter(Loads, Decays, c=colors)
    ax5.set_xticks(Loads)            
    ax5.set_xticklabels(loadSticks)
    ax5.set_ylabel('Decays (s)')
    ax5.set_xlabel('Load (W)')
    ax5.set_title('Decays in relation to the Training Loads')      
    #fig.savefig("H:\WinPython-64bit-3.5.3.0Qt5\HR_Without_Offset_Patient{}.pdf".format(patientNr), bbox_inches='tight')
    fig.savefig("W:\Documents\Forschungspraktikum_WS_16_17\HR_Without_Offset_Patient{}.pdf".format(patientNr), bbox_inches='tight')
    plt.show() 
    
    """     13. Extract Decays for the Statistical Analysis"""
    decaysMean = np.mean(Decays)
    decaysDict["Decays_Patient{}".format(patientNr)] = Decays
    decaysMeanDict["Patient{}_Decays_Mean".format(patientNr)] = decaysMean
    if (patternPatientNormal in file):
        decaysPatNormal["PatientNormal{}".format(patientNr)] = Decays
        decaysMeanPatNormal.append(decaysMean)
        patientNo.append("N{}".format(patientNr))
        patID.append(len(patientNo))
    else:
        decaysPatSick["PatientSick{}".format(patientNr)] = Decays
        decaysMeanPatSick.append(decaysMean)
        patientNo.append("S{}".format(patientNr))
        patID.append(len(patientNo))
    
"""     14. Plotting of the time constants for ALL Patients """
fig = plt.figure(figsize=(12, 8))
for pat, dec, patNo in zip(patID, decaysDict.values(), patientNo):
    plt.grid()
    plt.scatter([pat]*len(dec), dec)
    plt.axes().set_xticks(patID)
    plt.axes().set_xticklabels(patientNo)
    plt.ylabel('Decays (s)')
    plt.title('Patients and their Decays Overview')
plt.show()
#fig.savefig("H:\WinPython-64bit-3.5.3.0Qt5\Overview_Patients_Decays.pdf", bbox_inches='tight')
fig.savefig("W:\Documents\Forschungspraktikum_WS_16_17\Overview_Patients_Decays.pdf", bbox_inches='tight')
    
"""     15. Statistical Data Analysis of the extrated Time Constants   """
#WilcoxonTest =stats.ranksums(decaysMeanPatNormal, decaysMeanPatSick)   
#print(WilcoxonTest)
""" http://www.scipy-lectures.org/intro/matplotlib/matplotlib.html """    # Interessing tutorial for Matplotlib
""" http://www.scipy-lectures.org/packages/statistics/index.html """      # Interesting tutorial for Statistics