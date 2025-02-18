# -*- coding: utf-8 -*-
"""
Created on Fri Feb 24 13:55:15 2017

@author: z003kn2s
"""
import numpy as np
from scipy import signal
import matplotlib.pyplot as plt


sig = np.random.rand(20) - 0.5
t = range(len(sig))
wavelet = signal.ricker
widths = np.arange(1, 11)
cwtmatr = signal.cwt(sig, wavelet, widths)

xs = np.linspace(0, 90, 200)
data = np.sin(xs)
peakind = signal.find_peaks_cwt(data, np.arange(1,10))

print("sig= {}".format(sig))
print("")
print("wavelet= {}".format(wavelet))
print("")
print("widths= {}".format(widths))
print("")
print("cwtmatr= {}".format(cwtmatr))
print("")

plt.grid()
plt.plot(xs, data)
plt.show()