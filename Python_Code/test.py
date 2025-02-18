# -*- coding: utf-8 -*-
"""
Created on Thu Jul 26 10:10:16 2018

@author: wandji
"""

import os

sPath = "C:/UserData/WinPython/Pics"

files = [f for f in os.scandir(sPath)] 
print (files)