""" Find data based on provided URLs in a table and copy them to a particular folder """

""" Import of all relevant packages for the job"""
import numpy as np
import pandas as pd
import glob as gb
import re
from pathlib import Path
import os
import shutil
#import matplotlib.pyplot as plt
#import scipy as sci
#import lmfit as lm
#import datetime as dt
#import time as tm

"""     Define the path where to find the tables and open them sequentially """
pathData = 'C:/UserData/TE_CTRX/CTRX8191B/Site_0_Issue/*.csv'
files = gb.glob(pathData)
filesNamesArr = []
col = ['File']
for file, testerNr in zip(files, range(1, len(files)+1)):
    testerNr = re.search(r'UF\d+', file).group()
    print("")
    print("Tester {}".format(testerNr))
    filesNames = pd.read_csv(file, header=0, sep="\t", skipinitialspace=True, usecols=col) # Path names to the raw data
    #print (filesNames.head())
    filesNamesArr = np.array(filesNames)
    """ Create a folder for each tester where the raw data will be copied """
    pathFolder = Path('C:/UserData/TE_CTRX/CTRX8191B/Site_0_Issue/{}'.format(testerNr))
    if pathFolder.exists()==False:
        os.makedirs(pathFolder)
    """ In the tables, find the path name to the EFF raw data and copy them to the created folder """
    print("Starting copying data from Tester {} ...".format(testerNr))
    for nameIx in filesNamesArr:
        sourcePath = nameIx[0].replace('\\','/')
        destPath =  str(pathFolder).replace('\\','/') + "/" + os.path.basename(str(nameIx[0]))
        if os.path.exists(sourcePath) == True and os.path.exists(destPath) == False: 
            shutil.copyfile(sourcePath, destPath)
    print("Data copy completed...")
    
    