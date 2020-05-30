import pandas as pd
import random
import numpy as np
from math import nan


def extractRowsFromExcell(df, first_sheet_name):
    try:
        return df[first_sheet_name]["Browsers"], df[first_sheet_name]["Capacity"],df[first_sheet_name]["Device"], df[first_sheet_name]["Year"],df[first_sheet_name]["OS"],df[first_sheet_name]["Region"], df[first_sheet_name]["HeadQ"]
    except Exception as error:
        raise AttributeError("The excell title text is not matching ", error)

def getNanValueFromList(myList,i):
    if str(myList[i])=='nan':
        while True:
            x = random.choice(myList)
            if str(x) !='nan':
                return x
    else:
        return myList[i]



def GenerateReport(fileName):
    # YOU MUST PUT sheet_name=None TO READ ALL CSV FILES IN YOUR XLSM FILE
    df = pd.read_excel(fileName, sheet_name=None)

    # prints first sheet name or any sheet if you know it's index
    first_sheet_name = list(df.keys())[0]

    Browser, Capacity, Device, Year, OS, Region, HeadQ= extractRowsFromExcell(df, first_sheet_name)

    myFinalList=[]    

    for i in range(len(Browser)):
    
        for internalLoop in range(len(Browser)):
            iteam =[]
            iteam.append(getNanValueFromList(Browser,i))
            iteam.append(getNanValueFromList(Capacity,internalLoop))
            iteam.append(getNanValueFromList(Device,internalLoop))
            iteam.append(getNanValueFromList(Year, internalLoop))
            iteam.append(getNanValueFromList(OS,internalLoop))
            iteam.append(getNanValueFromList(Region,internalLoop))
            iteam.append(getNanValueFromList(HeadQ,internalLoop))

            myFinalList.append(iteam)

    newBrowser=[]
    newCapacity=[]
    newDevice=[]
    newYear=[]
    newOs=[]
    newRegion=[]
    newHeadQ=[]


    for iteam in myFinalList:
        newBrowser.append(iteam[0])
        newCapacity.append(iteam[1])
        newDevice.append(iteam[2])
        newYear.append(iteam[3])
        newOs.append(iteam[4])
        newRegion.append(iteam[5])
        newHeadQ.append(iteam[6])



    ExcellResult = {
            'Browser': newBrowser,
            'Capacity': newCapacity,
            'Device': newDevice,
            'Year': newYear,
            'Os': newOs,
            'Region': newRegion,
            'HeadQ': newHeadQ
            }

    df = pd.DataFrame(ExcellResult, columns = ['Browser', 'Capacity','Device','Year','Os','Region','HeadQ'])
    df.to_excel (r'OTG_OUTPUT.xlsx', index = False, header=True)