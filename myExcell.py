import pandas as pd
import random
import numpy as np
from math import nan
import os


def extractRowsFromExcell(df, first_sheet_name, headerList):
    columnValues={}
    for title in headerList:
        columnValues[title]=df[first_sheet_name][title]
    
    return columnValues


def getNanValueFromList(myList,i):
    if str(myList[i])=='nan':
        while True:
            x = random.choice(myList)
            if str(x) !='nan':
                return x
    else:
        return myList[i]


def GetExcellHeaders(df, first_sheet_name):
    return df[first_sheet_name].columns.values.tolist()


def getDestinationLocation(fileName):
    head_tail = os.path.split(fileName) 
    return head_tail[0]

def getTheCountExcludingNanValue(myList):
    count =0
    for i in myList:
        if str(i)!='nan':
            count+=1
    return count

def GetNonPriorityTestCase(headerList, columnValueList):
    TotalTestCase = 1
    for head in headerList:
        TotalTestCase = TotalTestCase * getTheCountExcludingNanValue(columnValueList[head])
    return TotalTestCase 

def GenerateReport(fileName):
    # YOU MUST PUT sheet_name=None TO READ ALL CSV FILES IN YOUR XLSM FILE
    df = pd.read_excel(fileName, sheet_name=None)

    # prints first sheet name or any sheet if you know it's index
    first_sheet_name = list(df.keys())[0]

    #Get the list of headder from the excell
    headerList = GetExcellHeaders(df, first_sheet_name)

    columnValueList = extractRowsFromExcell(df, first_sheet_name, headerList)

    #Getting the first Colum, since this shall be used as starting point for the Base Algorith
    firstColumnFromExcell = columnValueList[headerList[0]]

    myFinalList = []
    for i in range(len(firstColumnFromExcell)):

        for internalLoop in range(len(firstColumnFromExcell)):
            iteam =[]
            for header in headerList:

                if header == headerList[0]:
                    iteam.append(getNanValueFromList(firstColumnFromExcell,i))
                else:
                    iteam.append(getNanValueFromList(columnValueList[header], internalLoop))

            myFinalList.append(iteam)

    NonPriorityTestCase = GetNonPriorityTestCase(headerList, columnValueList)
        
    priorityTestcase = len(myFinalList)
    

    iteamLoopCouinter = 0
    ExcellResult = {}
    for header in headerList:
        iteam = []
        for singleEntry in myFinalList:
            iteam.append(singleEntry[iteamLoopCouinter])
        ExcellResult[header] = iteam
        iteamLoopCouinter += 1

    
    destinationFolderPath = getDestinationLocation(fileName)
        
    finalOutPutPath = os.path.join(destinationFolderPath, "E-OATS_OUTPUT" + "." + "xlsx")
    
    df = pd.DataFrame(ExcellResult,columns = headerList)
    df.to_excel (finalOutPutPath, index = False, header=True)

    return finalOutPutPath, priorityTestcase, NonPriorityTestCase