import pandas as pd
import random
import numpy as np
from math import nan
import os
import itertools
import xlsxwriter


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


def getTheActualIteamByRemovingNanFromList(iteamList):
    tempList = []
    for i in iteamList:
        if str(i)!= 'nan':
            tempList.append(i)
    return tempList




def getPriorityTestCase(headerList, columnValueList):

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

    return myFinalList

    

def incrementAiIncrement(aiIncrement, headerList):
    i=0
    for keys, values in aiIncrement.items():
        head, newValue = rollDice(values+i, headerList)
        aiIncrement[keys]= newValue
        i+=1
    
    return aiIncrement



def rollDice(number, Dice):

	low = 0
	high = len(Dice)

	for i in range(1,number+1):
		low = low +1

		if low >high:
			low =1
		
	return Dice[low-1], low

def DiceInitialisation(headerList):
    aiIncrement = {}
    Dice = []
    i=0
    for head in headerList:
        aiIncrement[head]=i
        Dice.append(i) 
        i+=1
    
    return Dice,aiIncrement

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

    Dice, aiIncrement = DiceInitialisation(headerList)

    myFinalList = []
    for i in range(len(firstColumnFromExcell)):
        if i > 1:
            aiIncrement = incrementAiIncrement(aiIncrement, headerList)
        
        for internalLoop in range(len(firstColumnFromExcell)):
                iteam =[]
                if i == 0:
                    for header in headerList:

                        if header == headerList[0]:

                            iteam.append(getNanValueFromList(firstColumnFromExcell,i))
                        else:

                            iteam.append(getNanValueFromList(columnValueList[header], internalLoop))
                else:
                    for header in headerList:
                        if header == headerList[0]:

                            iteam.append(getNanValueFromList(firstColumnFromExcell,i))
                        else:

                            newIncrement = internalLoop+aiIncrement[header]

                            #rollDice(newIncrement,Dice)
                            #if newIncrement > len(headerList)-1:
                            #   newIncrement = 0
                            key , iteration = rollDice(newIncrement,Dice)
                            iteam.append(getNanValueFromList(columnValueList[header], key))	            		

                myFinalList.append(iteam)


    priorityTestcase = len(myFinalList)
    

    iteamLoopCouinter = 0
    ExcellResult = {}
    for header in headerList:
        iteam = []
        for singleEntry in myFinalList:
            iteam.append(singleEntry[iteamLoopCouinter])
        ExcellResult[header] = iteam
        iteamLoopCouinter += 1






    #Code to generate nonPriority TestCase
    args=[]
    for head in headerList:
        args.append(getTheActualIteamByRemovingNanFromList(columnValueList[head]))

    myNonPriorityTestCaseList = []
    for combination in itertools.product(*args):
        myNonPriorityTestCaseList.append(list(combination))
        #print(list(combination))
    #Combination ends here
    

    NonPriorityTestCase = len(myNonPriorityTestCaseList)

    destinationFolderPath = getDestinationLocation(fileName)
    finalOutPutPath = os.path.join(destinationFolderPath, "E-OATS_OUTPUT" + "." + "xlsx")
    
    
    df1 = pd.DataFrame(ExcellResult,columns = headerList)
    df2 = pd.DataFrame(myNonPriorityTestCaseList, columns = headerList)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(finalOutPutPath, engine='xlsxwriter')
    
    #df.to_excel (finalOutPutPath, index = False, header=True)
    df1.to_excel(writer, sheet_name='PriorityTestCase', index=False, header=True)
    df2.to_excel(writer, sheet_name='TotalTestCase', index=False, header=True)
    
    writer.save()

    return finalOutPutPath, priorityTestcase, NonPriorityTestCase