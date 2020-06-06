import pandas as pd
import random
import numpy as np
from math import nan
import os
import itertools
import xlsxwriter



class Report:

    PRIORITY_TEST_CASE_COUNT = 0
    TOTAL_TEST_CASE_COUNT = 0
    NON_PRIORITY_TEST_CASE_COUNT = 0

    def __init__(self, fileFullPath):
        self.__filePath = fileFullPath
    
    def GetExcellHeaders(self):
        return self.df[self.first_sheet_name].columns.values.tolist()
        
    def extractRowsFromExcell(self):
        columnValues={}
        for title in self.headerList:
            columnValues[title]=self.df[self.first_sheet_name][title]
        
        return columnValues

    def incrementAiIncrement(self, aiIncrement, headerList):
        i=0
        for keys, values in aiIncrement.items():
            head, newValue = self.rollDice(values+i, headerList)
            aiIncrement[keys]= newValue
            i+=1
        
        return aiIncrement

    def getPriorityTestCase(self,aiIncrement, dice):
        myFinalList = []
        for i in range(len(self.firstColumnFromExcell)):
            if i > 1:
                aiIncrement = self.incrementAiIncrement(aiIncrement, self.headerList)
            
            for internalLoop in range(len(self.firstColumnFromExcell)):
                    iteam =[]
                    if i == 0:
                        for header in self.headerList:

                            if header == self.headerList[0]:

                                iteam.append(self.getNanValueFromList(self.firstColumnFromExcell,i))
                            else:
                                iteam.append(self.getNanValueFromList(self.columnValueList[header], internalLoop))
                    else:
                        for header in self.headerList:
                            if header == self.headerList[0]:

                                iteam.append(self.getNanValueFromList(self.firstColumnFromExcell,i))
                            else:

                                newIncrement = internalLoop+aiIncrement[header]

                                #rollDice(newIncrement,Dice)
                                #if newIncrement > len(headerList)-1:
                                #   newIncrement = 0
                                key , iteration = self.rollDice(newIncrement,dice)
                                iteam.append(self.getNanValueFromList(self.columnValueList[header], key))	            		

                    myFinalList.append(iteam)

        return myFinalList

    def diceInitialisation(self):
        aiIncrement = {}
        Dice = []
        i=0
        for head in self.headerList:
            aiIncrement[head]=i
            Dice.append(i) 
            i+=1
        
        return Dice,aiIncrement

    def getNanValueFromList(self, myList, i):
        if str(myList[i])=='nan':
            while True:
                x = random.choice(myList)
                if str(x) !='nan':
                    return x
        else:
            return myList[i]

    def rollDice(self, number, Dice):
        low = 0
        high = len(Dice)

        for i in range(1,number+1):
            low = low +1

            if low >high:
                low =1
            
        return Dice[low-1], low

    
    def getTheActualIteamByRemovingNanFromList(self, iteamList):
        tempList = []
        for i in iteamList:
            if str(i)!= 'nan':
                tempList.append(i)
        return tempList

    def getNonPriorityTestCase(self):
        #Code to generate nonPriority TestCase
        args=[]
        for head in self.headerList:
            args.append(self.getTheActualIteamByRemovingNanFromList(self.columnValueList[head]))

        myNonPriorityTestCaseList = []
        for combination in itertools.product(*args):
            myNonPriorityTestCaseList.append(list(combination))
            #print(list(combination))
        #Combination ends here

        return myNonPriorityTestCaseList

    def getDestinationLocation(self):
        head_tail = os.path.split(self.__filePath) 
        return head_tail[0]

    def generateReport(self):
        self.df = pd.read_excel(self.__filePath, sheet_name=None)

        # prints first sheet name or any sheet if you know it's index
        self.first_sheet_name = list(self.df.keys())[0]

        #Get the list of headder from the excell
        self.headerList = self.GetExcellHeaders()

        self.columnValueList = self.extractRowsFromExcell()

        #Getting the first Colum, since this shall be used as starting point for the Base Algorith
        self.firstColumnFromExcell = self.columnValueList[self.headerList[0]]

        dice, aiIncrement = self.diceInitialisation()

        self.priorityTestCase = self.getPriorityTestCase(aiIncrement,dice)

        self.PRIORITY_TEST_CASE_COUNT = len(self.priorityTestCase)
        
        self.nonPriorityTestCase = self.getNonPriorityTestCase()

        self.NON_PRIORITY_TEST_CASE_COUNT = len(self.nonPriorityTestCase)
        

        destinationFolderPath = self.getDestinationLocation()
        finalOutputPath = os.path.join(destinationFolderPath, "E-OATS_OUTPUT" + "." + "xlsx")
        
        
        df1 = pd.DataFrame(self.priorityTestCase,columns = self.headerList)
        df2 = pd.DataFrame(self.nonPriorityTestCase, columns = self.headerList)

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(finalOutputPath, engine='xlsxwriter')
        
        #df.to_excel (finalOutPutPath, index = False, header=True)
        df1.to_excel(writer, sheet_name='PriorityTestCase', index=False, header=True)
        df2.to_excel(writer, sheet_name='TotalTestCase', index=False, header=True)
        
        writer.save()

        return finalOutputPath, self.PRIORITY_TEST_CASE_COUNT, self.NON_PRIORITY_TEST_CASE_COUNT