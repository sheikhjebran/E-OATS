import pandas as pd
import random
import numpy as np
from math import nan
import os
import itertools
import xlsxwriter
import threading
import logging
import time
from enum import Enum
from collections import OrderedDict as od
import value


class Report:

    PRIORITY_TEST_CASE_COUNT = 0
    TOTAL_TEST_CASE_COUNT = 0
    NON_PRIORITY_TEST_CASE_COUNT = 0

    def __init__(self, fileFullPath):
        self.__filePath = fileFullPath
        self.STAGE_1 = Status.PENDING
        self.STAGE_2 = Status.PENDING
        self.STAGE_3 = Status.PENDING
        format = "%(asctime)s: %(message)s"
        logging.basicConfig(format=format, level=logging.INFO,
                        datefmt="%H:%M:%S")
    
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

                                key , iteration = self.rollDice(newIncrement,dice)
                                iteam.append(self.getNanValueFromList(self.columnValueList[header], key))	            		

                    myFinalList.append(iteam)

        
        self.priorityTestCase = myFinalList
        self.PRIORITY_TEST_CASE_COUNT = len(self.priorityTestCase)
        self.STAGE_1 = Status.COMPLETE
        logging.info("Thread %s: finishing", "ProrityTestCase")
        value.ProgressBarValue = 0.25
        
        

    def excellWriter(self ,df ,writer ,sheetName):
        value.MessageText = "Adding Data to : " + sheetName 
        logging.info("Writing to Excell for sheet %s Started", sheetName)
        df.to_excel(writer, sheet_name=sheetName, index=False, header=True)
        logging.info("Writing to Excell for sheet %s Finished", sheetName)
        value.MessageText = "Entries added to : " + sheetName 

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

    def getTotalTestCase(self):
        #Code to generate nonPriority TestCase
        args=[]
        for head in self.headerList:
            args.append(self.getTheActualIteamByRemovingNanFromList(self.columnValueList[head]))

        myTotalTestCaseList = []
        for combination in itertools.product(*args):
            myTotalTestCaseList.append(list(combination))
            #print(list(combination))
        #Combination ends here

        #return myTotalTestCaseList
        self.totalTestCase = myTotalTestCaseList
        self.TOTAL_TEST_CASE_COUNT = len(self.totalTestCase)
        self.STAGE_2 = Status.COMPLETE
        logging.info("Thread %s: finishing", "TotalTestCase")
        value.ProgressBarValue = 0.50

    def getDestinationLocation(self):
        head_tail = os.path.split(self.__filePath) 
        return head_tail[0]

    def validateTheDataFrame(self, rowDictionary, df):

        flag = False
        for row in df.itertuples(index=False):
            if row == rowDictionary:
                flag=True
                break
        
        if flag!=True:
            self.nonPriorityTestCaseDataList.append(list(rowDictionary))
            

    def generateReport(self):
        value.MessageText = "Reading File "
        self.df = pd.read_excel(self.__filePath, sheet_name=None)

        # prints first sheet name or any sheet if you know it's index
        self.first_sheet_name = list(self.df.keys())[0]

        #Get the list of headder from the excell
        self.headerList = self.GetExcellHeaders()

        self.columnValueList = self.extractRowsFromExcell()

        #Getting the first Colum, since this shall be used as starting point for the Base Algorith
        self.firstColumnFromExcell = self.columnValueList[self.headerList[0]]

        dice, aiIncrement = self.diceInitialisation()

        threads = list()
        
        value.MessageText = "Generating PriorityTestCase"
        logging.info("Main    : create and start thread %s.", "PriorityTestCase")
        myThread = threading.Thread(target=self.getPriorityTestCase, args=(aiIncrement,dice,))
        threads.append(myThread)
        myThread.start()

        
        value.MessageText = "Generating TotalTestCase"
        logging.info("Main    : create and start thread %s.", "TotalTestCase")
        myThread = threading.Thread(target=self.getTotalTestCase)
        threads.append(myThread)
        myThread.start()

     
        

        destinationFolderPath = self.getDestinationLocation()
        finalOutputPath = os.path.join(destinationFolderPath, "E-OATS_OUTPUT" + "." + "xlsx")
        
        
  
        for index, thread in enumerate(threads):
            logging.info("Main    : before joining thread %d.", index)
            thread.join()
            logging.info("Main    : thread %d done", index)


        df1 = pd.DataFrame(self.priorityTestCase,columns = self.headerList)
        df3 = pd.DataFrame(self.totalTestCase, columns = self.headerList)


        value.MessageText = "Generating NonPriorityTestCase"
        self.nonPriorityTestCaseDataList=[]
        for row in df3.itertuples(index=False):
            self.validateTheDataFrame(row,df1)
        
        df2 = pd.DataFrame(self.nonPriorityTestCaseDataList, columns = self.headerList)

        self.NON_PRIORITY_TEST_CASE_COUNT = len(self.nonPriorityTestCaseDataList)

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        value.MessageText = "Creating OutPut File "
        writer = pd.ExcelWriter(finalOutputPath, engine='xlsxwriter')
        
        value.ProgressBarValue = 0.75
        #df.to_excel (finalOutPutPath, index = False, header=True)

        
        myThread = threading.Thread(target=self.excellWriter, args=(df1,writer,"PriorityTestCase",))
        threads.append(myThread)
        myThread.start()

        myThread = threading.Thread(target=self.excellWriter, args=(df2,writer,"NonPriorityTestCase",))
        threads.append(myThread)
        myThread.start()

        myThread = threading.Thread(target=self.excellWriter, args=(df3,writer,"TotalTestCase",))
        threads.append(myThread)
        myThread.start()

        value.ProgressBarValue = 0.90

        for index, thread in enumerate(threads):
            logging.info("Main    : before joining thread %d.", index)
            thread.join()
            logging.info("Main    : thread %d done", index)
    
        myThread = threading.Thread(target=self.savingExcell, args=(writer,))
        threads.append(myThread)
        myThread.start()
        

        value.ProgressBarValue = 1

        return finalOutputPath, self.PRIORITY_TEST_CASE_COUNT, self.NON_PRIORITY_TEST_CASE_COUNT, self.TOTAL_TEST_CASE_COUNT

    def savingExcell(self, writer):
        value.MessageText = "Saving The Excell"
        writer.save()
        value.MessageText = "Completed all TASK .... !"
