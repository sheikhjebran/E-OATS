import kivy
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.button import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from kivy.lang.builder import Builder
from kivy.uix.widget import Widget
from kivy.core.window import Window
from kivy.uix.popup import Popup
from myExcell import Report
import os
import threading
import time
import value

kivy.require('1.10.1') # replace with your current kivy version !
Builder.load_file('frontScreen.kv')

class FrontScreenLayout(Widget):

    def InitialAction(self, path, filename):
        self.ids.OutPutFileLocation.text = ""
        self.ids.PriorityTestCase.text = ""
        self.ids.NonPriorityTestCase.text = ""
        self.ids.TotalTestCase.text = ""
        
        myThread = threading.Thread(target=self.getPriorityTestCase)
        myThread.start()


        myThread = threading.Thread(target=self.MainClass, args=(filename,))
        myThread.start()

    
    def getPriorityTestCase(self):
        
        self.ids.ProgressBar.opacity = 0

        value.ProgressBarValue = 0

        while value.ProgressBarValue!=1:
            if value.ProgressBarValue > 0.25:
                self.ids.ProgressBar.opacity = 1
            self.ids.ProgressBar.value = value.ProgressBarValue
            self.ids.OutPutFileLocation.text = value.MessageText
            time.sleep(1)

        self.ids.ProgressBar.value = value.ProgressBarValue
        self.ids.ProgressBar.height = '0dp'
        self.ids.ProgressBar.opacity = 0
        


    def MainClass(self, filename):
        report = Report(filename)
        ResultFileLocation, PriorityTestCase , NonPriorityTestCase, TotalTestCaseCount = report.generateReport()#GenerateReport(filename)
        
        self.ids.OutPutFileLocation.text = "OUTPUT file : " + ResultFileLocation
        self.ids.PriorityTestCase.text = "PriorityTestCase : " + str(PriorityTestCase)
        self.ids.NonPriorityTestCase.text = "NonPriorityTestCase : " + str(NonPriorityTestCase)
        self.ids.TotalTestCase.text = "TotalTestCase : "+ str(TotalTestCaseCount)

class MyApp(App):
    def build(self):
        self.title= "E-OATS v0.3"
        Window.size = (600,500)        
        return FrontScreenLayout()


if __name__ == '__main__':
    MyApp().run()