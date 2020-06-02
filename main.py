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
from myExcell import GenerateReport
import os


kivy.require('1.10.1') # replace with your current kivy version !
Builder.load_file('frontScreen.kv')

class FrontScreenLayout(Widget):

    def InitialAction(self, path, filename):
        ResultFileLocation, PriorityTestCase , NonPriorityTestCase = GenerateReport(filename)
        
        self.ids.OutPutFileLocation.text = "OUTPUT file : " + ResultFileLocation
        self.ids.PriorityTestCase.text = "PriorityTestCase : " + str(PriorityTestCase)
        self.ids.NonPriorityTestCase.text = "NonPriorityTestCase : " + str(NonPriorityTestCase)

class MyApp(App):
    def build(self):
        self.title= "E-OATS 0.1"
        Window.size = (600,400)        
        return FrontScreenLayout()


if __name__ == '__main__':
    MyApp().run()