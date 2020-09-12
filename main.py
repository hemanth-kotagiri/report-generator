import logging
import os
from datetime import date
from openpyxl import load_workbook
import api
from api import EditAPI
from api import SelectCircles
from api import IndividualEdit
from kivy.app import App
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.core.window import Window


#TODO: Add this inside the API
# Loading the workbook from the current directory
workbook = load_workbook(os.path.join(os.getcwd(), "4-08-2020 (2).xlsx"))

Window.clearcolor = (135/255,206/255,235/255,1)
Window.size = (360, 600)

class ReportGenerator(App):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def build(self):

        # Initializing a screen manager
        self.screen_manager = ScreenManager()
        api.root = self.screen_manager
        
        # Adding select circles(Main page) to the screen_manager
        self.select_circles_page = SelectCircles(screen_manager=self.screen_manager)
        screen = Screen(name="Main Page")
        screen.add_widget(self.select_circles_page)
        self.screen_manager.add_widget(screen)

        # Adding Individual editing Screen to the screen_manager
        self.next_page = IndividualEdit()
        screen = Screen(name="Next Page")
        screen.add_widget(self.next_page)
        self.screen_manager.add_widget(screen)

        return self.screen_manager


if __name__ == "__main__":
    Application = ReportGenerator()
    Application.run()
