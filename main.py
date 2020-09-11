import logging
import os
from datetime import date
from openpyxl import load_workbook
from api import EditAPI
from api import SelectCircles
from api import IndividualEdit
from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.screenmanager import Screen, ScreenManager


# Loading the workbook from the current directory
workbook = load_workbook(os.path.join(os.getcwd(), "4-08-2020 (2).xlsx"))

class ReportGenerator(App):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.sheet6_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 6")
        self.sheet7_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 7")
        self.sheet8_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 8")
        self.sheet9_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 9")
        self.sheet10_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 10")
        self.sheet11_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 11")

    def build(self):

        # Updating date for sheet 6
        # b = Button(text = "Update Date")
        # b.bind(on_press = self.sheet6_editor.update_date)

        # Initializing a screen manager
        self.screen_manager = ScreenManager()  
        
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
