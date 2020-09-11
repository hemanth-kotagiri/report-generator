import logging
import os
from datetime import date
from kivy.app import App
from kivy.uix.button import Button
from openpyxl import load_workbook
from api import EditAPI


# Loading the workbook from the current directory
workbook = load_workbook(os.path.join(os.getcwd(), "4-08-2020 (2).xlsx"))

class FirstApp(App):
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
        b = Button(text = "Update Date")
        b.bind(on_press = self.sheet6_editor.update_date)
        return b


FirstApp().run()
