import logging
from datetime import time
from datetime import date
from openpyxl import load_workbook
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput

class IndividualEdit(GridLayout):
    """ The next page to edit individual sheets """
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.cols = 2

        # Initializing all the sheets

        self.sheet6_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 6")
        self.sheet7_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 7")
        self.sheet8_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 8")
        self.sheet9_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 9")
        self.sheet10_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 10")
        self.sheet11_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 11")

        # Updating date for sheet 6
        b = Button(text = "Update Date")
        b.bind(on_press = self.sheet6_editor.update_date)
        self.add_widget(b)


class SelectCircles(GridLayout):
    def __init__(self, screen_manager, *kwargs):
        """ Opening Page of the App """
        super().__init__()
        self.screen_manager = screen_manager
        self.cols = 1
        self.add_widget(Label(text = "Click on any Circle to edit"))
        self.circle_6 = Button(text = "Circle 6")
        self.circle_6.bind(on_press = self.clicked)
        self.add_widget(self.circle_6)

        self.circle_7 = Button(text = "Circle 7")
        self.circle_7.bind(on_press = self.clicked)
        self.add_widget(self.circle_7)

        self.circle_8 = Button(text = "Circle 8")
        self.circle_8.bind(on_press = self.clicked)
        self.add_widget(self.circle_8)

        self.circle_9 = Button(text = "Circle 9")
        self.circle_9.bind(on_press = self.clicked)
        self.add_widget(self.circle_9)

        self.circle_10 = Button(text = "Circle 10")
        self.circle_10.bind(on_press = self.clicked)
        self.add_widget(self.circle_10)

        self.circle_11 = Button(text = "Circle 11")
        self.circle_11.bind(on_press = self.clicked)
        self.add_widget(self.circle_11)
    
    def clicked(self, instance):
        logging.info("GOING TO THE INDIVIDUAL CIRCLE EDITOR")
        self.screen_manager.current = "Next Page"


class EditAPI:
    """ This contains all the functions to edit a cell in a sheet """

    def __init__(self, workbook_name, sheet_name):
        """ workbook_name = name of the workbook, sheet_name = name of the sheet to edit """

        self.workbook_name = workbook_name
        self.workbook = load_workbook(workbook_name)
        self.sheet = self.workbook[sheet_name]
        self.date_updated = False
        logging.info("Workbook loaded successfully")


    def update_date(self, instance):
        """ Updates the date in a given sheet of the workbook """

        if self.date_updated:
            logging.info("DATE ALREADY UPDATED")
            self.workbook.save(self.workbook_name)
            return
        
        today = date.today().strftime("%d-%m-%Y")
        update_date = "Date: " + str(today)
        logging.info("Updating Date: {}".format(update_date))
        self.sheet["K2"].value = update_date

        self.date_updated = True
        logging.info("DATE UPDATED SUCCESSFULLY")
        self.workbook.save(self.workbook_name)












