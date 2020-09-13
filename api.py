import logging
from datetime import time
from datetime import date
from openpyxl import load_workbook
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput

class IndividualEdit(GridLayout):
    """ The next page to edit individual sheets """
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.cols = 1
        self.padding =20 
        self.spacing = 50

        # Initializing all the sheets

        self.sheet6_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 6")
        self.sheet7_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 7")
        self.sheet8_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 8")
        self.sheet9_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 9")
        self.sheet10_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 10")
        self.sheet11_editor = EditAPI("4-08-2020 (2).xlsx", "Circle 11")

        self.sheet_editors = {
            "Circle 6": self.sheet6_editor,
            "Circle 7": self.sheet7_editor,
            "Circle 8": self.sheet8_editor,
            "Circle 9": self.sheet9_editor,
            "Circle 10": self.sheet10_editor,
            "Circle 11": self.sheet11_editor
        }

    
    def update_date(self, instance):
        self.sheet_editors[root.circle].update_date(instance)
    
    def go_back(self, instance):
        root.current = "Main Page"
    
    def on_pre_enter(self, *args):
        self.clear_widgets()

        self.vmf_name = []
        self.pfm_1_km = []
        self.pfm_2_km = []
        self.vmf_km = []
        self.pfm_1_area = []
        self.pfm_2_area = []
        self.vmf_area = []

        for i in range(5, 12):
            # PFM-1 Input fields
            try:
                if len(self.sheet_editors[root.circle].sheet["B" + str(i)].value.split("-")) == 1: continue

                if self.sheet_editors[root.circle].sheet["B" + str(i)].value.split("-")[1].isnumeric() or self.sheet_editors[root.circle].sheet["B" + str(i)].value == "Complaint Machine":
                    row = BoxLayout(orientation="vertical", size_hint_y=None)
                    pfm_1_layout = BoxLayout()
                    pfm_2_layout = BoxLayout()
                    vmf_layout = BoxLayout()

                    divison_label = Label(text=self.sheet_editors[root.circle].sheet["B" + str(i)].value or "", color=(0,0,0,1))
                    self.add_widget(divison_label)

                    pfm1_info_label = Label(text="PFM-1", size_hint_x=None, width=20, bold=True, color=(0,0,0,1))
                    pfm_1_layout.add_widget(pfm1_info_label)

                    pfm2_info_label = Label(text="PFM-2", size_hint_x=None, width=20, bold=True, color=(0,0,0,1))
                    pfm_2_layout.add_widget(pfm2_info_label)

                    vmf_info_label = Label(text="VMF   ", size_hint_x=None, width=20, bold=True, color=(0,0,0,1))
                    vmf_layout.add_widget(vmf_info_label)

                    name_mobile_label = Label(text=self.sheet_editors[root.circle].sheet["C" + str(i)].value or "")
                    pfm_1_layout.add_widget(name_mobile_label)

                    name_mobile_label_2 = Label(text=self.sheet_editors[root.circle].sheet["F" + str(i)].value or "")
                    pfm_2_layout.add_widget(name_mobile_label_2)

                    vmf_name_in = TextInput(text="Name", multiline=False, width = 5)
                    self.vmf_name.append(vmf_name_in)
                    vmf_layout.add_widget(vmf_name_in)

                    km_input = TextInput(text="KM", multiline=False, width=20)
                    self.pfm_1_km.append(km_input)
                    pfm_1_layout.add_widget(km_input)

                    km_input_2 = TextInput(text="KM", multiline=False, width=20)
                    self.pfm_2_km.append(km_input_2)
                    pfm_2_layout.add_widget(km_input_2)

                    vmf_km_in = TextInput(text="KM", multiline=False, width=20)
                    self.vmf_km.append(vmf_km_in)
                    vmf_layout.add_widget(vmf_km_in)

                    area_input = TextInput(text="Areas Covered", multiline=False, width=20)
                    self.pfm_1_area.append(area_input)
                    pfm_1_layout.add_widget(area_input)

                    area_input_2 = TextInput(text="Areas Covered", multiline=False, width=20)
                    self.pfm_2_area.append(area_input_2)
                    pfm_2_layout.add_widget(area_input_2)

                    vmf_area = TextInput(text="Area", multiline=False, width=20)
                    self.vmf_area.append(vmf_area)
                    vmf_layout.add_widget(vmf_area)

                    row.add_widget(pfm_1_layout)
                    row.add_widget(pfm_2_layout)
                    row.add_widget(vmf_layout)
                    self.add_widget(row)
                    self.my_rows.append([pfm_1_layout, pfm_2_layout, vmf_layout])
                else:
                    continue
            except Exception:
                continue
        
        

    def on_parent(self, *args):
        self.parent.bind(on_pre_enter=self.on_pre_enter)
    
    def update_all(self, instance):
        # Updating the KM values for PFM

        for i in range(len(self.pfm_1_km)):
            self.sheet_editors[root.circle].sheet["D" + str(i + 5)].value = self.check_string_km(self.pfm_1_km[i].text)
        for i in range(len(self.pfm_2_km)):
            self.sheet_editors[root.circle].sheet["G" + str(i + 5)].value = self.check_string_km(self.pfm_2_km[i].text)
        
        # Updating the Areas covered for PFMs

        for i in range(len(self.pfm_1_area)):
            self.sheet_editors[root.circle].sheet["E" + str(i + 5)].value = self.check_area(self.pfm_1_area[i].text)
        
        for i in range(len(self.pfm_2_area)):
            self.sheet_editors[root.circle].sheet["H" + str(i + 5)].value = self.check_area(self.pfm_2_area[i].text)
        
        for i in range(len(self.vmf_km)):
            self.sheet_editors[root.circle].sheet["J" + str(i + 5)].value = self.check_string_km(self.vmf_km[i].text)
        
        for i in range(len(self.vmf_name)):
            self.sheet_editors[root.circle].sheet["I" + str(i + 5)].value = self.check_vmf_name(self.vmf_name[i].text)
        
        for i in range(len(self.vmf_area)):
            self.sheet_editors[root.circle].sheet["K" + str(i + 5)].value = self.check_vmf_area(self.vmf_area[i].text)

        
        self.sheet_editors[root.circle].update_date(instance)
    
    def check_string_km(self, s):
        if s.strip("KM"):
            return float(s.strip("KM"))
        else:
            return 0
    
    def check_area(self, s):
        if s == "Areas Covered" or "Areas" in s or "Covered" in s:
            return ""
        else:
            return s
    
    def check_vmf_area(self, s):
        if s != "Area":
            return s
        return ""
    
    def check_vmf_name(self, s):
        if s != "Name":
            return s
        return ""


class SelectCircles(GridLayout):
    def __init__(self, screen_manager, *kwargs):
        """ Opening Page of the App """
        super().__init__()
        self.screen_manager = screen_manager
        self.cols = 1
        self.spacing = 10
        self.padding = 40
        self.add_widget(Label(text="TAP BELOW TO EDIT YOUR CIRCLE",
                                color=(0,0,0,1), font_size="20sp", bold=True))
        
        self.circle_6 = Button(text = "Circle 6", size_hint=(0.5,0.5), font_size="20sp", background_color=(177/255, 126/255, 5/255, 1))
        self.circle_6.bind(on_press = self.clicked)
        self.add_widget(self.circle_6)

        self.circle_7 = Button(text = "Circle 7", size_hint=(0.5,0.5), font_size="20sp", background_color=(177/255, 126/255, 5/255, 1))
        self.circle_7.bind(on_press = self.clicked)
        self.add_widget(self.circle_7)

        self.circle_8 = Button(text = "Circle 8", size_hint=(0.5,0.5), font_size="20sp", background_color=(177/255, 126/255, 5/255, 1))
        self.circle_8.bind(on_press = self.clicked)
        self.add_widget(self.circle_8)

        self.circle_9 = Button(text = "Circle 9", size_hint=(0.5,0.5), font_size="20sp", background_color=(177/255, 126/255, 5/255, 1))
        self.circle_9.bind(on_press = self.clicked)
        self.add_widget(self.circle_9)

        self.circle_10 = Button(text = "Circle 10", size_hint=(0.5,0.5), font_size="20sp", background_color=(177/255, 126/255, 5/255, 1))
        self.circle_10.bind(on_press = self.clicked)
        self.add_widget(self.circle_10)

        self.circle_11 = Button(text = "Circle 11", size_hint=(0.5,0.5), font_size="20sp", background_color=(177/255, 126/255, 5/255, 1))
        self.circle_11.bind(on_press = self.clicked)
        self.add_widget(self.circle_11)
    
    def clicked(self, instance):
        logging.info("GOING TO THE INDIVIDUAL CIRCLE EDITOR")
        root.circle = instance.text
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
        self.workbook.save(today + "(2).xlsx")
    













