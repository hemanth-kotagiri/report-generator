import logging
import os
from datetime import time
from datetime import date
from openpyxl import load_workbook
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.lang import Builder
from kivy.properties import ObjectProperty
from kivy.app import App

from android.storage import primary_external_storage_path                      
primary_ext_storage = primary_external_storage_path()
from android.permissions import request_permissions, Permission                
request_permissions([Permission.WRITE_EXTERNAL_STORAGE])


class InfoLabel(Label):
    pass

class AreaInput(TextInput):
    pass


class IndividualEdit(GridLayout):
    """ The next page to edit individual sheets """
    sm = ObjectProperty(None)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.cols = 1
        self.padding =20
        self.spacing = 50
        self.bind(minimum_height=self.setter('height'))

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


    def update_date(self):
        self.sheet_editors[self.sm.circle].update_date()

    def go_back(self):
        self.sm.current = "Main Page"

    def on_pre_enter(self, *args):
        self.clear_widgets()

        self.vmf_name = []
        self.pfm_1_km = []
        self.pfm_2_km = []
        self.vmf_km = []
        self.pfm_1_area = []
        self.pfm_2_area = []
        self.vmf_area = []
        self.my_rows = []

        for i in range(5, 12):
            # PFM-1 Input fields
            try:
                if len(self.sheet_editors[self.sm.circle].sheet["B" + str(i)].value.split("-")) == 1: continue

                if self.sheet_editors[self.sm.circle].sheet["B" + str(i)].value.split("-")[1].isnumeric() or self.sheet_editors[self.sm.circle].sheet["B" + str(i)].value == "Complaint Machine":
                    row = BoxLayout(orientation="vertical", size_hint_y=None, height=350)
                    pfm_1_layout = BoxLayout(size_hint_y=1)
                    pfm_2_layout = BoxLayout(size_hint_y=1)
                    vmf_layout = BoxLayout(size_hint_y=1)

                    divison_label = Label(text=self.sheet_editors[self.sm.circle].sheet["B" + str(i)].value or "", color=(0,0,0,1))
                    self.add_widget(divison_label)

                    pfm1_info_label = InfoLabel(text="PFM-1")
                    pfm_1_layout.add_widget(pfm1_info_label)

                    pfm2_info_label = InfoLabel(text="PFM-2")
                    pfm_2_layout.add_widget(pfm2_info_label)

                    vmf_info_label = InfoLabel(text="VMF   ")
                    vmf_layout.add_widget(vmf_info_label)

                    name_mobile_label = Label(text=self.sheet_editors[self.sm.circle].sheet["C" + str(i)].value or "")
                    pfm_1_layout.add_widget(name_mobile_label)

                    name_mobile_label_2 = Label(text=self.sheet_editors[self.sm.circle].sheet["F" + str(i)].value or "")
                    pfm_2_layout.add_widget(name_mobile_label_2)

                    vmf_name_in = AreaInput(text="Name", width = 5)
                    self.vmf_name.append(vmf_name_in)
                    vmf_layout.add_widget(vmf_name_in)

                    km_input = AreaInput(text="KM")
                    self.pfm_1_km.append(km_input)
                    pfm_1_layout.add_widget(km_input)

                    km_input_2 = AreaInput(text="KM")
                    self.pfm_2_km.append(km_input_2)
                    pfm_2_layout.add_widget(km_input_2)

                    vmf_km_in = AreaInput(text="KM")
                    self.vmf_km.append(vmf_km_in)
                    vmf_layout.add_widget(vmf_km_in)

                    area_input = AreaInput(text="Areas Covered")
                    self.pfm_1_area.append(area_input)
                    pfm_1_layout.add_widget(area_input)

                    area_input_2 = AreaInput(text="Areas Covered")
                    self.pfm_2_area.append(area_input_2)
                    pfm_2_layout.add_widget(area_input_2)

                    vmf_area = AreaInput(text="Area")
                    self.vmf_area.append(vmf_area)
                    vmf_layout.add_widget(vmf_area)

                    row.add_widget(pfm_1_layout)
                    row.add_widget(pfm_2_layout)
                    row.add_widget(vmf_layout)
                    self.add_widget(row)
                    self.my_rows.append([pfm_1_layout, pfm_2_layout, vmf_layout])
                else:
                    continue
            except Exception as e:
                logging.info(e)

    def on_parent(self, *args):
        self.parent.bind(on_pre_enter=self.on_pre_enter)

    def update_all(self):
        # Updating the KM values for PFM
        for area_list, rowname in [(self.pfm_1_km, "D"),(self.pfm_2_km, "G")]:
            for i, area in enumerate(area_list):
                self.sheet_editors[self.sm.circle].sheet[f"{rowname}{i+5}"].value = self.check_string_km(area.text)

        # Updating the Areas covered for PFMs and VMFs
        for area_list, rowname in [(self.pfm_1_area, "E"), (self.pfm_2_area, "H"), (self.vmf_km, "J"), (self.vmf_name, "I"), (self.vmf_area, "K")]:
            for i, area in enumerate(area_list):
                self.sheet_editors[self.sm.circle].sheet[f"{rowname}{i+5}"].value = self.check_area(area.text)

        self.sheet_editors[self.sm.circle].update_date()

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


class CircleButton(Button):
    pass

class SelectCircles(GridLayout):
    sm = ObjectProperty(None)

    def __init__(self, **kwargs):
        """ Opening Page of the App """
        super().__init__(**kwargs)
        self.cols = 1
        self.spacing = 10
        self.padding = 40
        self.add_widget(Label(text="TAP BELOW TO EDIT YOUR CIRCLE",
                                color=(0,0,0,1), font_size="20sp", bold=True))

        for name in range(6,12):
            self.add_widget(CircleButton(text = f"Circle {name}"))

    def clicked(self, instance):
        logging.info("GOING TO THE INDIVIDUAL CIRCLE EDITOR")
        self.sm.circle = instance.text
        self.sm.current = "Next Page"


class EditAPI:
    """ This contains all the functions to edit a cell in a sheet """

    def __init__(self, workbook_name, sheet_name):
        """ workbook_name = name of the workbook, sheet_name = name of the sheet to edit """

        self.workbook_name = workbook_name
        self.workbook = load_workbook(workbook_name)
        self.sheet = self.workbook[sheet_name]
        self.date_updated = False
        logging.info("Workbook loaded successfully")


    def update_date(self):
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
        filename = os.path.join(primary_ext_storage, today + "(2).xlsx")
        self.workbook.save(filename)
