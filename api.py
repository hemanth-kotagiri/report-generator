import logging
import os
from datetime import date, time
from openpyxl import load_workbook
 
from kivy.uix.recycleview import RecycleView
from kivy.properties import ObjectProperty
from kivy.lang import Builder
from kivy.app import App
 
from android.storage import primary_external_storage_path
primary_ext_storage = primary_external_storage_path()
from android.permissions import request_permissions, Permission
request_permissions([Permission.WRITE_EXTERNAL_STORAGE])
 
 
class IndividualEdit(RecycleView):
    """ The next page to edit individual sheets """
    sm = ObjectProperty(None)
 
    def init_data(self, *args):
        app = App.get_running_app()
        self.data = []
        for i in range(5, 12):
            if app.sheet_editors[app.circle].sheet[f"B{i}"].value.title() != "Total":
                self.data.append(
                    {"l1":app.sheet_editors[app.circle].sheet[f"B{i}"].value or "",
                      "l2":app.sheet_editors[app.circle].sheet[f"C{i}"].value or "",
                      "l3":app.sheet_editors[app.circle].sheet[f"F{i}"].value or "",
                      "km_input1": "KM",
                      "area_input1": "Areas Covered",
                      "km_input2": "KM",
                      "area_input2": "Areas Covered",
                      "vmf_input": "Name",
                      "vmf_km_input": "KM",
                      "vmf_area_input": "Areas Covered",
                      "index": i-5}
                )
            else:
                break

 
    def update_all(self):
        app = App.get_running_app()
 
        for data in self.data:
            for inp, rowname in [("km_input1", "D"),("km_input2", "G"), ("vmf_km_input", "J"),]:
                app.sheet_editors[app.circle].sheet[f"{rowname}{data['index']+5}"].value = self.check_string_km(data.get(inp))
 
            for inp, rowname in [("area_input1", "E"), ("area_input2", "H"),  ("vmf_input", "I"), ("vmf_area_input", "K")]:
                #print(data.get(inp))
                app.sheet_editors[app.circle].sheet[f"{rowname}{data['index']+5}"].value = self.check_string_area(data.get(inp))
 
        app.sheet_editors[app.circle].update_date()
 
    def check_string_km(self, s):
        if s.strip("KM"):
            return float(s.strip("KM"))
        return 0
 
    def check_string_area(self, s):
        if s == "Areas Covered" or "Areas" in s or "Covered" in s or "Name" in s:
            return ""
        print(s)
        return s
 

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
        today = date.today().strftime("%d-%m-%Y")
 
        if self.date_updated:
            logging.info("DATE ALREADY UPDATED")
            self.workbook.save(primary_ext_storage, today + "(2).xlsx")
            return
 
        update_date = "Date: " + str(today)
        logging.info("Updating Date: {}".format(update_date))
        self.sheet["K2"].value = update_date
 
        self.date_updated = True
        logging.info("DATE UPDATED SUCCESSFULLY")
        filename = os.path.join(primary_ext_storage, today + "(2).xlsx")
        self.workbook.save(filename)
 