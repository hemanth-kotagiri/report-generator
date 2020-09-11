import logging
from datetime import time
from datetime import date
from openpyxl import load_workbook

#TODO: Add updated state in update_data method

class EditAPI:
    """ This contains all the functions to edit a cell in a sheet """

    def __init__(self, workbook_name, sheet_name):
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












