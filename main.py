import os
from datetime import date
from kivy.app import App
from kivy.uix.button import Button
from openpyxl import load_workbook


# Loading the workbook from the current directory
wb = load_workbook(os.path.join(os.getcwd(), "4-08-2020 (2).xlsx"))

class FirstApp(App):
    date_updated = False 

    def update_date(self, obj):
        if date_updated:
            print("Already Updated...")
            return

        print("Updating...")

        sheet1 = wb["Circle 6"] # hardcoding the sheetname for testing purposes
        today = date.today().strftime("%d-%m-%Y")
        sheet1["K2"].value = today 
        wb.save("4-08-2020 (2).xlsx")

        print("DATE UPDATED")


    def build(self):
        b = Button(text = "Update Date", on_press = self.update_date)
        return b


FirstApp().run()
