from datetime import date
from openpyxl import load_workbook

#Updating the cell data in circle 6 sheet

wb = load_workbook("4-08-2020 (2).xlsx")

print(wb.sheetnames)

sheet1 = wb["Circle 6"]

update_date = "Date: " + str(date.today())

print("Updating Date: {}".format(update_date))

(sheet1["K2"].value) = update_date

wb.save("4-08-2020 (2).xlsx")

