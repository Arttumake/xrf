import csv
import openpyxl as xl
import os
import glob

# group up all the csv files in this directory to a list
csv_files = glob.glob(os.path.join(os.getcwd(), "*.csv"))
excel_file = "Tulostiedosto ver2.xlsx"

wb = xl.load_workbook(excel_file)
ws = wb.active # define worksheet to work on
substance_row = 8 # the row where all the compounds are listed

for row in ws.iter_rows(min_row=substance_row):
    pass

"""
for  file in csv_files:
    with open(file) as csv_file:
        file_reader = csv.reader(csv_file, delimiter=)
"""

wb.save("newExcel.xlsx")
wb.close()