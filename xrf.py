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

compound_order = {} # dictionary to hold compound as key and its column number as value -> {'CaO' : 3}
for row in ws.iter_rows(min_row=substance_row, max_row=substance_row, min_col=2):
    for column, cell in enumerate(row):
        compound_order[cell.value] = column + 1
        
compound_order.pop(None)

for  file in csv_files:
    csv_data = {}
    with open(file) as csv_file:
        file_reader = csv.reader(csv_file, delimiter=',')
        for row in file_reader:
            for index, value in enumerate(row):
                if index > 1 and index % 2 != 0:
                    pass
                    #print(row[index-1])
                            
wb.close()


