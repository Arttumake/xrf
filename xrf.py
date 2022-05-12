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

compound_order = {} # dictionary to hold compound as key and its column number as value
for row in ws.iter_rows(min_row=substance_row, max_row=substance_row, min_col=2):
    for column, cell in enumerate(row):
        compound_order[cell.value] = column + 1
        
compound_order.pop(None, None) # remove trailing none-key from dict if it exists

for file in csv_files:
    with open(file) as csv_file:
        file_reader = csv.reader(csv_file, delimiter=',')
        for row_num, row in enumerate(file_reader):
            extras = 0 # count of extra compounds after "sum before norm."-cell
            for index, value in enumerate(row):
                if index > 1 and index % 2 != 0:
                    try:
                        this_row = substance_row + row_num + 1
                        this_column = compound_order[row[index-1]] + 1
                        ws.cell(row=this_row, 
                                column=this_column).value = float(value)
                    # catch all compounds not defined in the dictionary yet and place
                    # their values after "Sum Before Norm." cell    
                    except KeyError:
                        extras += 1
                        print(row[index-1])
                        ws.cell(row=this_row, 
                                column=len(compound_order)+extras).value = float(value)

wb.save("new_excel.xlsx")                           
wb.close()
