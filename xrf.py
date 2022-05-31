import csv
import os
import glob
import shutil

import openpyxl as xl
from openpyxl.styles import Font, Alignment

# group up all the csv files in this directory to a list
csv_files = glob.glob(os.path.join(os.getcwd(), "*.csv"))
excel_file = "Tulostiedosto ver2.xlsx"
excels = []

wb = xl.load_workbook(excel_file)
ws = wb.active # define worksheet to work on
substance_row = 8 # the row where all the compounds are listed

compound_order = {} # dictionary to hold compound as key and its column number as value
for row in ws.iter_rows(min_row=substance_row, max_row=substance_row, min_col=2):
    for column, cell in enumerate(row):
        compound_order[cell.value] = column + 1
        
compound_order.pop(None, None) # remove trailing none-key from dict if it exists

for num, file in enumerate(csv_files):
    with open(file) as csv_file:
        file_reader = csv.reader(csv_file, delimiter=',')
        for row_num, row in enumerate(file_reader):
            extras = 1 # count of extra compounds after "sum before norm."-cell
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
                        this_column = len(compound_order) + extras
                        compound_cell = ws.cell(row = substance_row, column = this_column)
                        compound_cell.value = row[index-1]
                        compound_cell.font = Font(bold=True)                                            
                        ws.cell(row = this_row, 
                                column = this_column).value = float(value)
                        
                elif index == 0:
                    ws.cell(row=substance_row + row_num + 1, 
                            column=1).value = value
                    ws.cell(row=substance_row + row_num + 1, 
                            column=1).alignment = Alignment(horizontal='right')                 
                    
    for row in ws.iter_rows(min_row=substance_row+1, min_col=2):
        for cell in row:
            if not cell.value:
                cell.value = "< 0.001"
            cell.alignment = Alignment(horizontal='right')
    
    excel_name = f"test_excel_{num+1}.xlsx"             
    wb.save(excel_name)  
    wb.close()     
    excel_path = os.path.join(os.getcwd(), excel_name)
    excels.append(excel_name)             
                        
def move_file(names: list, dst: str) -> None:
    """ Moves a list of files to the destination folder
        args:
            - name: name of the file with file extension
            - dst: destination folder
    """
    for name in names:
        dir = os.path.join(os.getcwd(), dst)
        if not os.path.exists(dir):
            os.mkdir(dir)
        shutil.move(name, dir)
    
#move_file(excels, 'Raportit')
#move_file(csv_files, 'CSV')
