import csv
import os
import glob
import shutil
import datetime
import ctypes
import warnings

import openpyxl as xl
from openpyxl.styles.fonts import Font
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import Border, Side

warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')
"""
Done by: Arttu Mäkelä 
For: CRS Laboratories
Started: June 2022

This script reads csv file in folder and outputs an Excel report based on
a template file. It renames the input CSV-file, copies the Excel report to
a subdirectory and moves the CSV-file to another. 

It can read multiple CSV-files, but currently it outputs a separate Excel-report
for each file.
"""

uniquant_template = "uniquant.xlsx"
puriste_template = "puriste.xlsx"
sulate_template = "sulate.xlsx"

# name of the määritysrajat-excel (only needed for puriste/sulate)
määritys_rajat_xl = "määritysrajat.xlsx"

csv_dir = "CSV" # CVS directory name
excel_dir = "Raportit" # Excel report directory name

# csv-file method and the excel template associated with it
method_files = {
    "X_UQ_3600W Oxides" : uniquant_template,
    "5. PhosphateConcentrateMajors_FB 0.2" : sulate_template,
    "1. PhosphateRocks_PP 1.0" : puriste_template
}

# group up all the csv files in this directory to a list
parent_path = os.path.abspath("..")
csv_files = glob.glob(os.path.join(parent_path, "*.csv"))

if not csv_files:
    ctypes.windll.user32.MessageBoxW(0, "Place CSV-file in the same folder as the xrf.py script", "Error", 0)

methods = {} # keep track of the methods of each row in csv
excels = []
csvs = [] 
date_format = "%Y-%b-%d %X" # how CSV-file represents a date
substance_row = 11 # the row where all the compounds are listed in excel templates

for num, file in enumerate(csv_files):
    with open(file) as csv_file:
        file_reader = csv.reader(csv_file, delimiter=',')
        dates_rows = {}
        # get each date in csv and put them in dictionary as datetime object keys
        # and row numbers being the values
        for row_num, row in enumerate(file_reader):
            date = datetime.datetime.strptime(row[3], date_format)
            dates_rows[date] = row_num+1
            
        # construct the date-portion of the report file's name
        current = datetime.datetime.now()
        file_date = f"{current.day}.{current.month}.{current.year}"
        # more exact time in the case that multiple reports are created from one csv
        file_date_exact = file_date + f" {current.hour} {current.minute} {current.second}"

        dates = []
        # sort the dates
        for key, value in dates_rows.items():
            dates.append(key)
        sorted_dates_rows = {}
        dates.sort()
        
        for i, key in enumerate(dates_rows.keys()):
            sorted_dates_rows[key] = dates.index(key) + 1
            
        # row_order shows where each row in csv is placed in excel report 
        row_order = [item for item in sorted_dates_rows.values()]
        csv_file.seek(0) # reset iterator to beginning
        
        
        compound_order = {} # dictionary to hold compound as key and its column number as value
        # loop through samples (sorted by date) and csv iterable
        for index, (row_num, row) in enumerate(zip(row_order, file_reader)):
            extras = 1 # count of extra compounds after "sum before norm."-cell
            for col, value in enumerate(row): # loop through each value in csv row
                current_row =  substance_row + row_num + 2  # +2 for the extra 2 rows under compounds
                if col == 0:
                    if index == 0:  # check what excel template to use and load the excel  
                        template = method_files[value]
                        wb = xl.load_workbook(template)
                        ws = wb.active
                        # read excel-template and get all the compounds in it to a dict
                        for rows in ws.iter_rows(min_row=substance_row, max_row=substance_row, min_col=2):
                            for column, cell in enumerate(rows):
                                compound_order[cell.value] = column + 1 
                                                      
                        # check if template excel is puriste/sulate and assign limits to compounds
                        if template != uniquant_template:
                            wb_limits = xl.load_workbook(määritys_rajat_xl)
                            ws_limits = wb_limits.active
                            if template == sulate_template:
                                limits_sulate = {}
                                for rows in ws_limits.iter_rows(min_row=4, min_col=2, max_col=4):
                                    if not rows[0].value:
                                        break
                                    limits_sulate[rows[0].value] = (rows[1].value, rows[2].value)
                            elif template == puriste_template:
                                limits_puriste = {}
                                for rows in ws_limits.iter_rows(min_row=4, min_col=6, max_col=8):
                                    if not rows[0].value:
                                        break
                                    limits_puriste[rows[0].value] = (rows[1].value, rows[2].value)    
                                                                                             
                        compound_order.pop(None, None) # remove trailing none-key from dict if it exists
    
                    ws.cell(row=5, column=2).value = value # place method name from csv to excel cell
                    methods[csv_file] = value
            
                elif col > 3 and col % 2 != 0:
                    try:
                        # set column number based on compound's order in excel template
                        this_column = compound_order[row[col-1]] + 1
                        ws.cell(row=current_row, 
                                column=this_column).value = float(value)
                        # calculate iron value for each row based on Fe2O3 value
                        if row[col-1] == "Fe2O3":
                            fe_column = compound_order["Fe*"] + 1
                            ws.cell(row = current_row, 
                                    column = fe_column).value = round(0.69945 * float(value),3)
                    # catch all compounds not defined in the dictionary and place
                    # their values after "Sum Before Norm." cell
                    except KeyError:
                        extras += 1
                        this_column = len(compound_order) + extras
                        compound_cell = ws.cell(row = substance_row, column = this_column)
                        compound = row[col-1]
                        def place_value(this_column):
                            try:
                                ws.cell(row = current_row, column = this_column).value = float(value)
                                compound_cell.font = Font(bold=True)
                                compound_cell.alignment = Alignment(horizontal='center')
                                compound_cell.border = Border(bottom=Side(style='thin'))
                                # styling for the 2 rows below compound
                                below_compound = ws.cell(row = substance_row + 1, column = this_column)
                                below_compound.value = "(%)"
                                below_compound.alignment = Alignment(horizontal='center')
                                below_compound.font = Font(size=8)
                                pp_xrf12 = ws.cell(row = substance_row + 2, column = this_column)
                                pp_xrf12.value = ws.cell(row = substance_row + 2, column = 2).value
                                pp_xrf12.font = Font(size=8)
                                pp_xrf12.alignment = Alignment(horizontal='center')
                            except ValueError:
                                return
                        # check if compound already placed on colum
                        if not compound_cell.value:
                            compound_cell.value = compound
                            place_value(this_column)
                        elif compound_cell.value == compound:
                            place_value(this_column)
                        # if compound is a different compound, put its value on the next empty or matching column
                        elif compound_cell.value != compound:
                            is_avail = False
                            col_idx = 1
                            while not is_avail:
                                compound_cell = ws.cell(row = substance_row, column = this_column + col_idx)
                                if compound == compound_cell.value or not compound_cell.value:
                                    compound_cell.value = compound
                                    place_value(this_column+col_idx)
                                    is_avail = True
                                col_idx += 1

                # sample name column from csv to excel report
                elif col == 1:
                    ws.cell(row=current_row, 
                            column=1).value = value
                    ws.cell(row=current_row, 
                            column=1).alignment = Alignment(horizontal='right')
                    
                # get SID2 from first row, column 3
                elif col == 2 and index == 0:
                    sid2 = value # to be used in naming the report
                    
    # check limits for sulate values and overwrite if over/under
    if template == sulate_template:
        for key, value in limits_sulate.items():
            for col in ws.iter_cols(min_row=substance_row+3, min_col=compound_order[key]+1, max_col=compound_order[key]+1):
                for cell in col:
                    if cell.value and cell.value < limits_sulate[key][0]:
                        cell.value = f"< {limits_sulate[key][0]}"
                    elif cell.value and cell.value > limits_sulate[key][1]:
                        cell.value = f"*{limits_sulate[key][1]}"
                        
    # check limits for puriste values and overwrite if over/under
    elif template == puriste_template:
        for key, value in limits_puriste.items():
            for col in ws.iter_cols(min_row=substance_row+3, min_col=compound_order[key]+1, max_col=compound_order[key]+1):
                for cell in col:
                    if cell.value and cell.value < limits_puriste[key][0]:
                        cell.value = f"< {limits_puriste[key][0]}"
                    elif cell.value and cell.value > limits_puriste[key][1]:
                        cell.value = f"> {limits_puriste[key][1]}"                    

    # check for None-values and insert value to them                
    for row in ws.iter_rows(min_row=substance_row+1+2, min_col=1, max_col=len(compound_order)):
        if row[0].value:    # check that row has a sample name         
            for cell in row:
                if not cell.value:
                    cell.value = "< 0.001"
                cell.alignment = Alignment(horizontal='right')
                              
    # rename csv file and save a new excel file
    excel_name = f"{sid2} - {file_date}.xlsx"
    csv_name = f"{sid2} - {file_date}.csv"
    
    csv_path = os.path.join(parent_path, csv_name)
    os.rename(file, csv_path)    
    wb.save(excel_name)  
    wb.close()     

    excel_path = os.path.join(os.getcwd(), excel_name)
    csvs.append(csv_path)
    excels.append(excel_path)

def move_files(names: list, dst: str, cwd=os.getcwd(), file_type=".csv"):
    """ Moves a list of files to the destination folder, script should be
        in the same folder as the files being moved.
        Renames copied file if file already exist in destination folder by
        adding more accurate timestamp
        
        args:
            - name: name of the file with file extension
            - dst: destination folder (not path, string only)
            - cwd: location of the target file in relation to this script
    """
    for name in names:
        dir = os.path.join(os.path.abspath(".."), dst)
        tail = os.path.split(name)
        file_path = os.path.join(dir, tail[1])
        if not os.path.exists(dir):
            os.mkdir(dir)
        if os.path.exists(file_path):
            file_path = os.path.join(cwd, name)
            new_name = f"{sid2} - {file_date_exact}{file_type}"
            os.rename(file_path, os.path.join(cwd, new_name)) 
            shutil.move(os.path.join(cwd, new_name), dir)
        else:
            shutil.move(name, dir) 
              
# present a warning if multiple methods in csv file        
if len(methods) > 1:
    ctypes.windll.user32.MessageBoxW(0, "More than 1 method present in CSV-file", "Warning", 0)  
      
move_files(excels, excel_dir, file_type=".xlsx")
move_files(csvs, csv_dir, cwd=parent_path)
