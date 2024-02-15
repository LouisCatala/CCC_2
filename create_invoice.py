import sys
import pandas as pd
import openpyxl
from pathlib import Path
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import shutil
from shutil import copy2


def read_excel_sge(file_path, sheet_name): #look up sheet name
    return pd.read_excel(file_path, sheet_name=sheet_name)

# Get the current location of the Python script
script_location = Path(__file__).resolve().parent
upper_directory = script_location.parent

def search_and_gather(excel_data, row_numbers, column_names=None):  #1.2
    result = {}
    for sheet, data in excel_data.items():
        for row_number in row_numbers:
            if row_number < len(data):
                row_data = data.iloc[row_number]
                if column_names is None:
                    # If no specific columns are specified, add the entire row
                    result.setdefault(sheet, []).append(row_data)
                else:
                    # If specific columns are specified, gather data from those columns
                    for column_name in column_names:
                        if column_name in row_data.index:
                            result.setdefault(sheet, {}).setdefault(column_name, []).append(row_data[column_name])
    return result


def process_string(txt): #Break up Coin colmun in Q1 sheet, to string_list  ↓
    # Split the string at every occurrence of ', '
    string_list = txt.split(', ')

    # Return the list of split strings
    return string_list


def create_qty_list(txt_items, qty_list):  #Remove all () in string and update string_list, and modify/create qty_list
    updated_txt_items = []

    for txt in txt_items:
        # Find the first occurrence of '()' and extract the number
        if '(' in txt and ')' in txt:
            start = txt.find('(') + 1
            end = txt.find(')', start)
            number = int(txt[start:end])
            qty_list.append(number)

            # Remove the first '()' and its contents from the string
            txt = txt[:start-1] + txt[end+1:]

        # Remove all other occurrences of '()' and its contents
        while '(' in txt and ')' in txt:
            start = txt.find('(')
            end = txt.find(')', start) + 1
            txt = txt[:start] + txt[end:]

        updated_txt_items.append(txt.strip())

    return updated_txt_items
# EXCEL_SHEET_PATH = r'C:\Users\13473\Desktop\CCC\Invoice temple_CCC.xlsx'
EXCEL_SHEET_PATH = script_location / 'Invoice temple_CCC.xlsx'
def fill_in(sheet, cell_number, data):
    sheet[cell_number] = data

def update_workbook(cell_updates, workbook_path):
    workbook = openpyxl.load_workbook(workbook_path)
    sheet = workbook.active

    for cell_number, data in cell_updates.items():
        fill_in(sheet, cell_number, data)

    workbook.save(workbook_path)
    workbook.close()
    
def combine_updates(*args):
    combined_dict = {}
    for dictionary in args:
        combined_dict.update(dictionary)
    return combined_dict

def main(row_number):
    adjusted_row_number = row_number - 2
    #CCC Phone Sales 2023.xlsx
    # source_file_path = r'C:\Users\13473\Desktop\CCC\CCC Phone Sales 2023.xlsx'  # Path to the source Excel file
    source_file_path = script_location / 'CCC Phone Sales 2023.xlsx'
    #---------------------------------------Q1 Page Update for CCC Phone Sales To Create new Invoice ----------------------------------
    source_data = read_excel_sge(source_file_path, 'Q1') #when put row number for Q1 sheet, do the actual row number -2, because starts at two.

    # Gather relevant data  for fill in section (Next Section is Fill In)
    gathered_data = search_and_gather(source_data,[adjusted_row_number]) # actual get line 2, so it is 0. which is acutal line -2
    cell_dict_Q1 = {'H3': '', 'H4': '', 'C6': '', 'B13': '', 'C13': '', 'D13': '', 'F6': ''}
    #For column QTY, GRADE, YEAR
    B2i_list = ["B" + str(16 + 2 * i) for i in range(10)]
    print("QTY",B2i_list)
    D2i_list = ["D" + str(16 + 2 * i) for i in range(10)]
    print("grade/coin name", D2i_list)
    year2i_list = ["C" + str(16 + 2 * i) for i in range(10)]
    print("year_list", year2i_list)
    # Convert lists into a dictionary for update_workbook method
    empty_list = [""] * 10
    dictionary_qty = dict(zip(B2i_list, empty_list))
    print(dictionary_qty)
    dictionary_grade = dict(zip(D2i_list, empty_list))
    print(dictionary_grade)
    dictionary_year = dict(zip(year2i_list, empty_list))
    print(dictionary_year)
    all_updates = combine_updates(cell_dict_Q1, dictionary_qty, dictionary_grade, dictionary_year)
    
    update_workbook(all_updates, EXCEL_SHEET_PATH)
    #----------------------------H3:Date----------------------------------------------
    formatted_date = "Unknown Date"
    date_object = gathered_data['Unnamed: 0'][0]
    if isinstance(date_object, datetime):   
        formatted_date = date_object.strftime("%m/%d/%Y")  
    cell_dict_Q1['H3'] = formatted_date
    #---------------------------------- H4 INV. # ---------------------------------------#
    InvNUMBER_object = gathered_data['INV. #'][0]
    cell_dict_Q1['H4'] = InvNUMBER_object
    cell_dict_Q1['C13'] = InvNUMBER_object
    cell_dict_Q1['D13'] = InvNUMBER_object
    #----------------------------------C6 Customer---------------------------------------#
    Cust_Obj = gathered_data['Customer'][0]
    cell_dict_Q1['C6'] = Cust_Obj
    cell_dict_Q1['F6'] = Cust_Obj
    #----------------------------------B13 salesman-------------------------------------#
    Salesman_Obj = gathered_data['Salesman'][0]

    cell_dict_Q1['B13'] = Salesman_Obj

    #Fill in Section

    #--------------------------------Q1 single cell updates ended-------------------------#
    Qty_list = []
    coin_list = gathered_data['Coin'][0]
    string_list1 = process_string(coin_list)
    string_list = create_qty_list(string_list1, Qty_list)
    year = date_object.strftime("%Y") 
    year_list = []
    for i in range(len(Qty_list)):
        year_list.append(year)
    #-----------Above clear--------------#
    #For column QTY, GRADE, YEAR
    B2i_list = ["B" + str(16 + 2 * i) for i in range(len(Qty_list))]
    print("QTY",B2i_list)
    D2i_list = ["D" + str(16 + 2 * i) for i in range(len(string_list))]
    print("grade/coin name", D2i_list)
    year2i_list = ["C" + str(16 + 2 * i) for i in range(len(Qty_list))]
    print("year_list", year2i_list)
    # Convert lists into a dictionary for update_workbook method
    dictionary_qty = dict(zip(B2i_list, Qty_list))
    print(dictionary_qty)
    dictionary_grade = dict(zip(D2i_list, string_list))
    print(dictionary_grade)
    dictionary_year = dict(zip(year2i_list, year_list))
    print(dictionary_year)

    #----------------------Q1 muti cells updates end------------------------------------------#

    #---------------------------------------Q1 Page Update for CCC Phone Sales To Create new Invoice ----------------------------------#


    #-------------------------Money Owed to Meyer Page  Update for CCC Phone Sales Ends---------------------------------#
    new_file_name = f"{Cust_Obj} Inv {InvNUMBER_object}.xlsx"
    executable_dir = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
    target_dir = executable_dir.parent.parent  # Move up two levels
    new_file_path = target_dir / new_file_name
    shutil.copy(EXCEL_SHEET_PATH, new_file_path)
    
    all_updates = combine_updates(cell_dict_Q1, dictionary_qty, dictionary_grade, dictionary_year)
    
    update_workbook(all_updates, new_file_path)
    print("Saving file to:", new_file_path)


#Make user able to swap the path for excel_sheet_path
    

    
if __name__ == "__main__":
    if len(sys.argv) > 1:
        row_numbers_str = sys.argv[1]
        row_numbers = row_numbers_str.split(',')

        for row_str in row_numbers:
            try:
                row_number = int(row_str)
                main(row_number)
            except ValueError:
                print(f"Invalid row number: {row_str}")
    else:
        print("No row number provided")
