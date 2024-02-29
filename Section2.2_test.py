# Make a better code, consider goal, the process,
# after (the possible changes required), Connect with acutal situation. 
# pyinstaller --onefile --windowed --add-data "Invoice temple_CCC.xlsx;." -n Create_Invoice Section1.py
import os
from openpyxl import load_workbook
import random
import tkinter as tk
#this should be inside section 2 function, and filled with values define before and assign into the function
#So nothing need to assigned in parameter
date = "1/9/2024" #inovce data
id = '7035' #inv number
name = "K.Moran" #costimer name
salesman = "David" #invoice salesman 
grade = []#invoice_data['GRADE'] 
qty_list = [random.randint(1, 100) for _ in range(5)]
grade_list = [random.choice(['A', 'B', 'C', 'D']) for _ in range(5)]
desc_list = [random.choice(['Widget', 'Gadget', 'Tool', 'Item']) + f" {i+1}" for i in range(5)]

details = [f"({qty}) {grade} {desc}" for qty, grade, desc in zip(qty_list, grade_list, desc_list)]
print(details)
combined_string = ' (), '.join(details)
print(combined_string)
def format_data_for_excel():
    formatted_data = f"{date}\t{id}\t{name}\t{salesman}\t\"{combined_string}\""
    return formatted_data
def copy_to_clipboard():
    formatted_data = format_data_for_excel()
    app.clipboard_clear()  # Clear the clipboard
    app.clipboard_append(formatted_data)  # Append the formatted data to the clipboard
    print("Data copied to clipboard.")


# Example usage
data_to_paste = format_data_for_excel()
print(data_to_paste)
app = tk.Tk()

copy_button = tk.Button(app, text='Copy', command=copy_to_clipboard)
copy_button.pack(side=tk.RIGHT, padx=5, pady=5)

app.mainloop()

