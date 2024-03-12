# Make a better code, consider goal, the process,
# after (the possible changes required), Connect with acutal situation. 
# pyinstaller --onefile --windowed --add-data "Invoice temple_CCC.xlsx;." -n Create_Invoice Section1.py
import os
from openpyxl import load_workbook

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the path to the Excel template using a relative path
template_path = os.path.join(script_dir, 'Invoice temple_CCC.xlsx')

def create_excel_invoice(invoice_data):
    global formatted_data_2        
    # template_path = r"C:\Users\13473\Desktop\CCC\Inovices\Invoice temple_CCC.xlsx"  # Update this path to your template's location
    invoices_folder = os.path.expanduser("~/Desktop/invoices")
    if not os.path.exists(invoices_folder):
        os.makedirs(invoices_folder)

    # Load the workbook and select the active worksheet
    wb = load_workbook(template_path)
    ws = wb.active

    # Fill in the details in the Excel sheet
    ws['H4'] = ws['C13'] = ws['D13'] = invoice_data['INV#']
    ws['C6'] = ws['F6'] = invoice_data['SOLD TO']
    ws['H3'] = invoice_data['DATE']
    ws['C7'] = ws['F7'] = invoice_data['Address']
    ws['B13'] = invoice_data['SALESMAN']
    ws['C8'] = invoice_data['City']
    ws['F8'] = invoice_data['City']
    ws['D8'] = invoice_data['State'] + ' ' + invoice_data['ZIP']  # Assuming State and ZIP should be in the same cell separated by a space
    ws['G8'] = invoice_data['State'] + ' ' + invoice_data['ZIP'] 
    ws['E13'] = invoice_data['Shipping Terms']
    ws['F13'] = invoice_data['Amount down']
    ws['G13'] = invoice_data['Payment Method']
    ws['H13'] = invoice_data['Due Date']

    # Fill in the line items
    line_item_fields = ['QTY', 'DATE(Year)', 'GRADE', 'Description', 'Unit Price']
    line_item_cells = {
        'QTY': ['B17', 'B19', 'B21', 'B23', 'B25'],
        'DATE(Year)': ['C17', 'C19', 'C21', 'C23', 'C25'],
        'GRADE': ['D17', 'D19', 'D21', 'D23', 'D25'],
        'Description': ['E17', 'E19', 'E21', 'E23', 'E25'],
        'Unit Price': ['F17', 'F19', 'F21', 'F23', 'F25']
    }

    # Before processing line items, find the minimum length of the provided lists to avoid "index out of range" errors
    min_line_items_length = min(len(invoice_data[field]) for field in line_item_fields)

    for field in line_item_fields:
        for i in range(min_line_items_length):
            value = invoice_data[field][i]
            cell = line_item_cells[field][i]
            if value:  # Check if there's actually a value to avoid writing empty strings for missing data
                ws[cell] = value


    # Calculate Line Totals and Subtotal
    # Calculate Line Totals and Subtotal
    subtotal = 0
    for i, qty_str in enumerate(invoice_data['QTY']):
        # Check if qty_str is not empty, else default to 0
        qty = int(qty_str) if qty_str else 0
        
        unit_price_str = invoice_data['Unit Price'][i]
        # Check if unit_price_str is not empty, else default to 0.0
        unit_price = float(unit_price_str) if unit_price_str else 0.0
        
        line_total = qty * unit_price
        if line_total != 0:  # Only write line_total if it's not 0
            ws[line_item_cells['QTY'][i].replace('B', 'H')] = line_total
        subtotal += line_total

    # Write Subtotal, Shipping, and Balance Due only if they are not 0
    if subtotal != 0:
        ws['H38'] = subtotal

    try:
        shipping_str = invoice_data.get('Shipping', '0')  # Default to '0' if not provided
        shipping_cost = float(shipping_str) if shipping_str else 0.0
    except ValueError:
        shipping_cost = 0.0  # Default to 0.0 if conversion fails
    try:
        amount_down_str = invoice_data.get('Amount down', '0')  # Default to '0' if not provided
        amount_down = float(amount_down_str) if amount_down_str else 0.0
    except ValueError:
        amount_down = 0.0  # Default to 0.0 if conversion fails
    if shipping_cost != 0.0:
        ws['H39'] = shipping_cost

    balance_due = subtotal - amount_down
    if balance_due != 0.0:
        ws['H40'] = balance_due

    costumer_name = invoice_data['SOLD TO']
    first_initial = costumer_name.split()[0][0]
    last_name = costumer_name.split()[-1]
    shortened_name = f"{first_initial}. {last_name}"
    # Save the filled-in invoice to a new file
    new_invoice_path = os.path.join(invoices_folder, f"{shortened_name} Inv {invoice_data['INV#']}.xlsx")
    wb.save(new_invoice_path)
    
    # Close the workbook
    wb.close()
    details = [f"({qty}) {grade} {desc}" for qty, grade, desc in zip(invoice_data['QTY'], invoice_data['GRADE'], invoice_data['Description'])]
    combined_string = ' (), '.join(details)
    formatted_data_2 = f"{invoice_data['DATE']}\t{invoice_data['INV#']}\t{shortened_name}\t{invoice_data['SALESMAN']}\t\"{combined_string}\"" 
    return new_invoice_path, formatted_data_2 

# Example usage (you need to replace 'path_to_template/template.xlsx' with your actual template's path and provide the actual data in 'invoice_data'):
# invoice_data = {
#     'INV#': '12345',
#     ... # other fields and values as required
# }
# new_invoice_path = create_excel_invoice(invoice_data)
# print("Invoice saved to:", new_invoice_path)
import tkinter as tk

def focus_next_widget(event):
    event.widget.tk_focusNext().focus()
    return("break")
    
app = tk.Tk()
app.title("Invoice Input Form")

instruction_label = tk.Label(app, text="Press Tab/ enter go Next box. After click on text box, \n if needed, Ctrl+A to select infos. Ctrl+C to copy, Ctrl+V to paste infos")
instruction_label.pack(side=tk.TOP, pady=5)

# Define the regular fields
regular_fields = [
    'INV#', 'SOLD TO', 'DATE', 'Address', 'SALESMAN', 
    'City', 'State', 'ZIP', 'Shipping Terms', 'Amount down', 
    'Payment Method', 'Due Date'
]

# Define the fields that need multiple entries on the same line
multi_entry_fields = [
    ('QTY', 5),
    ('DATE(Year)', 5),
    ('GRADE', 5),
    ('Description', 5),
    ('Unit Price', 5)
]
under_multi_entry_fields = {
    'Shipping', 'Balance Due'
}
entries = {}
for field in regular_fields:
    row = tk.Frame(app)
    label = tk.Label(row, width=22, text=field+": ", anchor='w')
    entry = tk.Entry(row)
    entry.bind("<Return>", focus_next_widget)
    row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
    label.pack(side=tk.LEFT)
    entry.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
    entries[field] = entry

# Handling multiple entries
for field, num in multi_entry_fields:
    row = tk.Frame(app)
    tk.Label(row, width=22, text=field+": ", anchor='w').pack(side=tk.LEFT)
    row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

    # Create the specified number of entries for each field
    entries[field] = []
    for i in range(num):
        entry = tk.Entry(row)
        entry.bind("<Return>", focus_next_widget)
        entry.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, padx=2)
        entries[field].append(entry)
        
#Handling under_multi_entry_field
for field in under_multi_entry_fields:
    row = tk.Frame(app)
    label = tk.Label(row, width=22, text=field+": ", anchor='w')
    entry = tk.Entry(row)
    entry.bind("<Return>", focus_next_widget)
    row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
    label.pack(side=tk.LEFT)
    entry.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
    entries[field] = entry
    
def on_submit():
    global global_invoice_data
    invoice_data = {}
    for field, entry in entries.items():
        if isinstance(entry, list):  # For fields with multiple entries
            invoice_data[field] = [e.get() for e in entry]
        else:
            invoice_data[field] = entry.get()
    global_invoice_data = invoice_data
    
    # Assuming create_excel_invoice is properly imported or defined in this script
    try:
        new_invoice_path, formatted_data_2 = create_excel_invoice(invoice_data)
        print(f"Invoice saved to: {new_invoice_path}")
        # Optionally, you can show a success message to the user here
    except Exception as e:
        print(f"Error generating invoice: {e}")
        # Optionally, handle or show the error to the user here
    def copy_to_clipboard():
        app.clipboard_clear()  # Clear the clipboard
        app.clipboard_append(formatted_data_2)  # Append the formatted data to the clipboard
        print("Data copied to clipboard.")
#Copy Button for copy and paste
    copy_button = tk.Button(app, text='Copy', command=copy_to_clipboard)
    copy_button.pack(side=tk.RIGHT, padx=0, pady=5)

#submit button
submit_button = tk.Button(app, text='Submit', command=on_submit)
submit_button.pack(side=tk.RIGHT, padx=5, pady=5)


if __name__ == "__main__":
    app.mainloop()

