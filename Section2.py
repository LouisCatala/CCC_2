import os
from openpyxl import load_workbook

def create_excel_invoice(invoice_data):
    template_path = r"C:\Users\13473\Desktop\CCC\Inovices\Invoice temple_CCC.xlsx"  # Update this path to your template's location
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

    for field in line_item_fields:
        for i, cell in enumerate(line_item_cells[field]):
            ws[cell] = invoice_data[field][i]

    # Calculate Line Totals and Subtotal
    subtotal = 0
    for i, qty in enumerate(invoice_data['QTY']):
        unit_price = invoice_data['Unit Price'][i]
        line_total = qty * unit_price
        ws[line_item_cells['QTY'][i].replace('B', 'H')] = line_total  # Replace 'B' with 'H' in QTY cells for Line Total cells
        subtotal += line_total

    # Write Subtotal, Shipping, and Balance Due
    ws['H38'] = subtotal
    shipping_cost = invoice_data.get('Shipping', 0)  # Use 0 as default if Shipping is not provided
    ws['H39'] = shipping_cost
    ws['H40'] = subtotal + shipping_cost
    
    salesman_name = invoice_data['SOLD TO']
    first_initial = salesman_name.split()[0][0]
    last_name = salesman_name.split()[-1]
    shortened_name = f"{first_initial}. {last_name}"
    # Save the filled-in invoice to a new file
    new_invoice_path = os.path.join(invoices_folder, f"{shortened_name} Inv {invoice_data['INV#']}.xlsx")
    wb.save(new_invoice_path)
    
    # Close the workbook
    wb.close()

    return new_invoice_path  # Returns the path of the saved invoice

# Example usage (you need to replace 'path_to_template/template.xlsx' with your actual template's path and provide the actual data in 'invoice_data'):
# invoice_data = {
#     'INV#': '12345',
#     ... # other fields and values as required
# }
# new_invoice_path = create_excel_invoice(invoice_data)
# print("Invoice saved to:", new_invoice_path)