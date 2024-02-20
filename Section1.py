import tkinter as tk
from Section2 import create_excel_invoice 

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
    invoice_data = {}
    for field, entry in entries.items():
        if isinstance(entry, list):  # For fields with multiple entries
            invoice_data[field] = [e.get() for e in entry]
        else:
            invoice_data[field] = entry.get()
    
    # Assuming create_excel_invoice is properly imported or defined in this script
    try:
        new_invoice_path = create_excel_invoice(invoice_data)
        print(f"Invoice saved to: {new_invoice_path}")
        # Optionally, you can show a success message to the user here
    except Exception as e:
        print(f"Error generating invoice: {e}")
        # Optionally, handle or show the error to the user here

submit_button = tk.Button(app, text='Submit', command=on_submit)
submit_button.pack(side=tk.RIGHT, padx=5, pady=5)

if __name__ == "__main__":
    app.mainloop()
