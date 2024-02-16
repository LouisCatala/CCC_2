import tkinter as tk
from Section1 import app, entries  # Assuming your main file is named section1.py

# Define a dictionary with sample data for each field
sample_data = {
    'INV#': '12345',
    'SOLD TO': 'Company ABC',
    'DATE': '02/16/2024',
    'Address': '1234 Main St',
    'SALESMAN': 'John Doe',
    'City': 'Anytown',
    'State': 'Anystate',
    'ZIP': '12345',
    'Shipping Terms': 'FOB Destination',
    'Amount down': '500',
    'Payment Method': 'Credit Card',
    'Due Date': '03/16/2024',
    'QTY': ['10', '20', '30', '40', '50'],
    'DATE(Year)': ['2020', '2021', '2022', '2023', '2024'],
    'GRADE': ['A', 'B', 'C', 'D', 'E'],
    'Description': ['Item A', 'Item B', 'Item C', 'Item D', 'Item E'],
    'Unit Price': ['100', '200', '300', '400', '500'],
    'Shipping': 'Standard',
    'Balance Due': '2500'
}

def fill_fields_with_sample_data():
    # Iterate over the entries and set the sample data for each field
    for field, entry in entries.items():
        if isinstance(entry, list):
            for i, e in enumerate(entry):
                e.insert(0, sample_data[field][i])
        else:
            entry.insert(0, sample_data[field])

# Call the function to fill the fields with sample data
fill_fields_with_sample_data()

# Start the tkinter loop
app.mainloop()
