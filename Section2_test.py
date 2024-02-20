from Section2 import create_excel_invoice

# This replicates the structure of sample_data from Section1_test.py
invoice_data = {
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
    'QTY': [10, 20, 30, 40, 50],  # Ensure numeric values for calculations
    'DATE(Year)': ['2020', '2021', '2022', '2023', '2024'],
    'GRADE': ['A', 'B', 'C', 'D', 'E'],
    'Description': ['Item A', 'Item B', 'Item C', 'Item D', 'Item E'],
    'Unit Price': [100, 200, 300, 400, 500],  # Ensure numeric values for calculations
    'Shipping': 100,  # Assuming shipping cost is a numeric value
    'Balance Due': '2500'  # This might be calculated based on other values, ensure it's correct or recalculated as needed
}

# Call the function to create an Excel invoice using the sample data
new_invoice_path = create_excel_invoice(invoice_data)

# Print the path of the saved invoice for confirmation
print(f"Invoice saved to: {new_invoice_path}")
