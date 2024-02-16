#pyinstaller --onefile --windowed --icon=CCC.ico --add-data="*.xlsx;." GUI.py BORKEN
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import json
from pathlib import Path
from create_invoice import main

# Get the current location of the Python script
script_location = Path(__file__).resolve().parent
config_directory = Path.home() / '.CreateInvoiceApp'
config_directory.mkdir(exist_ok=True)  # Create the directory if it doesn't exist
config_path = config_directory / 'config.json'


def save_config(file_path):
    with open(config_path, 'w') as config_file:
        json.dump({'excel_file_path': file_path}, config_file)

def load_config():
    if config_path.exists():
        with open(config_path, 'r') as config_file:
            config = json.load(config_file)
            return config.get('excel_file_path')
    return None

def parse_row_numbers(input_str):
    """Parse the input string into a list of row numbers, including expanding ranges."""
    row_numbers = []
    for part in input_str.split(','):
        if '-' in part:
            start, end = map(int, part.split('-'))
            row_numbers.extend(range(start, end + 1))  # end + 1 because range is exclusive at the end
        else:
            row_numbers.append(int(part))
    return row_numbers

def prompt_for_file():
    """Prompt the user for the file when the application starts."""
    messagebox.showinfo("Select File", "Please select the CCC phone sales file you have.")
    file_path = load_config()
    if not file_path or not Path(file_path).exists():
        file_path = filedialog.askopenfilename(
            title="Select the CCC Phone Sales Excel file",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if not file_path:
            messagebox.showerror("No file selected", "You must select a file to proceed.")
            root.quit()  # Close the application if no file is selected
            return
        save_config(file_path)
    return file_path

def run_create_invoice():
    input_str = entry.get()
    row_numbers = parse_row_numbers(input_str)
    # Let user choose the directory to save the file
    for i, row_number in enumerate(row_numbers, start=1):
        main(row_number, file_path)  # Pass each row number and save directory to main
        result_label.config(text=f"Processing row: {row_number} ({i}/{len(row_numbers)})")
        root.update()  # This updates the GUI.

    result_label.config(text=f"Executed for rows: {input_str}")

def on_validate(P):
    # Updated regex to include hyphen (-) for ranges like "1-6"
    return re.match("^[0-9,-]*$", P) is not None
    
# Create the main window
root = tk.Tk()
root.title("Fill Excel Sheet")
file_path = prompt_for_file()
# def check_number():
#     # Get the text from the entry widget
#     user_input = entry.get()

#     # You can add your logic here what to do with the number
#     # For now, let's just display it in the result_label
#     result_label.config(text=f"Entered Row Number: {user_input}")
    

# Register the validation command
validate_command = root.register(on_validate), '%P'

# Create a frame for the row number input and the button
frame = tk.Frame(root)
frame.pack(padx=20, pady=10)

# Add a label for row number
label = tk.Label(frame, text="Row number:")
label.pack(side=tk.LEFT)

# Add a textbox for user input
entry = tk.Entry(frame, validate="key", validatecommand=validate_command)
entry.pack(side=tk.LEFT)

# Add a label for instructions
# Add a label for instructions
instruction_label = tk.Label(root, text="Click on White Box enter something like:\n1-6 or 1-6,2-7 or 4,5,6 or 702 in it")
instruction_label.pack(pady=5)  
# Add a button to trigger the validation
check_button = tk.Button(frame, text="Generate", command=run_create_invoice)
check_button.pack(side=tk.LEFT)

# # Add a label for the button description
# button_description_label = tk.Label(button_frame, text="To Save As and Generate Files")
# button_description_label.pack(side=tk.LEFT, padx=5)  # Added next to the button



# Label to display the result of validation
result_label = tk.Label(root, text="")
result_label.pack()

# Start the GUI event loop
root.mainloop()
