import tkinter as tk

def focus_next_widget(event):
    event.widget.tk_focusNext().focus()
    return("break")

app = tk.Tk()
app.title("Invoice Input Form")

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
    for field, entry in entries.items():
        if isinstance(entry, list):
            values = [e.get() for e in entry]
            print(f"{field}: {values}")
        else:
            print(f"{field}: {entry.get()}")  # Here you can process or print the entry data

submit_button = tk.Button(app, text='Submit', command=on_submit)
submit_button.pack(side=tk.RIGHT, padx=5, pady=5)

instruction_label = tk.Label(app, text="Press enter to see something fancy")
instruction_label.pack(side=tk.BOTTOM, pady=5)

if __name__ == "__main__":
    app.mainloop()
