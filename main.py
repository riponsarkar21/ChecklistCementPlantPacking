import pandas as pd
from tkinter import Tk, Label, Entry, Button, StringVar, OptionMenu, messagebox
from datetime import datetime

EXCEL_FILE = "maintenance_data.xlsx"

def initialize_database():
    try:
        pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Date", "Shift", "Equipment", "Item Checked", "Status", "Remarks", "Checked By"])
        df.to_excel(EXCEL_FILE, index=False)

def add_record(date, shift, equipment, item_checked, status, remarks, checked_by):
    record = {
        "Date": date,
        "Shift": shift,
        "Equipment": equipment,
        "Item Checked": item_checked,
        "Status": status,
        "Remarks": remarks,
        "Checked By": checked_by
    }
    try:
        df = pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        initialize_database()
        df = pd.read_excel(EXCEL_FILE)
    df = df.append(record, ignore_index=True)
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False)
    messagebox.showinfo("Success", "Record added successfully.")

# Initialize Tkinter window
root = Tk()
root.title("Maintenance Entry")

# Fields
Label(root, text="Date").grid(row=0, column=0)
date_entry = Entry(root)
date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
date_entry.grid(row=0, column=1)

Label(root, text="Shift").grid(row=1, column=0)
shift_var = StringVar(root)
shift_var.set("A")
OptionMenu(root, shift_var, "A", "B", "C").grid(row=1, column=1)

Label(root, text="Equipment").grid(row=2, column=0)
equipment_entry = Entry(root)
equipment_entry.grid(row=2, column=1)

Label(root, text="Item Checked").grid(row=3, column=0)
item_entry = Entry(root)
item_entry.grid(row=3, column=1)

Label(root, text="Status").grid(row=4, column=0)
status_var = StringVar(root)
status_var.set("OK")
OptionMenu(root, status_var, "OK", "Not OK").grid(row=4, column=1)

Label(root, text="Remarks").grid(row=5, column=0)
remarks_entry = Entry(root)
remarks_entry.grid(row=5, column=1)

Label(root, text="Checked By").grid(row=6, column=0)
checked_by_entry = Entry(root)
checked_by_entry.grid(row=6, column=1)

# Add Record Button
def submit():
    add_record(
        date=date_entry.get(),
        shift=shift_var.get(),
        equipment=equipment_entry.get(),
        item_checked=item_entry.get(),
        status=status_var.get(),
        remarks=remarks_entry.get(),
        checked_by=checked_by_entry.get()
    )
Button(root, text="Add Record", command=submit).grid(row=7, column=1)

root.mainloop()
