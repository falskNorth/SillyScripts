import tkinter as tk
from tkinter import filedialog
import pandas as pd


def select_xlsx_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        xlsx_file_entry.delete(0, tk.END)  # Clear the entry field
        xlsx_file_entry.insert(0, file_path)  # Insert the selected file path


def select_csv_output_location():
    output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if output_path:
        csv_file_entry.delete(0, tk.END)  # Clear the entry field
        csv_file_entry.insert(0, output_path)  # Insert the selected output path


def convert_to_csv():
    xlsx_file = xlsx_file_entry.get()
    csv_file = csv_file_entry.get()

    if xlsx_file and csv_file:
        try:
            # Read the Excel file
            df = pd.read_excel(xlsx_file)
            # Save as CSV
            df.to_csv(csv_file, index=False)
            status_label.config(text="Conversion successful!", fg="green")
        except Exception as e:
            status_label.config(text=f"Error: {e}", fg="red")
    else:
        status_label.config(text="Please select both files.", fg="red")


# Create the main window
root = tk.Tk()
root.title("Excel to CSV Converter (WorkflowMax Imports)")

# Create and place the widgets
tk.Label(root, text="Select .xlsx file:").grid(row=0, column=0, padx=10, pady=10)
xlsx_file_entry = tk.Entry(root, width=50)
xlsx_file_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_xlsx_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Save as .csv file:").grid(row=1, column=0, padx=10, pady=10)
csv_file_entry = tk.Entry(root, width=50)
csv_file_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_csv_output_location).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Convert", command=convert_to_csv).grid(row=2, column=1, pady=20)

status_label = tk.Label(root, text="")
status_label.grid(row=3, column=0, columnspan=3)

# Start the main loop
root.mainloop()
