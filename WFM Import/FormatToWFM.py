import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import sys


class ExcelToCSVConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to CSV Converter (WorkflowMax Import)")

        # Default header row (0-based index, 9 means row 7 in Excel)
        self.header_row = tk.IntVar(value=10)

        self.create_widgets()

    def create_widgets(self):
        # File selection
        tk.Label(self.root, text="Excel File:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        self.xlsx_file_entry = tk.Entry(self.root, width=50)
        self.xlsx_file_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.root, text="Browse", command=self.select_xlsx_file).grid(row=0, column=2, padx=5, pady=5)

        # Output file
        tk.Label(self.root, text="CSV Output:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        self.csv_file_entry = tk.Entry(self.root, width=50)
        self.csv_file_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self.root, text="Browse", command=self.select_csv_output_location).grid(row=1, column=2, padx=5,
                                                                                          pady=5)

        # Header row selection
        tk.Label(self.root, text="Header Row (Excel row number):").grid(row=2, column=0, padx=10, pady=5, sticky='e')
        tk.Entry(self.root, textvariable=self.header_row, width=5).grid(row=2, column=1, padx=5, pady=5, sticky='w')

        # Convert button
        tk.Button(self.root, text="Convert", command=self.convert_to_csv).grid(row=3, column=1, pady=10)

        # Status
        self.status_label = tk.Label(self.root, text="", fg="green")
        self.status_label.grid(row=4, column=0, columnspan=3)

    def select_xlsx_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.xlsx_file_entry.delete(0, tk.END)
            self.xlsx_file_entry.insert(0, file_path)

    def select_csv_output_location(self):
        output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if output_path:
            self.csv_file_entry.delete(0, tk.END)
            self.csv_file_entry.insert(0, output_path)

    def convert_to_csv(self):
        xlsx_file = self.xlsx_file_entry.get()
        csv_file = self.csv_file_entry.get()
        header_row_idx = self.header_row.get() - 1  # Convert to 0-based index

        if not xlsx_file or not csv_file:
            self.show_error("Please select both input and output files")
            return

        try:
            print(f"\nProcessing file: {xlsx_file}")

            # Read Excel ignoring columns A and B (0 and 1)
            df = pd.read_excel(xlsx_file, sheet_name=0, header=None).drop([0, 1], axis=1, errors='ignore')

            required_columns = ['Qty', 'Unit', 'Part N°', 'Parts Description', 'Supplier',
                                'List Price (£)', 'Nett Total (£)', 'Date Order Issued to Procurement Dept.', 'Status']

            print(f"Checking headers in row {header_row_idx + 1}")
            header_row = df.iloc[header_row_idx].values
            print("Header row contents:", header_row)

            # Clean and verify headers
            cleaned_headers = [str(col).strip() for col in header_row]
            missing = [col for col in required_columns if col not in cleaned_headers]

            if missing:
                self.show_error(f"Missing columns in row {header_row_idx + 1}: {', '.join(missing)}")
                return

            # Process data
            df.columns = cleaned_headers
            df = df[header_row_idx + 1:]

            # Date formatting
            date_col = 'Date Order Issued to Procurement Dept.'
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%d/%m/%Y')

            # Filter invalid rows
            df['Delete'] = df['Parts Description'].isnull()
            for col in required_columns[1:]:
                df['Delete'] |= df['Parts Description'].notnull() & df[col].isnull()

            # Create output mapping
            output_map = {
                "Job Number": "",
                "Cost Name": "Parts Description",
                "Date (DD/MM/YYYY)": date_col,
                "Unit Cost": "List Price (£)",
                "Unit Price": "Nett Total (£)",
                "Quantity": "Qty",
                "Phase": "",
                "Supplier": "Supplier",
                "Cost Code": "Part N°",
                "Type": "",
                "Billable": "",
                "Cost Note": "Status",  # Changed to use the Status column
                "Estimated/Actual": ""
            }

            # Create and save CSV
            pd.DataFrame({k: df[v] if v else "" for k, v in output_map.items()}) \
                .loc[~df['Delete']] \
                .to_csv(csv_file, index=False)

            self.show_success(f"Success! Saved {len(df[~df['Delete']])} records to {csv_file}")

        except Exception as e:
            self.show_error(f"Conversion failed: {str(e)}")
            print(f"ERROR: {str(e)}", file=sys.stderr)

    def show_error(self, message):
        self.status_label.config(text=message, fg="red")
        messagebox.showerror("Error", message)

    def show_success(self, message):
        self.status_label.config(text=message, fg="green")
        messagebox.showinfo("Success", message)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToCSVConverter(root)
    root.mainloop()

