import pandas as pd
import os

# DEFINE HEADERS FOR THE CVS FILE AND CREATE df WITH IT
headers = [
    "Job Number", "Cost Name", "Date (DD/MM/YYYY)", "Unit Cost",
    "Unit Price", "Quantity", "Phase", "Supplier",
    "Cost Code", "Type", "Billable", "Cost Note", "Estimated/Actual"
]

df = pd.DataFrame(columns=headers)

# CREATES Output DIR IF ONE DOESNT EXIST
output_dir = "../Output"
os.makedirs(output_dir, exist_ok=True)

# DEFINE BASE FILE NAME/ INITIAL FILE PATH
base_file_name = "JOB PO"
file_extension = ".csv"
file_path = os.path.join(output_dir, base_file_name + file_extension)

# CHECK FOR EXISTING FILES AND CREATE NEW UNIQUE FILE NAME IF NEEDED
n = 1
while os.path.exists(file_path):
    file_path = os.path.join(output_dir, f"{base_file_name}_{n}{file_extension}")
    n += 1

# Save the DataFrame to a CSV file in the Output folder
df.to_csv(file_path, index=False)

print(f"CSV file '{os.path.basename(file_path)}' created successfully in the '{output_dir}' folder.")
