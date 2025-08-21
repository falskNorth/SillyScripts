import pandas as pd
import os
import warnings

# Suppress specific warnings
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

# Define the file path
file_path = r'C:\Users\KaitlanBerg\Desktop\GITHUB\SillyScripts\WFM Import\TestDocs\ARJ-CLA-J0006-01-04-CS001.xlsx'

# Check if the file exists
if os.path.exists(file_path):
    # Load the Excel file, specifying that the header is in the seventh row (index 6)
    df = pd.read_excel(file_path, header=6)

    # Print the actual column names to verify
    print("Actual column names:", df.columns.tolist())

    # Strip whitespace from column names
    df.columns = df.columns.str.strip()

    # Handle duplicate column names
    df = df.loc[:, ~df.columns.duplicated()]  # Keep only the first occurrence of each column

    # Print the first few rows of the DataFrame to inspect the data
    print("DataFrame preview:")
    print(df.head())

    # Check for null values in the relevant columns
    null_counts = df[['Parts Description', 'List Price (£)']].isnull().sum()
    print("Null value counts:")
    print(null_counts)

    # Define the required columns based on the actual column names
    required_columns = [
        'Qty', 'Unit', 'Part N°', 'Parts Description', 'Supplier',
        'List Price (£)', 'Nett Total (£)', 'Date Order Issued to Procurement Dept.', 'Status'
    ]

    # Filter the DataFrame to only include the required columns
    try:
        df_filtered = df[required_columns]

        # Filter rows where 'Parts Description' and 'List Price (£)' are not null
        result = df_filtered[df_filtered['Parts Description'].notnull() & df_filtered['List Price (£)'].notnull()]

        # Print each row of the filtered result separately
        print("Filtered result:")
        for index, row in result.iterrows():
            print("=" * 50)  # Separator for readability
            print(f"Row {index}:")
            print(row.to_dict())  # Print the entire row as a dictionary
            print("\n")  # Extra space for clarity

    except KeyError as e:
        print(f"Error: {e}. Please check the column names.")
else:
    print("File does not exist. Please check the path.")
