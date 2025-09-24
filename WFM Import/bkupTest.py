import pandas as pd
import os
import warnings

# SUPPRESS SPECIFIC WARNINGS
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

# DEFINE THE FILE PATH
file_path = r'C:\Users\KaitlanBerg\Desktop\GITHUB\SillyScripts\WFM Import\TestDocs\ARJ-CLA-J0006-01-04-CS001.xlsx'

# CHECK IF THE FILE EXISTS
if os.path.exists(file_path):
    # LOAD THE EXCEL FILE, SPECIFYING THAT THE HEADER IS IN THE SEVENTH ROW (INDEX 6)
    df = pd.read_excel(file_path, header=6)

    # PRINT THE ACTUAL COLUMN NAMES TO VERIFY
    print("Actual column names:", df.columns.tolist())

    # STRIP WHITESPACE FROM COLUMN NAMES
    df.columns = df.columns.str.strip()

    # HANDLE DUPLICATE COLUMN NAMES
    df = df.loc[:, ~df.columns.duplicated()]  # KEEP ONLY THE FIRST OCCURRENCE OF EACH COLUMN

    # PRINT THE FIRST FEW ROWS OF THE DATAFRAME TO INSPECT THE DATA
    print("DataFrame preview:")
    print(df.head())

    # CHECK FOR NULL VALUES IN THE RELEVANT COLUMNS
    null_counts = df[['Parts Description', 'List Price (£)']].isnull().sum()
    print("Null value counts:")
    print(null_counts)

    # DEFINE THE REQUIRED COLUMNS BASED ON THE ACTUAL COLUMN NAMES
    required_columns = [
        'Qty', 'Unit', 'Part N°', 'Parts Description', 'Supplier',
        'List Price (£)', 'Nett Total (£)', 'Date Order Issued to Procurement Dept.', 'Status'
    ]

    print(df)
    # FILTER THE DATAFRAME TO ONLY INCLUDE THE REQUIRED COLUMNS
    try:
        df_filtered = df[required_columns]

        # IDENTIFY ROWS WHERE 'PARTS DESCRIPTION' OR 'LIST PRICE (£)' ARE NULL
        null_rows = df_filtered[df_filtered['Parts Description'].isnull() | df_filtered['List Price (£)'].isnull()]

        # PRINT EACH ROW OF THE NULL RESULT SEPARATELY
        print("Rows with null 'Parts Description' or 'List Price (£)':")
        for index, row in null_rows.iterrows():
            print("=" * 50)  # SEPARATOR FOR READABILITY
            print(f"Row {index}:")
            print(row.to_dict())  # PRINT THE ENTIRE ROW AS A DICTIONARY
            print("\n")  # EXTRA SPACE FOR CLARITY

    except KeyError as e:
        print(f"Error: {e}. Please check the column names.")
else:
    print("File does not exist. Please check the path.")
