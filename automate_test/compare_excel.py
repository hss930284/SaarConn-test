import pandas as pd

import tkinter as tk

from tkinter import filedialog

# Function to get file path from user

def get_file_path(prompt):

    root = tk.Tk()

    root.withdraw()

    return filedialog.askopenfilename(title=prompt, filetypes=[("Excel files", "*.xlsx *.xls")])

# Get Excel files from user

file1_path = get_file_path("Select the Original Excel File")

file2_path = get_file_path("Select the Updated Excel File")

# Get output log file paths

excel_log_path = filedialog.asksaveasfilename(title="Save Log File As (Excel)", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

text_log_path = filedialog.asksaveasfilename(title="Save Log File As (Text)", defaultextension=".txt", filetypes=[("Text files", "*.txt")])

# Read Excel files

xls1 = pd.ExcelFile(file1_path)

xls2 = pd.ExcelFile(file2_path)

diff_data = []

# Compare sheets

for sheet_name in xls1.sheet_names:

    if sheet_name in xls2.sheet_names:

        df1 = xls1.parse(sheet_name)

        df2 = xls2.parse(sheet_name)

        # Ensure same shape (fill missing cells with empty string to prevent NaN issues)

        df1, df2 = df1.align(df2, fill_value='')

        # Iterate over each cell

        for row in range(df1.shape[0]):

            for col in range(df1.shape[1]):

                val1 = str(df1.iat[row, col]).strip()

                val2 = str(df2.iat[row, col]).strip()

                # Log only if values are different

                if val1 != val2:

                    cell_ref = f"{chr(65 + col)}{row + 1}"  # Convert column index to Excel column letter

                    diff_data.append([sheet_name, cell_ref, val1, val2])

# Save to Excel

if diff_data:

    diff_df = pd.DataFrame(diff_data, columns=["Sheet", "Cell", "Original Value", "Updated Value"])

    diff_df.to_excel(excel_log_path, index=False)

    # Save to Text File

    with open(text_log_path, "w") as txt_file:

        for sheet, cell, original, updated in diff_data:

            txt_file.write(f"Sheet: {sheet}, Cell: {cell}\n")

            txt_file.write(f"  - Original Value: {original}\n")

            txt_file.write(f"  - Updated Value: {updated}\n")

            txt_file.write("-" * 50 + "\n")

    print(f"Changes saved in:\nExcel: {excel_log_path}\nText: {text_log_path}")

else:

    print("No differences found.") 