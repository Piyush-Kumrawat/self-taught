import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import subprocess
import sys
import openpyxl
import numpy as np

def run_script(input_file, output_file):
    # Read the original Excel file
    df_original = pd.read_excel(input_file)

    # Create a new Excel writer object with the xlsxwriter engine
    excel_writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

    # Write the selected columns to sheet 1
    selected_columns = ['csid', 'SurveyDate', 'orid', 'Comments']
    df_original[selected_columns].to_excel(excel_writer, sheet_name='Sheet1', index=False)

    # Identify columns with integer values from the 5th column onwards
    numeric_columns = df_original.iloc[:, 4:].select_dtypes(include='number').columns

    # Create an empty DataFrame for the final result
    df_final = pd.DataFrame()

    # Iterate over non-empty columns from the 5th column onwards
    for idx, col_name in enumerate(df_original.columns[4:], start=1):
        if df_original[col_name].notna().any():
            # Filter rows with non-null values in the iterating column
            filtered_rows = df_original[df_original[col_name].notna()]

            # Create a copy of the original DataFrame with filtered rows
            df_iteration = filtered_rows.copy()

            # Rename the column for 'Response Code' based on the current iteration
            df_iteration.rename(columns={col_name: 'Response Code'}, inplace=True)

            # Add 'Priority' column with values based on the iteration
            df_iteration['Priority'] = idx

            # Concatenate the current iteration to the final result
            df_final = pd.concat([df_final, df_iteration[['csid', 'SurveyDate', 'orid', 'Comments', 'Response Code', 'Priority']]])

    # Custom sorting function to handle both numeric and alphanumeric values
    def custom_sort(x):
        try:
            return float(x)
        except (ValueError, TypeError):
            return x

    # Sort the final DataFrame with the custom sorting function by 'csid'
    df_final['csid'] = df_final['csid'].apply(custom_sort)
    df_final = df_final.sort_values(by='csid', key=lambda x: x.astype(str))

    # Write the final DataFrame to sheet 2
    df_final.to_excel(excel_writer, sheet_name='Sheet2', index=False)

    # Close the Excel writer to save the file
    excel_writer.close()

def run_gui():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Ask for input and output file paths using file dialogs
    input_file = filedialog.askopenfilename(title="Select Input Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not input_file:
        messagebox.showinfo("Info", "Please select an input file.")
        return

    output_file = filedialog.asksaveasfilename(title="Save Output Excel File", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_file:
        messagebox.showinfo("Info", "Please select an output file.")
        return

    try:
        # Run the data processing script
        run_script(input_file, output_file)
        messagebox.showinfo("Info", "Script completed.")
    except Exception as e:
        messagebox.showerror("Error", f"Error running script: {str(e)}")

if __name__ == "__main__":
    run_gui()