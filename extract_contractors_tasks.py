# Author: Panos Bonotis -> https://www.linkedin.com/in/panagiotis-bonotis-351a7996/
# Date: Jul-2024
# Description: This program is designed to process a master Excel file containing multiple sheets. 
# It prompts the user to select the master Excel file and an output directory.
# Extracts unique values from the "Όνομα" column in the "Aggregated Data" sheet.
# For each unique value in the "Όνομα" column, creates a new Excel file:
# Includes a sheet named "mail copy/paste" containing specific columns.
# Copies and filters all other sheets based on the unique value in the "Όνομα" column.
# Saves the individual Excel files in the specified output directory.
# This program automates the creation of multiple Excel files based on unique identifiers, 
# making it easier to manage and analyze segmented data.

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from tkinter import ttk
from datetime import datetime

def get_file_path():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

def get_output_folder():
    folder_path = filedialog.askdirectory(title="Select Output Folder")
    return folder_path

def get_date():
    today_str = datetime.today().strftime('%Y-%m-%d')
    date_str = simpledialog.askstring("Input Date", f"Enter the date (YYYY-MM-DD) to filter by 'Ημ/νία Αίτησης':", initialvalue=today_str)
    return date_str

def split_excel_by_unique_names(input_file, output_folder, filter_date):
    # Load the entire Excel file
    xls = pd.ExcelFile(input_file)
    
    # Get the unique values in the "Όνομα" column from the "Aggregated Data" sheet
    df_aggregated = pd.read_excel(input_file, sheet_name='Aggregated Data')
    unique_names = df_aggregated['Όνομα'].dropna().unique()  # Drop NaN values
    
    # Columns to include in the "mail copy/paste" sheet
    mail_copy_columns = [
        'SR ID', 'Τύπος εργασίας', 'Κατάσταση', 'Ημ/νία Αίτησης', 'Διεύθυνση πελάτη', 
        'Αριθμός Οδού', 'BUILDING ID', 'A/K', 'ΤΗΛΕΦΩΝΟ ΠΑΡΑΓΓΕΛΙΑΣ', 'FLOOR', 
        'PILOT', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ', 'ΚΙΝΗΤΟ ΠΕΛΑΤΗ', 'ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΔΙΑΧΕΙΡΙΣΤΗ', 'ΚΙΝΗΤΟ ΔΙΑΧΕΙΡΙΣΤΗ'
    ]
    
    # Convert filter_date to datetime for comparison
    filter_date = pd.to_datetime(filter_date).date()
    
    # Iterate over each unique name and create a new Excel file
    for name in unique_names:
        file_name = f'{name}_{filter_date}.xlsx'
        with pd.ExcelWriter(os.path.join(output_folder, file_name)) as writer:
            # Create the "mail copy/paste" sheet
            mail_copy_df = df_aggregated[(df_aggregated['Όνομα'] == name) & (pd.to_datetime(df_aggregated['Ημ/νία Αίτησης']).dt.date == filter_date)][mail_copy_columns]
            mail_copy_df.to_excel(writer, sheet_name='mail copy&paste', index=False)
            
            # Store the relevant SR IDs
            relevant_sr_ids = set(mail_copy_df['SR ID'])
            
            # Iterate over all sheets and create filtered sheets in the new Excel file
            for sheet_name in xls.sheet_names:
                if sheet_name == 'Last Drop':
                    continue
                df = pd.read_excel(input_file, sheet_name=sheet_name)
                if 'Όνομα' in df.columns:
                    date_column = 'Ημ/νία Αίτησης'
                    if date_column in df.columns:
                        # Filter the rows that match the current name and date
                        filtered_df = df[(df['Όνομα'] == name) & (pd.to_datetime(df[date_column]).dt.date == filter_date)]
                    else:
                        filtered_df = df[df['Όνομα'] == name]
                else:
                    filtered_df = df
                
                # Filter SR IDs for the "Summary of Actions" sheet if it exists
                if sheet_name == 'Summary of Actions':
                    filtered_df = filtered_df[filtered_df['SR ID'].isin(relevant_sr_ids)]
                
                # Write the filtered DataFrame to the new Excel file
                filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Created {os.path.join(output_folder, file_name)}")

def run_splitter():
    input_file = file_path.get()
    output_folder = output_folder_path.get()
    filter_date = date_entry.get()
    
    if not input_file or not output_folder or not filter_date:
        messagebox.showerror("Error", "Please select the input file, output folder, and enter the date.")
        return
    
    try:
        # Ensure the date is in the correct format
        filter_date = pd.to_datetime(filter_date).date()
    except ValueError:
        messagebox.showerror("Error", "Invalid date format. Please enter the date in YYYY-MM-DD format.")
        return
    
    try:
        # Split the Excel file by unique names and filter by date
        split_excel_by_unique_names(input_file, output_folder, filter_date)
        messagebox.showinfo("Success", "Files created successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the main application window
root = tk.Tk()
root.title("Excel Splitter")

# Create and set variables
file_path = tk.StringVar()
output_folder_path = tk.StringVar()
date_entry = tk.StringVar()

# Create and place widgets
tk.Label(root, text="Select Excel File:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=file_path, width=50).grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=lambda: file_path.set(get_file_path())).grid(row=0, column=2, padx=10, pady=5)

tk.Label(root, text="Select Output Folder:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=output_folder_path, width=50).grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=lambda: output_folder_path.set(get_output_folder())).grid(row=1, column=2, padx=10, pady=5)

tk.Label(root, text="Enter Date (YYYY-MM-DD):").grid(row=2, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=date_entry, width=50).grid(row=2, column=1, padx=10, pady=5)

tk.Button(root, text="Run", command=run_splitter).grid(row=3, column=1, padx=10, pady=20)

# Start the application
root.mainloop()