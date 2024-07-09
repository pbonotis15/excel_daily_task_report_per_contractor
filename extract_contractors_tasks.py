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
from tkinter import filedialog, simpledialog
from datetime import datetime

def get_file_path():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

def get_output_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory(title="Select Output Folder")
    return folder_path

def get_date():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
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
        with pd.ExcelWriter(os.path.join(output_folder, f'{name}.xlsx')) as writer:
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
            print(f"Created {os.path.join(output_folder, f'{name}.xlsx')}")

if __name__ == "__main__":
    # Prompt the user to select the input file
    input_file = get_file_path()
    
    if not input_file:
        print("No file selected.")
        exit()

    # Prompt the user to select the output folder
    output_folder = get_output_folder()
    
    if not output_folder:
        print("No output folder selected.")
        exit()
    
    # Prompt the user to enter the date for filtering, prefilled with today's date
    filter_date = get_date()
    
    if not filter_date:
        print("No date entered.")
        exit()

    try:
        # Ensure the date is in the correct format
        filter_date = pd.to_datetime(filter_date).date()
    except ValueError:
        print("Invalid date format. Please enter the date in YYYY-MM-DD format.")
        exit()
    
    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)
    
    # Split the Excel file by unique names and filter by date
    split_excel_by_unique_names(input_file, output_folder, filter_date)