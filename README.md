# Excel Daily task report creator per Contractor

## Overview

This is a Python-based tool designed to automate the segmentation of a master Excel file into multiple individual files. Each individual file corresponds to a unique value in the "Όνομα" column of the "Aggregated Data" sheet. This tool also creates a special sheet titled "mail copy/paste" in each output file, containing specific columns for easy copying and pasting.

## Features

- Prompts the user to select the input Excel file and output directory.
- Extracts unique values from the "Όνομα" column.
- Creates individual Excel files for each unique value.
- Includes a "mail copy/paste" sheet with specified columns in each output file.
- Copies and filters all other sheets based on the unique value in the "Όνομα" column.

## Requirements

- Python 3.x
- pandas
- openpyxl
- tkinter

## Installation

1. Clone this repository or download the script files.
    ```sh
    git clone https://github.com/pbonotis15/excel-data-merger-and-summary-based-on-IDs.git
    cd your-repo-name
    ```

2. Install the required Python libraries:
    ```bash
    pip install pandas openpyxl
    ```

## Usage

1. Run the `extract_contractors_tasks.py` script:
    ```bash
    python extract_contractors_tasks.py
    ```
2. When prompted, select the master Excel file (`final_results.xlsx`).
3. Select the output directory where the individual Excel files will be saved.
4. The script will generate Excel files for each unique value in the "Όνομα" column (Contractors names), each containing the specified columns and filtered sheets.

## Contributing

If you'd like to contribute to this project, please fork the repository and use a feature branch. Pull requests are warmly welcome.