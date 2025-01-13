<h1>OT16 - Additional Calculations Model - Aggregagtion</h1>

This repository reflects the code used for OT16 Additional Calculations Model - Aggregation. It reflects the final set of code use for FD. The path for this is: 

...\Data\Python work\Other\OT16 Additional Calculations Model\OT16a V9 - SBB split - FINAL

Please note: 
* It does not cover the OT16 calculations python code. These are a set of very specific calculations which are of limited value elsewhere.
* the .gitignore file ensures that the excel and text files are not included in this repository.

**1. Overview**
This Python script is designed to:

Import company level data for the components for the cost data of the key datasets
Import the aggregated data from OT16 as marked in spreadsheet "PR24OT16_Copy_1.3".
Automatically calculate and aggregate data for specific years (2025-30).
Generate an f_output file to update the Fountain system.
Create pivot tables for QA purposes.
Ensure only aggregated values are added to the output, avoiding double-counting.
The script reads data from an Excel file, processes it, and outputs it back into the same file with aggregated results and pivot tables for quality assurance.

**2. Purpose**
Input Excel File: The script assumes you have an Excel file structured with the following sheets:
F_Inputs: Contains the key data for cost components.
F_Inputs_APR: Contains additional data for aggregation.
The columns should be in a format that includes company acronyms, cost data for various years (e.g., 2025-26, 2026-27, etc.), and references. Data should be cleaned and free from duplicates.

Run the Script: The script will open a file dialog to let you select the input Excel file.
The script will automatically process the data, remove duplicates, calculate aggregated data for several groups (e.g., ENG, WAL, IND, WASC, WOC), and generate the output.
It will then generate pivot tables for the years 2025-26 to 2029-30 and append them to the same Excel file.

Output File: The results are saved back into the same Excel file you selected as input.
Aggregated data will be placed in a new sheet called F_Outputs.
Pivot tables for the years 2025-26 to 2029-30 will be added in separate sheets.
A timestamp of when the f_output was generated will also be included in the F_Outputs sheet. 

**3. Methodology**
Data Cleaning: It removes duplicates from the data to prevent double-counting.
Aggregated Data Calculation: The script calculates aggregated data for different regions and sectors (ENG, WAL, IND, WASC, WOC).
The data for each group is summed across different companies in the group.
Aggregated data is then added to the output file under the sheet F_Outputs.
Pivot Table Generation: Pivot tables are generated for each year (2025-26 to 2029-30).
The pivot tables summarize the data for each company and the calculated aggregates.
Output to Excel: The final results are written back into the Excel file with the updated data.
Pivot tables are added to the Excel file in separate sheets for each year.
Metadata such as the script path and execution timestamp is inserted into the F_Outputs sheet.

**4. Important Notes**
The aggregated data is calculated based on specific groups of companies (e.g., ENG, WAL, etc.), so the dataset must have the appropriate company acronyms for this to work correctly.
The script will automatically drop any pre-existing aggregated data in the dataset to prevent double counting.
The generated Excel file will be overwritten if it already exists, so make sure to backup the original file before running the script.

**5. License** This code is published under the Open Government Licence v3.0. You are encouraged to use and reuse the information provided here, subject to the terms and conditions of the Open Government Licence

**6. Author** Alex Whitmarsh alex.whitmarsh@ofwat.gov.uk



