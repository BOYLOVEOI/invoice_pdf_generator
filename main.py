# Import Statement
import pandas as pd
import glob

# Creating a list of the file paths for the multiple csv files to be processed
# * is a wildcard that means to get everything that matches .xlsx in the
# /invoices folder
# If we put glob.glob('*') all the directories and files in the current folder
# would be returned
file_paths = glob.glob("invoices/*.xlsx")

# for loop to iterate over and create dataframes for each spreadsheet
for i in file_paths:
    # read_excel is to read in xlsx files
    df = pd.read_excel(i, sheet_name='Sheet 1')
