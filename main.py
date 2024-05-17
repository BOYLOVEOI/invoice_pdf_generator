# Import Statement
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Creating a list of the file paths for the multiple csv files to be processed
# * is a wildcard that means to get everything that matches .xlsx in the
# /invoices folder
# If we put glob.glob('*') all the directories and files in the current folder
# would be returned
file_paths = glob.glob('invoices/*.xlsx')


# for loop to iterate over and create dataframes for each spreadsheet
for i in file_paths:
    # read_excel is to read in xlsx files
    df = pd.read_excel(i, sheet_name='Sheet 1')
    # getting the file name ready dynamically through path.stem
    filename = Path(i)
    # splitting the filename and unpacking the two elements, the invoice number 
    # and date into two variables, invoice_nr and date 
    invoice_nr, date = filename.stem.split('-')

    # setting the format of the pdf document
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_font(family="Times", size=16, style="B")
    # adding the first page
    pdf.add_page()
    # adding the title text cell
    pdf.cell(w=50, h=8, txt="Invoice N.{}".format(invoice_nr), ln=1)
    # adding the date below the title
    pdf.cell(w=50, h=8, txt="Date: {}".format(date))

    # outputting each newly created pdf with their own names to the PDF folder
    pdf.output("PDFS/{}.pdf".format(filename.stem))


    