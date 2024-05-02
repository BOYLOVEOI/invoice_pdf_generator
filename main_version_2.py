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
    # getting the file str from filename.stem and splitting it and receiving
    # only the first item, the invoice nr 
    invoice_nr = filename.stem.split('-')[0]

    # creating the pdf document
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_font(family="Times", size=16, style="B")
    pdf.add_page()
    pdf.cell(w=50, h=8, txt="Invoice nr " +invoice_nr)
    pdf.ln(8)
    
    
    # Setting text for the table
    pdf.set_font(family='Times', size=10, style='b')
    # Setting column width of table (dividing width of page by # of columns)
    col_width = pdf.w / len(df.columns)
    # Setting row height of table to size of font 
    row_height = pdf.font_size

    # going over each row in the dataframe to create each row for the table
    for i in df.columns:
        # Setting format for column names
        if i.lower() == 'product_id':
            i = 'Product ID'
        elif i.lower() == 'product_name':
            i = 'Product Name'
        elif i.lower() == 'amount_purchase':
            i = 'Amount Purchased'
        elif i.lower() == 'price_per_unit':
            i = 'Price Per Unit'
        else:
            i = 'Total Price'
        pdf.set_text_color(r=55, g=55, b=55)
        pdf.cell(col_width, row_height, txt=str(i), border=1)
    pdf.ln(row_height)
    pdf.set_text_color(r=0, g=0, b=0)

    pdf.set_font(family='Times', size=10)
    for index, row in df.iterrows():
        for i in (row):
            pdf.cell(w= col_width, h=row_height, txt=str(i), border=1)
        pdf.ln(row_height)

    pdf.output('PDFs/{}.pdf'.format(filename.stem))
    
