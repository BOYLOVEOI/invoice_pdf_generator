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
file_paths = glob.glob("invoices/*.xlsx")


# for loop to iterate over and create dataframes for each spreadsheet
for i in file_paths:

    # getting the file name ready dynamically through path.stem
    filename = Path(i)
    # splitting the filename and unpacking the two elements, the invoice number 
    # and date into two variables, invoice_nr and date 
    invoice_nr, date = filename.stem.split("-")

    # setting the format of the pdf document
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_font(family="Times", size=16, style="B")
    # adding the first page
    pdf.add_page()
    # adding the title text cell
    pdf.cell(w=50, h=8, txt="Invoice N.{}".format(invoice_nr), ln=1)
    # adding the date below the title
    pdf.cell(w=50, h=8, txt="Date: {}".format(date), ln=1)
    # adding in an extra break for the table
    pdf.ln(h=8)

    # reading in the excel files
    # read_excel is to read in xlsx files
    df = pd.read_excel(i, sheet_name="Sheet 1")
    
    # storing the column names for the headers
    column_names = df.columns
    # list comprehension to reformat each column name
    column_names = [i.replace("_", " ").title() for i in column_names]

    # creating the column headers
    # setting font for the table content
    pdf.set_font(family="Times", size=10, style="B")
    # setting color of the font RGB of (80,80,80) = gray
    pdf.set_text_color(r=80, g=80, b=80)
    # adding in the first 'block' of the table, the product_id header
    pdf.cell(w=30, h=8, txt=column_names[0], border=True)
    # adding in the second 'block' of the table, the product_name header
    pdf.cell(w=70, h=8, txt=column_names[1], border=True)
    # adding in the third 'block', the amount_purchased header
    pdf.cell(w=35, h=8, txt=column_names[2], border=True)
    # adding in the fourth 'block' the price_per_unit header
    pdf.cell(w=30, h=8, txt=column_names[3], border=True)
    # adding in the last 'block' the total_price header
    # adding in a ln argument for the last 'block' as that is where the
    # row in the invoice table is completed/ended
    pdf.cell(w=30, h=8, txt=column_names[4], border=True, ln=1)

    # reading each row in the dataframe and adding it into the invoice
    # table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(r=80, g=80, b=80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=True)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=True)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=True)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=True)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=True, ln=1)

    # Creating the total price to add to the last row
    total_price = df["total_price"].sum()
    # Creating the last row for the total (only ONE row so OUTSIDE for loop)
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, border=True)
    pdf.cell(w=70, h=8, border=True)
    pdf.cell(w=35, h=8, border=True)
    pdf.cell(w=30, h=8, border=True) 
    pdf.cell(w=30, h=8, txt=str(total_price), border=True, ln=1)

    # Adding spacer between table and footer lines
    pdf.ln()

    # Adding the total price footer line
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(r=0, g=0, b=0)
    pdf.cell(w=30, h=8, txt=f"The total price is {total_price}", ln=1)
    
    # Adding the company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.set_text_color(r=0, g=0, b=0)
    pdf.cell(w=25, h=8, txt="PythonHow")
    pdf.image(name="pythonhow.png", w=10, h=10)



    # outputting each newly created pdf with their own names to the PDF folder
    pdf.output("PDFS/{}.pdf".format(filename.stem))


    