# This Proram can convert excel files to pdf invoices

import pandas as pd
import glob
from fpdf import FPDF
import pathlib

# Read all excel files in the folder

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:

    # Read the excel file using pandas and create a pdf object using FPDF and add a page to it

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # Get the filename without the extension and split it by "-" to get the invoice number
    filename = pathlib.Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Add the invoice number to the pdf
    pdf.set_font(family="Times", size=20, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice No: {invoice_nr}", ln=2, align='L')

    pdf.set_font(family="Times", size=14, style='B')
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1, align='L')

    pdf.cell(w=50, h=8, txt=" ", ln=1, align='L')  # Add a space line

    df = pd.read_excel(filepath, sheet_name='Sheet 1')  # Read the excel file

    # Add the column names to the pdf as a header

    columns = list(df.columns)  # convert the columns to a list
    # Replace the underscore with a space and capitalize the first letter of each word
    columns = [column.replace("_", " ").title() for column in columns]

    pdf.set_font(family="Times", size=11, style='B')
    pdf.set_text_color(80, 80, 80)

    pdf.cell(w=30, h=10, txt=columns[0], border=1, ln=0, align='C')
    pdf.cell(w=70, h=10, txt=columns[1], border=1, ln=0, align='C')
    pdf.cell(w=40, h=10, txt=columns[2], border=1, ln=0, align='C')
    pdf.cell(w=30, h=10, txt=columns[3], border=1, ln=0, align='C')
    pdf.cell(w=25, h=10, txt=columns[4], border=1, ln=1, align='C')

    # Add the data to the pdf file row by row

    for index, row in df.iterrows():

        pdf.set_font(family="Times", size=12, style='B')
        pdf.set_text_color(80, 80, 80)

        pdf.cell(w=30, h=8, txt=str(
            row["product_id"]), border=1, ln=0, align='L')
        pdf.cell(w=70, h=8, txt=str(
            row["product_name"]), border=1, ln=0, align='L')
        pdf.cell(w=40, h=8, txt=str(
            row["amount_purchased"]), border=1, ln=0, align='C')
        pdf.cell(w=30, h=8, txt=str(
            row["price_per_unit"]), border=1, ln=0, align='C')
        pdf.cell(w=25, h=8, txt=str(
            row["total_price"]), border=1, ln=1, align='C')

    pdf.output(f"PDFs/{filename}.pdf")
