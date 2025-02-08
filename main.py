# This Proram can convert excel files to pdf invoices

import pandas as pd
import glob
from fpdf import FPDF
import pathlib

# Read all excel files in the folder

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:

    # Read the excel file using pandas and create a pdf object using FPDF and add a page to it
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
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

    pdf.output(f"PDFs/{filename}.pdf")
