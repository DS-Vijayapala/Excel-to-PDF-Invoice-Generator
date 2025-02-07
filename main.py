# This Proram can convert excel files to pdf invoices

import pandas as pd
import glob
from fpdf import FPDF
import pathlib

# Read all excel files in the folder

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:

    # Read the excel file using pandas and create a pdf object using FPDF and add a page to it
    df = pd.read_excel(filepath , sheet_name='Sheet 1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # Get the filename without the extension and split it by "-" to get the invoice number 
    filename = pathlib.Path(filepath).stem 
    invoice_nr = filename.split("-")[0]

    # Add the invoice number to the pdf
    pdf.set_font(family="Times", size = 18, style='B')
    pdf.cell(w=50, h=8, txt = f"Invoice {invoice_nr}", ln = True, align = 'L')

    pdf.output(f"PDFs/{filename}.pdf")

