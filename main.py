# This Proram can convert excel files to pdf invoices

import pandas as pd
import glob
from fpdf import FPDF
import pathlib

# Read all excel files in the folder

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath , sheet_name='Sheet 1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    filename = pathlib.Path(filepath).stem
    invoice_nr = filename.split("-")[0]


    pdf.set_font("Arial", size = 18, style='B')
    pdf.cell(w=50, h=8, txt = f"Invoice {invoice_nr}", ln = True, align = 'L')

    pdf.output(f"PDFs/{filename}.pdf")

