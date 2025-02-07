# This Proram can convert excel files to pdf invoices

import pandas as pd
import glob

# Read all excel files in the folder

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath , sheet_name='Sheet 1')
    print(df)