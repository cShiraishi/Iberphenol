import pandas as pd
import os

excel_path = "PhenolDB_estruturado.xlsx"
if os.path.exists(excel_path):
    xl = pd.ExcelFile(excel_path)
    print "Sheets in Excel:"
    for sheet in xl.sheet_names:
        print "- " + sheet
else:
    print "Excel file not found at " + excel_path
