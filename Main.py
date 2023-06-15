# APP TO GENERATE PDF INVOICES FROM EXCEL FILES
import pandas as pd
import glob

filePaths = glob.glob('Invoices/*.xlsx')  # Getting all the exel files into a list

for filePath in filePaths:
    df = pd.read_excel(filePath,sheet_name='Sheet 1')
    print(df)
