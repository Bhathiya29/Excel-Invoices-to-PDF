# APP TO GENERATE PDF INVOICES FROM EXCEL FILES
import pandas as pd
import glob  # glob module returns all file paths that match a specific pattern
from fpdf import FPDF
from pathlib import Path

filePaths = glob.glob('Invoices/*.xlsx')  # Getting all the exel files into a list

try:
    for filePath in filePaths:
        df = pd.read_excel(filePath, sheet_name='Sheet 1')

        pdf = FPDF(orientation='P', unit='mm', format='A4')

        pdf.add_page()

        fileName = Path(filePath).stem  # getting the no of the file stem gives the name without the extension.xlsx
        invoiceNo = fileName.split('-')[0]  # only getting the number eg:-10001
        date = fileName.split('-')[1]  # getting the date

        pdf.set_font(family='Times', size=16, style='B')
        pdf.cell(w=50, h=8, txt=f"Invoice No {invoiceNo}",ln=1)  # writing the invoice no into the pdf

        pdf.set_font(family='Times', size=16, style='B')
        pdf.cell(w=50, h=8, txt=f"Date {date}")  # writing the Date into the pdf

        pdf.output(f"PDFs/{fileName}.pdf")


except:
    print('AN ERROR OCCURRED, PLEASE TRY AGAIN')
