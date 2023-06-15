# APP TO GENERATE PDF INVOICES FROM EXCEL FILES
import pandas as pd
import glob  # glob module returns all file paths that match a specific pattern
from fpdf import FPDF
from pathlib import Path

filePaths = glob.glob('Invoices/*.xlsx')  # Getting all the exel files into a list

for filePath in filePaths:
    # df = pd.read_excel(filePath, sheet_name='Sheet 1')

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    fileName = Path(filePath).stem  # getting the no of the file stem gives the name without the extension.xlsx
    invoiceNo = fileName.split('-')[0]  # only getting the number eg:-10001
    date = fileName.split('-')[1]  # getting the date

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice No {invoiceNo}", ln=1)  # writing the invoice no into the pdf

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Date {date}",ln=1)  # writing the Date into the pdf

    # Reading the excel files and iterating the rows
    df = pd.read_excel(filePath, sheet_name='Sheet 1')
    columns = list(df.columns)  # getting the column headers of the excel file using df.columns method and putting it to a list
    columns = [item.replace('_', ' ').title() for item in columns]  # replacing the _ and capitalizing the first letter
    # printing the column headers
    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Adding rows to the pdf
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1,ln=1)

    pdf.output(f"PDFs/{fileName}.pdf")
