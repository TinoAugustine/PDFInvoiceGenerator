import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Read data from the excel file
filepaths = glob.glob("invoices/*.xlsx")

# read each row in the sheet 1
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Set the PDF file payout and add pages
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # set the output file name
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Print data in the  pdf file
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}")

    pdf.output(f"PDFs/{filename}.pdf")
    print(df)
