import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Read data from the excel file
filepaths = glob.glob("invoices/*.xlsx")

# read each row in the sheet 1
for filepath in filepaths:

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
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    #Add Table from Excel to PDF
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    #Add a header for the pdf file.

    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    #Add Total under the Total Price Column

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=" ", border=1)
    pdf.cell(w=70, h=8, txt=" ", border=1)
    pdf.cell(w=30, h=8, txt=" ", border=1)
    pdf.cell(w=30, h=8, txt=" ", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    #Add Total Sum
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(255, 0, 0)
    pdf.cell(w=30, h=8, txt=f"The Total price is {total_sum}", ln=1)

    # Add Company Name and Logo
    pdf.set_font(family="Times", size=16, style="B" )
    pdf.set_text_color(0, 0, 255)
    pdf.cell(w=30, h=8, txt=f"The Awsome Company", ln=1)
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
    print(df)
