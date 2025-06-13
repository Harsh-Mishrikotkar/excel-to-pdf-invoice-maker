import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*.xlsx")

for fp in filepaths:
    df = pd.read_excel(fp, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    fName = Path(fp).stem
    invoiceNumber, date = fName.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50,h=8, txt=f"Invoice nr.{invoiceNumber}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50,h=8, txt=f"Date: {date}")


    pdf.output(f"PDFS/{fName}.pdf")