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
    pdf.cell(w=50,h=8, txt=f"Date: {date}", ln=1)

    columns = df.columns
    columns = [i.replace("_", " ").title() for i in columns]
    pdf.set_text_color(80,80,80)
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for i,j in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(j["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(j["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(j["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(j["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(j["total_price"]),border=1, ln=1)

    totalSum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(totalSum),border=1, ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The Total price is {totalSum}", ln=1)

    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFS/{fName}.pdf")