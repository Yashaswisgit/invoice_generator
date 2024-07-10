import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_no, date = filename.split("_")


    pdf.set_font(family = "Times", style = "B", size=16, ln = 1)
    pdf.cell(w ="8", h ="6", txt = f"Invoice_No: {invoice_no}")


    pdf.set_font(family = "Times", style = "B", size=16, ln = 1)
    pdf.cell(w ="8", h ="6", txt = f"Date: {Date}", ln = 1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    columns = df.columns
    columns = [item.replace("-", " ").title() for item in columns]
    pdf.set_font(family="Times", size=12, style = "B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family = "Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w = 30, h = 8, txt = str(row["product_id"]), border = 1)
        pdf.cell(w = 70, h = 8, txt = str(row["product_name"]), border = 1)
        pdf.cell(w = 30, h = 8, txt = str(row["amount_purchased"]), border = 1)
        pdf.cell(w = 30, h = 8, txt = str(row["price_per_unit"]), border = 1)
        pdf.cell(w = 30, h = 8, txt = str(row["total_price"]), border = 1, ln = 1)
    total = df["total_price"].sum
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total), border=1, ln=1)
    pdf.set_font(family="Times", size=12, style = "B")
    pdf.cell(w=30, h=8, txt=f"The total price:{total}", ln = 1)


    pdf.output(f"Pdfs/{filename}.pdf")