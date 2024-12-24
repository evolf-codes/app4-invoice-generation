import pandas as pd
import os
from fpdf import FPDF

# Path to the Excel file
directory_path = "./invoices"
invoice_names = os.listdir(directory_path)

print(invoice_names)

# grab date from file name
for i in invoice_names:
    # set up pdf
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False, margin=0)

    invoice_number = i.split("-")[0]
    print(invoice_number)

    invoice_date = i.split("-")[1].strip(".xlsx")
    print(invoice_date)

    # Load the Excel file into a DataFrame
    df_path = os.path.join(directory_path, i)
    df = pd.read_excel(df_path, sheet_name="Sheet 1")
    print(df)

    pdf.add_page()

    pdf.set_font(family="Times", style="B", size=18)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=12, txt="Invoce nr. " +
             invoice_number, align="L", ln=1, border=0)
    pdf.cell(w=0, h=12, txt="Date " + invoice_date, align="L", ln=1, border=0)

    pdf.set_font(family="Times", size=12)
    columns = list(df.columns)

    pdf.cell(w=30, h=8, txt="ID", border=1)
    pdf.cell(w=70, h=8, txt="Name", border=1)
    pdf.cell(w=30, h=8, txt="Amount", border=1)
    pdf.cell(w=30, h=8, txt="Price", border=1)
    pdf.cell(w=30, h=8, txt="Total", border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # total
    total_due = str(df["total_price"].sum())
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=70, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=30, h=8, txt=total_due, border=1, ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=60, txt="The total due is: $" +
             total_due,  align="L", ln=1, border=0)

    pdf.output("PDF/"+invoice_number+invoice_date+".pdf")
