import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

#usando glob nos devuelve una lista con las tres direcciones de los documentos
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    #creamos un objeto pdf
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # extraemos informacion del la hoja de excel con el metodo stem nos regresa el nombre
    # pero sin extension
    filename = Path(filepath).stem 

    #con el metodo split nos devuelve una lista con dos objetos el primero el nombre y el segundo la fecha
    invoice_number, date = filename.split("-")

    # heater
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Factura\Invoice num. {invoice_number}", ln=1)

    # subheater
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50,h=8, txt=f"Date: {date}", ln=1 )

    dt = pd.read_excel(filepath, sheet_name="Sheet 1") # Leemos  el archivo excelel y lo guardamos en una variable

    # encabezados de columnas
    columns_table = dt.columns
    columns_table = [item.replace("_", " ").title() for item in columns_table]
    pdf.set_font(family="Times", size=10, style="B") 
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns_table[0], border=1) 
    pdf.cell(w=70, h=8, txt=columns_table[1],border=1) 
    pdf.cell(w=30, h=8, txt=columns_table[2], border=1) 
    pdf.cell(w=30, h=8, txt=columns_table[3], border=1) 
    pdf.cell(w=30, h=8, txt=columns_table[4], border=1, ln=1) 

    # filas de la tabla
    for index, row in dt.iterrows():
      pdf.set_font(family="Times", size=10) 
      pdf.set_text_color(80, 80, 80)
      pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1) 
      pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1) 
      pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1) 
      pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1) 
      pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1) 

    #fila del total
    total = dt["total_price"].sum()

    pdf.set_font(family="Times", size=10) 
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1) 
    pdf.cell(w=70, h=8, txt="", border=1) 
    pdf.cell(w=30, h=8, txt="", border=1) 
    pdf.cell(w=30, h=8, txt="", border=1) 
    pdf.cell(w=30, h=8, txt=str(total), border=1, ln=1) 

    # se agrega una oracion conformando el total
    pdf.set_font(family="Times", size=10, style="B") 
    pdf.cell(w=30, h=8, txt=f"The total price is {total}", ln=1)

    # se agrega el nombre de la company y la imagen 
    pdf.set_font(family="Times", size=14, style="B") 
    pdf.cell(w=25, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    #imprimimos y guardamos el resultado en la carpeta pdfs 
    pdf.output(f"PDFs/{filename}.pdf")


