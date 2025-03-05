import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

#usando glob nos devuelve una lista con las tres direcciones de los documentos
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    data = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(data)
    #creamos un objeto pdf
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # extraemos informacion del la hoja de excel con el metodo stem nos regresa el nombre
    # pero sin extension
    filename = Path(filepath).stem 

    #con el metodo split nos devuelve una lista con dos objetos el primero el nombre y el segundo la fecha
    invoice_number = filename.split("-")[0]
    date = filename.split("-")[1]

    # heater
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Factura\Invoice num. {invoice_number}", ln=True)

    # subheater
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50,h=8, txt=f"Date: {date}", ln=True )




    #imprimimos y guardamos el resultado en la carpeta pdfs 
    pdf.output(f"PDFs/{filename}.pdf")


