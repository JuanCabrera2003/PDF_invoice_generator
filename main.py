import pandas as pd
import glob

#usando glob nos devuelve una lista con las tres direcciones de los documentos
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    data = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(data)

