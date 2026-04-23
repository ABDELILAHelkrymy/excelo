import pdfplumber
import pandas as pd

path = r"c:\Users\aelkrymy\Desktop\excel scripts\example\Liste_Demandes_Maintien_22-04-2026 12-08-55.pdf"

all_data = []
with pdfplumber.open(path) as pdf:
    for page in pdf.pages[:2]: # just first 2 pages for test
        tables = page.extract_tables()
        for table in tables:
            for row in table:
                all_data.append(row)

df = pd.DataFrame(all_data[1:], columns=all_data[0])
df.to_excel("test_pdf.xlsx", index=False)
print("Saved to test_pdf.xlsx")
