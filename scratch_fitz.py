import fitz
import pandas as pd
import re

ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

def clean_text(val):
    if isinstance(val, str):
        return ILLEGAL_CHARACTERS_RE.sub("", val)
    return val

path = r"c:\Users\aelkrymy\Desktop\excel scripts\example\Liste_Demandes_Maintien_22-04-2026 12-08-55.pdf"

doc = fitz.open(path)
all_data = []

for page in doc[:2]:
    tabs = page.find_tables()
    for tab in tabs:
        for row in tab.extract():
            cleaned_row = [clean_text(cell) for cell in row]
            all_data.append(cleaned_row)

if all_data:
    df = pd.DataFrame(all_data[1:], columns=all_data[0])
    df.to_csv("test_fitz.csv", index=False, encoding="utf-8-sig")
    print("Saved test_fitz.csv")
