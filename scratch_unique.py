import pandas as pd

df = pd.read_csv('test_fitz.csv')
unique_chars = set()
for col in df.columns:
    unique_chars.update(col)
for row in df.itertuples(index=False):
    for cell in row:
        if isinstance(cell, str):
            unique_chars.update(cell)

print(f"Total unique weird chars: {len(unique_chars)}")
print("Chars:", [hex(ord(c)) for c in unique_chars])
