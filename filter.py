import pandas as pd

# Read both Excel files
df_global = pd.read_excel("data1/liste-global.xlsx")
df_payees = pd.read_excel("data1/liste-payées.xlsx")

# Get the set of paid IDs from liste-payées
paid_ids = set(df_payees["N_Identifiant"].astype(str).str.strip())

# Filter out rows where 'رقم بطاقة التعريف الوطنية' is in the paid list
df_global["_match_key"] = df_global["رقم بطاقة التعريف الوطنية"].astype(str).str.strip()
df_filtered = df_global[~df_global["_match_key"].isin(paid_ids)].drop(columns=["_match_key"])

# Save the result to a new Excel file
df_filtered.to_excel("data1/liste-non-payées.xlsx", index=False)

print(f"Global list rows:   {len(df_global)}")
print(f"Paid list rows:     {len(df_payees)}")
print(f"Filtered list rows: {len(df_filtered)}")
print("Result saved to: data1/liste-non-payées.xlsx")