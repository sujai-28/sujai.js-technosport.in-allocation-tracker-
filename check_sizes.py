import pandas as pd

files = {
    "allocation": "D:\\downloads\\Bhubhneshwar allocation data.xlsx",
    "stock": "D:\\downloads\\Current Stock For Bhubhneshwar.xlsx",
    "inventory": "D:\\downloads\\Inventory Available for Sales - OMS-2026-08-17T12_52_57.842+05_30.csv"
}

stock_df = pd.read_excel(files['stock'])
alloc_df = pd.read_excel(files['allocation'])
inv_df = pd.read_csv(files['inventory'])

print("Stock unique sizes:", stock_df['SIZE'].unique()[:20])
print("Inv unique sizes:", inv_df['Size'].dropna().unique()[:20])
print("Alloc unique sizes:", alloc_df['Size'].dropna().unique()[:20])
