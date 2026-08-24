import pandas as pd
import json

files = {
    "allocation": "D:\\downloads\\Bhubhneshwar allocation data.xlsx",
    "stock": "D:\\downloads\\Current Stock For Bhubhneshwar.xlsx",
    "inventory": "D:\\downloads\\Inventory Available for Sales - OMS-2026-08-17T12_52_57.842+05_30.csv"
}

for name, path in files.items():
    print(f"--- {name} ---")
    try:
        if path.endswith('.csv'):
            df = pd.read_csv(path, nrows=5)
        else:
            df = pd.read_excel(path, nrows=5)
            print(df.head())
        print("\nColumns:", df.columns.tolist())
    except Exception as e:
        print("Error:", e)
    print("\n")
