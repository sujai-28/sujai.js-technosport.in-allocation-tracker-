import pandas as pd

xl = pd.ExcelFile(r'D:\downloads\S102&S103 BALANCE QTY DETAILS 07.07.2026.xlsx')

for sheet in xl.sheet_names:
    df = xl.parse(sheet)
    print(f"\n===== Sheet: {sheet} =====")
    # The Row Labels + Sum of QTY columns look like a SKU summary pivot
    summary = df[['Row Labels','Sum of QTY']].dropna(subset=['Row Labels','Sum of QTY'])
    summary = summary[summary['Row Labels'].astype(str).str.startswith('MFW')]
    summary.columns = ['SKU', 'Total_Qty']
    summary['SKU'] = summary['SKU'].astype(str).str.strip().str.upper()
    summary['Total_Qty'] = pd.to_numeric(summary['Total_Qty'], errors='coerce').fillna(0).astype(int)
    print(summary.to_string())
    total = summary['Total_Qty'].sum()
    print(f"  Total SKU rows: {len(summary)}, Grand Total Qty: {total}")
