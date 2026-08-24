import pandas as pd
import numpy as np

files = {
    "allocation": "D:\\downloads\\Bhubhneshwar allocation data.xlsx",
    "stock": "D:\\downloads\\Current Stock For Bhubhneshwar.xlsx",
    "inventory": "D:\\downloads\\Inventory Available for Sales - OMS-2026-08-17T12_52_57.842+05_30.csv"
}

stock_df = pd.read_excel(files['stock'])
alloc_df = pd.read_excel(files['allocation'])
inv_df = pd.read_csv(files['inventory'])

stock_df.rename(columns={'STYLE': 'Style', 'COLOR': 'Color', 'SIZE': 'Size', 'STOCK QUANTITY': 'StockQty'}, inplace=True)
alloc_df.rename(columns={'Max Allocated Qty': 'AllocQty'}, inplace=True)
inv_df.rename(columns={'Total Available Quantity': 'InvQty'}, inplace=True)

stock_agg = stock_df.groupby(['Style', 'Color', 'Size'])['StockQty'].sum().reset_index()
alloc_agg = alloc_df.groupby(['Style', 'Color', 'Size'])['AllocQty'].sum().reset_index()

current_total = pd.merge(stock_agg, alloc_agg, on=['Style', 'Color', 'Size'], how='outer').fillna(0)
current_total['TotalStoreStock'] = current_total['StockQty'] + current_total['AllocQty']

print("Sample total store stock:")
print(current_total.head())

inv_agg = inv_df.groupby(['Style', 'Color', 'Size'])['InvQty'].sum().reset_index()

# Find existing Style-Color
existing_sc = current_total[['Style', 'Color']].drop_duplicates()
existing_sc['IsExisting'] = True

print("Existing Style-Colors:", len(existing_sc))

inv_styles = inv_df[['Style', 'Color']].drop_duplicates()
merged_sc = pd.merge(inv_styles, existing_sc, on=['Style', 'Color'], how='left')

new_sc = merged_sc[merged_sc['IsExisting'].isna()]
print("New Style-Colors in Inv:", len(new_sc))

