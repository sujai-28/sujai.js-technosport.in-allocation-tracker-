import pandas as pd
import numpy as np

def run_allocation():
    files = {
        "allocation": r"D:\downloads\Bhubhneshwar allocation data (1).xlsx",
        "stock": r"D:\downloads\Bhubhneshwar Store Stock.xlsx",
        "inventory": r"D:\downloads\Inventory Available for Sales - OMS-2026-08-18T11_52_32.411+05_30.csv"
    }

    print("Loading data...")
    stock_df = pd.read_excel(files['stock'])
    alloc_df = pd.read_excel(files['allocation'])
    inv_df = pd.read_csv(files['inventory'])

    # Standardize column names
    stock_df.rename(columns={'STYLE': 'Style', 'COLOR': 'Color', 'SIZE': 'Size', 'STOCK QUANTITY': 'StockQty'}, inplace=True)
    alloc_df.rename(columns={'Max Allocated Qty': 'AllocQty'}, inplace=True)
    inv_df.rename(columns={'Total Available Quantity': 'InvQty'}, inplace=True)
    
    # Ensure Size is string for consistency
    stock_df['Size'] = stock_df['Size'].astype(str).str.strip()
    alloc_df['Size'] = alloc_df['Size'].astype(str).str.strip()
    inv_df['Size'] = inv_df['Size'].astype(str).str.strip()
    
    # Create SKU mapping
    sku_mapping = inv_df[['Style', 'Color', 'Size', 'Client SKU Id / EAN']].drop_duplicates()
    sku_dict = {}
    for _, row in sku_mapping.iterrows():
        sku_dict[(row['Style'], row['Color'], row['Size'])] = row['Client SKU Id / EAN']

    stock_agg = stock_df.groupby(['Style', 'Color', 'Size'], as_index=False)['StockQty'].sum()
    alloc_agg = alloc_df.groupby(['Style', 'Color', 'Size'], as_index=False)['AllocQty'].sum()
    inv_agg = inv_df.groupby(['Style', 'Color', 'Size'], as_index=False)['InvQty'].sum()

    # Current Store Total
    current_total = pd.merge(stock_agg, alloc_agg, on=['Style', 'Color', 'Size'], how='outer').fillna(0)
    current_total['TotalStoreStock'] = current_total['StockQty'] + current_total['AllocQty']

    # Existing Styles and Style-Colors
    existing_sc = current_total[['Style', 'Color']].drop_duplicates()
    existing_styles = set(current_total['Style'].unique())

    print(f"Found {len(existing_styles)} existing styles in store.")

    # --- REPLENISHMENT REPORT ---
    replenish_rows = []
    
    # Iterate over existing Style-Colors
    for _, row in existing_sc.iterrows():
        style = row['Style']
        color = row['Color']
        
        # Get sizes from store and inventory
        store_sizes_df = current_total[(current_total['Style'] == style) & (current_total['Color'] == color)]
        inv_sizes_df = inv_agg[(inv_agg['Style'] == style) & (inv_agg['Color'] == color)]
        
        all_sizes = set(store_sizes_df['Size']).union(set(inv_sizes_df['Size']))
        
        for size in all_sizes:
            # Skip excluded sizes if they are new (from inventory but not in store)
            is_in_store = size in store_sizes_df['Size'].values
            if not is_in_store and size in ['3XL', '4XL', '5XL']:
                continue
                
            base = 5 if size in ['XLR', 'LAR'] else 3
            
            store_match = store_sizes_df[store_sizes_df['Size'] == size]
            inv_match = inv_sizes_df[inv_sizes_df['Size'] == size]
            
            total_current = store_match['TotalStoreStock'].values[0] if not store_match.empty else 0
            stock_qty = store_match['StockQty'].values[0] if not store_match.empty else 0
            alloc_qty = store_match['AllocQty'].values[0] if not store_match.empty else 0
            inv_qty = inv_match['InvQty'].values[0] if not inv_match.empty else 0
            
            replenish_qty = max(0, base - total_current)
            final_replenish = min(replenish_qty, inv_qty)
            
            if final_replenish > 0:
                replenish_rows.append({
                    'SKU': sku_dict.get((style, color, size), ""),
                    'Style': style,
                    'Color': color,
                    'Size': size,
                    'Store_Stock': stock_qty,
                    'Allocated_Stock': alloc_qty,
                    'Total_Current': total_current,
                    'Inventory_Available': inv_qty,
                    'Target_Base': base,
                    'Needed_Qty': replenish_qty,
                    'Allocated_Qty': final_replenish
                })
                
    replenish_df = pd.DataFrame(replenish_rows)
    print(f"Replenishment rows generated: {len(replenish_df)}")

    # --- NEW STYLE REPORT ---
    new_style_rows = []
    
    # Styles in inventory that are NOT in existing store styles
    inv_styles = inv_agg['Style'].unique()
    new_styles = [s for s in inv_styles if s not in existing_styles]
    
    print(f"Found {len(new_styles)} new styles in inventory.")
    
    for style in new_styles:
        style_inv = inv_agg[inv_agg['Style'] == style]
        
        for color, color_df in style_inv.groupby('Color'):
            avail_sizes = set(color_df[color_df['InvQty'] > 0]['Size'].str.upper())
            valid_other_sizes = avail_sizes - {'XLR', 'LAR', '3XL', '4XL', '5XL'}
            
            if 'XLR' in avail_sizes and 'LAR' in avail_sizes and len(valid_other_sizes) >= 1:
                # The color is eligible for new style allocation
                for _, row in color_df.iterrows():
                    size = row['Size']
                    size_upper = str(size).upper()
                    inv_qty = row['InvQty']
                    
                    if size_upper in ['3XL', '4XL', '5XL']:
                        continue
                        
                    base = 5 if size_upper in ['XLR', 'LAR'] else 3
                    
                    final_alloc = min(base, inv_qty)
                    if final_alloc > 0:
                        new_style_rows.append({
                            'SKU': sku_dict.get((style, color, size), ""),
                            'Style': style,
                            'Color': color,
                            'Size': size,
                            'Inventory_Available': inv_qty,
                            'Target_Base': base,
                            'Allocation_Qty': final_alloc
                        })
                
    new_style_df = pd.DataFrame(new_style_rows)
    print(f"New style rows generated: {len(new_style_df)}")

    # Save to Excel
    output_file = r"D:\downloads\Replenishment_and_New_Style_Report_Bhubhneshwar.xlsx"
    with pd.ExcelWriter(output_file) as writer:
        if not replenish_df.empty:
            replenish_df.to_excel(writer, sheet_name='Replenishment', index=False)
        else:
            pd.DataFrame([{'Message': 'No replenishment needed'}]).to_excel(writer, sheet_name='Replenishment', index=False)
            
        if not new_style_df.empty:
            new_style_df.to_excel(writer, sheet_name='New_Style', index=False)
        else:
            pd.DataFrame([{'Message': 'No new styles found'}]).to_excel(writer, sheet_name='New_Style', index=False)

    print(f"Report successfully saved to {output_file}")

if __name__ == "__main__":
    run_allocation()
