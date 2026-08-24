import os
import glob
import pandas as pd
import numpy as np

def get_latest_file(directory, ext='*.xlsx'):
    files = glob.glob(os.path.join(directory, ext))
    files = [f for f in files if not os.path.basename(f).startswith('~')]
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def clean_store_code(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        return s[:-2]
    return s

def main():
    print("Finding data files...")
    base_dir = os.path.dirname(os.path.abspath(__file__))
    stock_dir = os.path.join(base_dir, "ebo stock track data", "current stock")
    transit_dir = os.path.join(base_dir, "ebo stock track data", "intransit")
    alloc_dir = os.path.join(base_dir, "ebo stock track data", "ALLOCATION")
    ag_dir = os.path.join(base_dir, "ebo stock track data", "STYLE WISE AG")
    
    stock_file = get_latest_file(stock_dir, '*.xlsx')
    transit_file = get_latest_file(transit_dir, '*.xlsx')
    
    # Search for both csv and xlsx allocation files
    alloc_csv = get_latest_file(alloc_dir, '*.csv')
    alloc_xlsx = get_latest_file(alloc_dir, '*.xlsx')
    if alloc_csv and alloc_xlsx:
        alloc_file = alloc_csv if os.path.getmtime(alloc_csv) > os.path.getmtime(alloc_xlsx) else alloc_xlsx
    else:
        alloc_file = alloc_csv or alloc_xlsx
        
    ag_file = get_latest_file(ag_dir, '*.xlsx')
    
    if not stock_file: return print("Error: Could not find Stock file.")
    if not transit_file: return print("Error: Could not find Transit file.")
    if not alloc_file: return print("Error: Could not find Allocation file.")
    if not ag_file: return print("Error: Could not find AG file.")
        
    print(f"Reading Stock: {stock_file}")
    # Peek at columns to handle case-insensitivity and header variation
    peek_stock = pd.read_excel(stock_file, nrows=0)
    available_stock_cols = [str(c).strip() for c in peek_stock.columns]
    
    store_name_col = next((c for c in peek_stock.columns if str(c).strip().lower() in ['store name', 'store_name', 'owner site', 'owner_site']), 'Store Name')
    # Prioritise site_code / site code over store code for the numeric store code
    store_code_col = next((c for c in peek_stock.columns if str(c).strip().lower() in ['site_code', 'site code', 'sitecode', 'store code', 'store_code', 'storecode']), 'store_code')
    style_col = next((c for c in peek_stock.columns if str(c).strip().lower() == 'style'), 'Style')
    color_col = next((c for c in peek_stock.columns if str(c).strip().lower() in ['color', 'colour']), 'Colour')
    size_col = next((c for c in peek_stock.columns if str(c).strip().lower() == 'size'), 'Size')
    qty_col = next((c for c in peek_stock.columns if str(c).strip().lower() in ['quantity', 'qty', 'stock quantity', 'stock_quantity']), 'quantity')
    
    df_stock = pd.read_excel(stock_file, usecols=[store_name_col, store_code_col, style_col, color_col, size_col, qty_col])
    df_stock.rename(columns={
        store_name_col: 'Store Name',
        store_code_col: 'store_code',
        style_col: 'Style',
        color_col: 'Color',
        size_col: 'Size',
        qty_col: 'Qty'
    }, inplace=True)
    df_stock['Source'] = 'stock'
    
    print(f"Reading Transit: {transit_file}")
    peek_tran = pd.read_excel(transit_file, nrows=0)
    
    t_store_code_col = next((c for c in peek_tran.columns if str(c).strip().lower() in ['store code', 'store_code', 'storecode', 'site code', 'site_code', 'sitecode']), 'Store Code')
    t_store_name_col = next((c for c in peek_tran.columns if str(c).strip().lower() in ['ebo name', 'ebo_name', 'store name', 'store_name', 'owner site', 'owner_site']), 'EBO Name')
    t_style_col = next((c for c in peek_tran.columns if str(c).strip().lower() == 'style'), 'STYLE')
    t_color_col = next((c for c in peek_tran.columns if str(c).strip().lower() in ['color', 'colour']), 'COLOR')
    t_size_col = next((c for c in peek_tran.columns if str(c).strip().lower() == 'size'), 'SIZE')
    t_qty_col = next((c for c in peek_tran.columns if str(c).strip().lower() in ['transit qty', 'transit_qty', 'quantity', 'qty']), 'Transit Qty')
    
    df_tran = pd.read_excel(transit_file, usecols=[t_store_code_col, t_store_name_col, t_style_col, t_color_col, t_size_col, t_qty_col])
    df_tran.rename(columns={
        t_store_code_col: 'store_code',
        t_store_name_col: 'Store Name',
        t_style_col: 'Style',
        t_color_col: 'Color',
        t_size_col: 'Size',
        t_qty_col: 'Qty'
    }, inplace=True)
    df_tran['Source'] = 'transit'
    
    print(f"Reading Allocation: {alloc_file}")
    if alloc_file.endswith('.csv'):
        # Peek at columns to handle case-insensitivity
        peek_df = pd.read_csv(alloc_file, nrows=0)
    else:
        peek_df = pd.read_excel(alloc_file, nrows=0)
        
    available_cols = peek_df.columns.tolist()
    store_code_col = next((c for c in available_cols if str(c).strip().lower() in ['store code', 'store_code', 'storecode']), 'store code')
    style_col = next((c for c in available_cols if str(c).strip().lower() == 'style'), 'Style')
    color_col = next((c for c in available_cols if str(c).strip().lower() == 'color'), 'Color')
    size_col = next((c for c in available_cols if str(c).strip().lower() == 'size'), 'Size')
    qty_col = next((c for c in available_cols if str(c).strip().lower() in ['max allocated qty', 'allocated qty', 'quantity', 'qty']), 'Max Allocated Qty')
    
    cols_to_use = [store_code_col, style_col, color_col, size_col, qty_col]
    if alloc_file.endswith('.csv'):
        df_alloc = pd.read_csv(alloc_file, usecols=cols_to_use)
    else:
        df_alloc = pd.read_excel(alloc_file, usecols=cols_to_use)
        
    df_alloc.rename(columns={
        store_code_col: 'store_code',
        style_col: 'Style',
        color_col: 'Color',
        size_col: 'Size',
        qty_col: 'Qty'
    }, inplace=True)
    df_alloc['Store Name'] = '' # Alloc doesn't have store name, we'll map it later
    df_alloc['Source'] = 'alloc'

    print(f"Reading AG Mapping: {ag_file}")
    df_ag = pd.read_excel(ag_file, usecols=['Style', 'AG Name'])
    
    # Combine all data
    print("Combining data...")
    df_all = pd.concat([df_stock, df_tran, df_alloc], ignore_index=True)
    
    # Clean up strings
    df_all['Style'] = df_all['Style'].astype(str).str.strip().str.upper()
    df_all['Color'] = df_all['Color'].astype(str).str.strip().str.upper()
    df_all['Size'] = df_all['Size'].astype(str).str.strip().str.upper()
    df_all['store_code'] = df_all['store_code'].apply(clean_store_code)
    df_all['Qty'] = pd.to_numeric(df_all['Qty'], errors='coerce').fillna(0)
    
    # Composite key for Option counting
    df_all['StyleColor'] = df_all['Style'] + "_" + df_all['Color']
    
    df_ag['Style'] = df_ag['Style'].astype(str).str.strip().str.upper()

    # Merge AG Name into df_all early
    df_all = pd.merge(df_all, df_ag, on='Style', how='left').fillna({'AG Name': 'UNKNOWN AG'})

    # Create Store Name mapping
    # Filter out empty or null store names to build a reliable map
    valid_names = df_all[(df_all['Store Name'].notnull()) & (df_all['Store Name'].astype(str).str.strip() != '')]
    store_map = valid_names.drop_duplicates('store_code').set_index('store_code')['Store Name'].to_dict()
    
    # Apply the mapping uniformly. If a code isn't in the map, fallback to the code itself.
    df_all['Store Name'] = df_all['store_code'].map(store_map).fillna(df_all['store_code'])

    # Helper function to evaluate validity PER STORE
    def get_valid_store_options(df_subset):
        # Group by Store, StyleColor, AG Name, Size
        agg = df_subset.groupby(['store_code', 'Store Name', 'StyleColor', 'AG Name', 'Size'])['Qty'].sum().reset_index()
        agg = agg[agg['Qty'] > 0]
        
        totals = agg.groupby(['store_code', 'Store Name', 'StyleColor', 'AG Name'])['Qty'].sum()
        sizes_set = agg.groupby(['store_code', 'Store Name', 'StyleColor', 'AG Name'])['Size'].apply(set)
        
        df_eval = pd.DataFrame({'TotalQty': totals, 'Sizes': sizes_set}).reset_index()
        df_eval['Style'] = df_eval['StyleColor'].apply(lambda x: x.split('_')[0])
        
        def is_valid(row):
            if row['TotalQty'] < 7: return False
            # Bypass size check for Socks, Men's Shoes, and Caps
            if str(row['AG Name']).strip().lower() in ['socks', "men's shoes", 'caps']:
                return True
            sz = row['Sizes']
            has_adult = ('LAR' in sz) and ('XLR' in sz) and (len(sz) >= 3)
            has_kids = (('12' in sz) or ('12Y' in sz)) and (('14Y' in sz) or ('14' in sz)) and (('10Y' in sz) or ('8Y' in sz) or ('10' in sz) or ('8' in sz))
            return has_adult or has_kids
            
        df_eval['Valid'] = df_eval.apply(is_valid, axis=1)
        return df_eval[df_eval['Valid']].copy()

    print("Evaluating ONLY STOCK...")
    valid_stock = get_valid_store_options(df_all[df_all['Source'] == 'stock'])
    
    print("Evaluating STOCK + TRANSIT...")
    valid_st = get_valid_store_options(df_all[df_all['Source'].isin(['stock', 'transit'])])
    
    print("Evaluating ALL (Stock + Transit + Alloc)...")
    valid_all = get_valid_store_options(df_all)

    # ------------------
    # AGGREGATE STORE AND AG WISE (RAW)
    # ------------------
    store_ag_stock = valid_stock.groupby(['store_code', 'Store Name', 'AG Name']).size().rename('Only Stock Valid Options')
    store_ag_st = valid_st.groupby(['store_code', 'Store Name', 'AG Name']).size().rename('Stock and Transit Valid Options')
    store_ag_all = valid_all.groupby(['store_code', 'Store Name', 'AG Name']).size().rename('All Valid Options')
    
    df_store_ag_out = pd.concat([store_ag_all, store_ag_st, store_ag_stock], axis=1).fillna(0).reset_index()
    
    # LOAD AG WORKING FOR CAPPING
    ag_working_dir = os.path.join(base_dir, "AG VALIDATION DATA", "AG WORKING")
    ag_working_file = get_latest_file(ag_working_dir, '*.xlsx')
    
    if ag_working_file:
        print(f"Reading AG Working for Capping: {ag_working_file}")
        df_ag_working = pd.read_excel(ag_working_file, sheet_name="AG Working")
        df_ag_working['Store Code'] = df_ag_working['Store Code'].apply(clean_store_code)
        ag_col = 'AG' if 'AG' in df_ag_working.columns else 'AG Name'
        df_ag_working[ag_col] = df_ag_working[ag_col].astype(str).str.strip()
        df_ag_working['Final Options'] = pd.to_numeric(df_ag_working['Final Options'], errors='coerce').fillna(0)
        
        # Group by Store and AG to get Max Final Options
        ag_caps = df_ag_working.groupby(['Store Code', ag_col], as_index=False)['Final Options'].sum()
        ag_caps.rename(columns={'Store Code': 'store_code', ag_col: 'AG Name', 'Final Options': 'Cap'}, inplace=True)
        
        # Merge caps into raw data with OUTER JOIN to keep all Final Options
        df_store_ag_out = pd.merge(df_store_ag_out, ag_caps, on=['store_code', 'AG Name'], how='outer')
        
        # Fill missing Store Name
        store_names = df_all.drop_duplicates('store_code').set_index('store_code')['Store Name'].to_dict()
        df_store_ag_out['Store Name'] = df_store_ag_out.apply(
            lambda r: store_names.get(r['store_code'], r['Store Name']) if pd.isna(r['Store Name']) else r['Store Name'],
            axis=1
        )
        
        # Treat missing valid options and missing caps (NaN) as 0
        df_store_ag_out['Cap'] = df_store_ag_out['Cap'].fillna(0)
        df_store_ag_out['All Valid Options'] = df_store_ag_out['All Valid Options'].fillna(0)
        df_store_ag_out['Stock and Transit Valid Options'] = df_store_ag_out['Stock and Transit Valid Options'].fillna(0)
        df_store_ag_out['Only Stock Valid Options'] = df_store_ag_out['Only Stock Valid Options'].fillna(0)        
        # Cap values where Cap is not null
        for col in ['All Valid Options', 'Stock and Transit Valid Options', 'Only Stock Valid Options']:
            df_store_ag_out[col] = df_store_ag_out.apply(
                lambda r: min(r[col], r['Cap']),
                axis=1
            )
            
        df_store_ag_out.rename(columns={'Cap': 'Final Options'}, inplace=True)
    else:
        print("Warning: Could not find AG Working file for capping.")
        
    if 'Final Options' not in df_store_ag_out.columns:
        df_store_ag_out['Final Options'] = 0
        
    # ------------------
    # LOAD PRIORITY LISTS FOR SORTING
    # ------------------
    priority_dir = os.path.join(base_dir, "ag validation priority")
    store_pri_file = os.path.join(priority_dir, "Store wise May sales qty.xlsx")
    style_pri_file = os.path.join(priority_dir, "Style wise May Bill qty (1).xlsx")
    
    store_pri_dict = {}
    if os.path.exists(store_pri_file):
        try:
            df_store_pri = pd.read_excel(store_pri_file)
            # Normalize Store Code to string
            df_store_pri['Store Code'] = df_store_pri['Store Code'].apply(clean_store_code)
            store_pri_dict = dict(zip(df_store_pri['Store Code'], df_store_pri['priority']))
            print(f"Loaded {len(store_pri_dict)} store priorities for sorting.")
        except Exception as e:
            print(f"Error loading store priority file: {e}")
            
    style_pri_dict = {}
    if os.path.exists(style_pri_file):
        try:
            df_style_pri = pd.read_excel(style_pri_file)
            col = next((c for c in df_style_pri.columns if str(c).strip().lower() in ['style', 'option', 'option name']), None)
            if col:
                df_style_pri[col] = df_style_pri[col].astype(str).str.strip().str.upper()
                temp_dict = {}
                for _, row in df_style_pri.iterrows():
                    val = str(row[col]).strip().upper()
                    pri = row['priority']
                    style = val.split('_')[0]
                    if style not in temp_dict or pri < temp_dict[style]:
                        temp_dict[style] = pri
                style_pri_dict = temp_dict
                print(f"Loaded {len(style_pri_dict)} style priorities for sorting.")
            else:
                print("Warning: Style/option column not found in style priority file.")
        except Exception as e:
            print(f"Error loading style priority file: {e}")

    def sort_df_by_store_priority(df, store_code_col='store_code'):
        if df.empty:
            return df
        df = df.copy()
        clean_codes = df[store_code_col].apply(clean_store_code)
        df['_store_pri'] = clean_codes.map(store_pri_dict).fillna(999999)
        # Sort by priority ascending, then by store code
        df_sorted = df.sort_values(by=['_store_pri', store_code_col]).drop(columns=['_store_pri'])
        return df_sorted

    # Re-calculate Store Wise and AG Wise from Capped Raw Data
    def get_ag_category(ag_name):
        name = str(ag_name).strip().lower()
        if name.startswith('b'):
            return 'K'
        elif name.startswith('w'):
            return 'W'
        elif name.startswith('m'):
            return 'M'
        return 'M'

    df_store_ag_out['Category'] = df_store_ag_out['AG Name'].apply(get_ag_category)

    df_store_out = df_store_ag_out.groupby(['store_code', 'Store Name'], as_index=False)[['All Valid Options', 'Stock and Transit Valid Options', 'Only Stock Valid Options', 'Final Options']].sum()
    df_store_out.rename(columns={'Final Options': 'Sum of Final Options'}, inplace=True)
    
    # Add M, W, K breakdowns for each of the metrics
    for cat in ['M', 'W', 'K']:
        df_cat = df_store_ag_out[df_store_ag_out['Category'] == cat]
        if not df_cat.empty:
            df_cat_sum = df_cat.groupby(['store_code', 'Store Name'], as_index=False)[['All Valid Options', 'Stock and Transit Valid Options', 'Only Stock Valid Options', 'Final Options']].sum()
            df_cat_sum.rename(columns={
                'All Valid Options': f'All Valid Options - {cat}',
                'Stock and Transit Valid Options': f'Stock and Transit Valid Options - {cat}',
                'Only Stock Valid Options': f'Only Stock Valid Options - {cat}',
                'Final Options': f'Sum of Final Options - {cat}'
            }, inplace=True)
            df_store_out = pd.merge(df_store_out, df_cat_sum, on=['store_code', 'Store Name'], how='left')
        else:
            df_store_out[f'All Valid Options - {cat}'] = 0
            df_store_out[f'Stock and Transit Valid Options - {cat}'] = 0
            df_store_out[f'Only Stock Valid Options - {cat}'] = 0
            df_store_out[f'Sum of Final Options - {cat}'] = 0
            
    # Fill any NaNs with 0
    df_store_out.fillna(0, inplace=True)
    
    # Ensure all option columns are integers
    option_cols = [
        'Sum of Final Options', 'Sum of Final Options - M', 'Sum of Final Options - W', 'Sum of Final Options - K',
        'All Valid Options', 'All Valid Options - M', 'All Valid Options - W', 'All Valid Options - K',
        'Stock and Transit Valid Options', 'Stock and Transit Valid Options - M', 'Stock and Transit Valid Options - W', 'Stock and Transit Valid Options - K',
        'Only Stock Valid Options', 'Only Stock Valid Options - M', 'Only Stock Valid Options - W', 'Only Stock Valid Options - K'
    ]
    for col in option_cols:
        if col in df_store_out.columns:
            df_store_out[col] = df_store_out[col].astype(int)

    # Reorder columns explicitly to match UI order
    store_cols = [
        'store_code', 'Store Name', 'Sum of Final Options',
        'Sum of Final Options - M', 'Sum of Final Options - W', 'Sum of Final Options - K',
        'Only Stock Valid Options', 'Only Stock Valid Options - M', 'Only Stock Valid Options - W', 'Only Stock Valid Options - K',
        'Stock and Transit Valid Options', 'Stock and Transit Valid Options - M', 'Stock and Transit Valid Options - W', 'Stock and Transit Valid Options - K',
        'All Valid Options', 'All Valid Options - M', 'All Valid Options - W', 'All Valid Options - K'
    ]
    store_cols = [c for c in store_cols if c in df_store_out.columns]
    df_store_out = df_store_out[store_cols]

    # --- FORMAT TO MATCH UI ---
    # 1. Calculate GRAND TOTAL row
    totals = {'store_code': 'TOTAL', 'Store Name': 'GRAND TOTAL'}
    for c in store_cols[2:]:
        totals[c] = df_store_out[c].sum()
    
    df_store_out_fmt = pd.concat([pd.DataFrame([totals]), df_store_out], ignore_index=True)

    # 2. Add Percentages
    for suffix in ['', ' - M', ' - W', ' - K']:
        denom_col = 'Sum of Final Options' + suffix
        for prefix in ['Only Stock Valid Options', 'Stock and Transit Valid Options', 'All Valid Options']:
            num_col = prefix + suffix
            if num_col in df_store_out_fmt.columns and denom_col in df_store_out_fmt.columns:
                def fmt_val(row):
                    num = row[num_col]
                    denom = row[denom_col]
                    if denom > 0:
                        pct = int(round((num / denom) * 100))
                        return f"{num} \n({pct}%)"
                    return str(num)
                df_store_out_fmt[num_col] = df_store_out_fmt.apply(fmt_val, axis=1)

    # 3. Create Stacked Headers using newlines
    flat_cols = []
    for c in df_store_out_fmt.columns:
        if c == 'store_code': flat_cols.append('store_code')
        elif c == 'Store Name': flat_cols.append('Store Name')
        elif c.startswith('Sum of Final Options'):
            sub = c.split(' - ')[1] if ' - ' in c else 'Overall'
            flat_cols.append(f"Final Options\n{sub}")
        elif c.startswith('Only Stock Valid Options'):
            sub = c.split(' - ')[1] if ' - ' in c else 'Overall'
            flat_cols.append(f"Only Stock\n{sub}")
        elif c.startswith('Stock and Transit Valid Options'):
            sub = c.split(' - ')[1] if ' - ' in c else 'Overall'
            flat_cols.append(f"Stock & Transit\n{sub}")
        elif c.startswith('All Valid Options'):
            sub = c.split(' - ')[1] if ' - ' in c else 'Overall'
            flat_cols.append(f"All Valid\n{sub}")
        else:
            flat_cols.append(c)
            
    df_store_out_fmt.columns = flat_cols

    df_ag_out = df_store_ag_out.groupby(['AG Name'], as_index=False)[['All Valid Options', 'Stock and Transit Valid Options', 'Only Stock Valid Options', 'Final Options']].sum()
    df_ag_out.rename(columns={'Final Options': 'Sum of Final Options'}, inplace=True)
    # Reorder columns
    ag_cols = ['AG Name', 'Sum of Final Options', 'Only Stock Valid Options', 'Stock and Transit Valid Options', 'All Valid Options']
    ag_cols = [c for c in ag_cols if c in df_ag_out.columns]
    df_ag_out = df_ag_out[ag_cols]
    
    # Ensure AG-wise columns are also cast to int
    for col in ['Sum of Final Options', 'Only Stock Valid Options', 'Stock and Transit Valid Options', 'All Valid Options']:
        if col in df_ag_out.columns:
            df_ag_out[col] = df_ag_out[col].astype(int)
    
    # Load CM mapping
    cm_dict = {}
    cm_file = os.path.join(base_dir, "AG VALIDATION DATA", "CM", "Name_Zones (2).xlsx")
    if os.path.exists(cm_file):
        try:
            df_cm = pd.read_excel(cm_file, sheet_name="Sheet1")
            df_cm['Store Code'] = df_cm['Store Code'].apply(clean_store_code)
            df_cm['CM'] = df_cm['CM'].astype(str).str.strip()
            for _, r in df_cm.iterrows():
                sc = r['Store Code']
                cm_val = r['CM']
                if pd.notna(cm_val) and str(cm_val).strip().lower() != 'nan':
                    cm_dict[sc] = str(cm_val).strip()
        except Exception as e:
            print(f"Error loading CM mapping file in validate_styles: {e}")

    # ------------------
    # AGGREGATE STYLE WISE (Pivot)
    # ------------------
    # Start with unique store_code and Store Name
    df_style_out = df_store_out[['store_code', 'Store Name']].copy()
    df_style_out['CM'] = df_style_out['store_code'].apply(clean_store_code).map(cm_dict).fillna('Unmapped')

    # Pivot Stock
    if not valid_stock.empty:
        style_counts_stock = valid_stock.groupby(['store_code', 'Store Name', 'Style']).size().reset_index(name='Valid Options')
        df_style_stock = style_counts_stock.pivot_table(index=['store_code', 'Store Name'], columns='Style', values='Valid Options', aggfunc='sum', fill_value=0)
        df_style_stock.columns = [f"{c}_Stock" for c in df_style_stock.columns]
        df_style_stock = df_style_stock.reset_index()
        df_style_out = pd.merge(df_style_out, df_style_stock, on=['store_code', 'Store Name'], how='left')
        
    # Pivot Stock + Transit
    if not valid_st.empty:
        style_counts_st = valid_st.groupby(['store_code', 'Store Name', 'Style']).size().reset_index(name='Valid Options')
        df_style_st = style_counts_st.pivot_table(index=['store_code', 'Store Name'], columns='Style', values='Valid Options', aggfunc='sum', fill_value=0)
        df_style_st.columns = [f"{c}_Trans" for c in df_style_st.columns]
        df_style_st = df_style_st.reset_index()
        df_style_out = pd.merge(df_style_out, df_style_st, on=['store_code', 'Store Name'], how='left')
        
    # Pivot All (Stock + Transit + Alloc)
    if not valid_all.empty:
        style_counts_all = valid_all.groupby(['store_code', 'Store Name', 'Style']).size().reset_index(name='Valid Options')
        df_style_all = style_counts_all.pivot_table(index=['store_code', 'Store Name'], columns='Style', values='Valid Options', aggfunc='sum', fill_value=0)
        df_style_all.columns = [f"{c}_Allocation" for c in df_style_all.columns]
        df_style_all = df_style_all.reset_index()
        df_style_out = pd.merge(df_style_out, df_style_all, on=['store_code', 'Store Name'], how='left')
        
    # Fill any NaNs from the merge with 0 and convert to int
    df_style_out = df_style_out.fillna(0)
    style_cols = [c for c in df_style_out.columns if c not in ['store_code', 'Store Name', 'CM']]
    df_style_out[style_cols] = df_style_out[style_cols].astype(int)
    
    # Apply Store wise priority sorting
    df_store_out = sort_df_by_store_priority(df_store_out, 'store_code')
    df_store_ag_out = sort_df_by_store_priority(df_store_ag_out, 'store_code')
    df_style_out = sort_df_by_store_priority(df_style_out, 'store_code')
    
    # Apply Style wise priority sorting (reordering style columns)
    def get_style_col_sort_key(col_name):
        style_name = col_name
        suffix_rank = 3
        if col_name.endswith('_Stock'):
            style_name = col_name[:-6]
            suffix_rank = 1
        elif col_name.endswith('_Trans'):
            style_name = col_name[:-6]
            suffix_rank = 2
        elif col_name.endswith('_Allocation'):
            style_name = col_name[:-11]
            suffix_rank = 3
            
        pri = style_pri_dict.get(style_name.upper(), 999999)
        return (pri, style_name.upper(), suffix_rank)

    style_cols = [c for c in df_style_out.columns if c not in ['store_code', 'Store Name', 'CM']]
    style_cols_sorted = sorted(style_cols, key=get_style_col_sort_key)
    df_style_out = df_style_out[['store_code', 'Store Name', 'CM'] + style_cols_sorted]
    
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file = os.path.join(base_dir, "VALID STYLE OUTPUT", f"AG_Validation_Output_v2_{timestamp}.xlsx")
    os.makedirs(os.path.dirname(out_file), exist_ok=True)
    print(f"Writing to {out_file}...")
    
    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
        df_store_out_fmt.to_excel(writer, sheet_name='STORE WISE', index=False)
        df_ag_out.to_excel(writer, sheet_name='AG WISE', index=False)
        df_store_ag_out.to_excel(writer, sheet_name='STORE AG RAW', index=False)
        df_style_out.to_excel(writer, sheet_name='STYLE WISE', index=False)
        
        # Load and write CM sheet
        cm_file = os.path.join(base_dir, "AG VALIDATION DATA", "CM", "Name_Zones (2).xlsx")
        if os.path.exists(cm_file):
            try:
                df_cm = pd.read_excel(cm_file, sheet_name="Sheet1")
                df_cm.to_excel(writer, sheet_name='CM', index=False)
                print("Added CM sheet to output Excel.")
            except Exception as e:
                print(f"Could not add CM sheet to output Excel: {e}")
        
    print("Done!")
    
    try:
        import shutil
        shutil.copy(out_file, os.path.join(base_dir, "VALID STYLE OUTPUT", "AG_Validation_Output_v2_LATEST.xlsx"))
    except Exception as e:
        print(f"Could not save LATEST copy: {e}")

if __name__ == "__main__":
    main()
