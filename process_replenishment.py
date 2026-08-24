import pandas as pd
import os

# Configuration: adjust these if your actual column names differ
INPUT_FILE = r"DATA\Requirement as of 06.07.2026.xlsx"

# Sheet 1 config
SHEET1_NAME = "Sheet1"
EBO_COLUMN_NAME = "EBO NAME"
# Map the columns from Sheet1 to the expected output columns
# Format: "Output Column Name": "Sheet1 Column Name"
COLUMN_MAPPING = {
    "clientSkuId": "Client SKU Id / EAN",
    "orderedQuantity": "Allocated Qty",
    "sellingPrice": "sp"
}

# Sheet 2 config
SHEET2_NAME = "Sheet2"  # Note: no trailing space in actual sheet name
ROW_LABEL_COL = "Row Labels"

def process_replenishment_data():
    import sys
    input_file = sys.argv[1] if len(sys.argv) > 1 else INPUT_FILE
    if not os.path.exists(input_file):
        print(f"Error: {input_file} not found.")
        return

    print("Loading data...")
    try:
        # Check sheet names dynamically
        xl = pd.ExcelFile(input_file)
        sheet_names = xl.sheet_names

        # Load Sheet 1 (df_sheet1)
        if SHEET1_NAME in sheet_names:
            df_sheet1 = pd.read_excel(input_file, sheet_name=SHEET1_NAME)
        else:
            df_sheet1 = pd.read_excel(input_file, sheet_name=0)

        # Normalize replenishment columns
        col_mapping = {
            "EBO NAME": ["EBO NAME", "ebo name", "Ebo Name", "EBO Name", "STORE NAME", "store name", "Store Name"],
            "Client SKU Id / EAN": ["Client SKU Id / EAN", "client sku id / ean", "Client Sku Id / Ean", "Client SKU ID / EAN", "SKU", "sku", "Sku"],
            "Allocated Qty": ["Allocated Qty", "allocated qty", "Allocated QTY", "Allocated quantity", "allocated quantity", "Allocated  Qty"],
            "sp": ["sp", "selling price", "Selling Price", "SP"]
        }
        rename_map = {}
        for standard_col, options in col_mapping.items():
            for col in df_sheet1.columns:
                normalized_col = " ".join(col.split()).lower()
                normalized_std = " ".join(standard_col.split()).lower()
                normalized_opts = [" ".join(opt.split()).lower() for opt in options]
                if normalized_col == normalized_std or normalized_col in normalized_opts:
                    rename_map[col] = standard_col
                    break
        df_sheet1 = df_sheet1.rename(columns=rename_map)

        # Remove duplicates based on Store Name, SKU, and Allocated Qty
        subset_cols = [col for col in ["EBO NAME", "Client SKU Id / EAN", "Allocated Qty"] if col in df_sheet1.columns]
        if subset_cols:
            initial_len = len(df_sheet1)
            df_sheet1 = df_sheet1.drop_duplicates(subset=subset_cols)
            if len(df_sheet1) < initial_len:
                print(f"Removed {initial_len - len(df_sheet1)} duplicate rows based on Store, SKU, and Qty.")

        # Load Sheet 2 (df_sheet2) if it exists
        df_sheet2 = None
        if SHEET2_NAME in sheet_names:
            df_sheet2_raw = pd.read_excel(input_file, sheet_name=SHEET2_NAME, header=None)
            df_sheet2_raw.columns = ["Row Labels", "Sum of Allocated Qty"]
            df_sheet2 = df_sheet2_raw[
                df_sheet2_raw["Row Labels"].notna() &
                (df_sheet2_raw["Row Labels"] != "Row Labels")
            ].reset_index(drop=True)
    except Exception as e:
        print(f"Error reading excel file: {e}")
        return

    # Extract unique Row Labels
    if EBO_COLUMN_NAME not in df_sheet1.columns:
        print(f"Error: Column '{EBO_COLUMN_NAME}' not found in Sheet 1")
        return

    if df_sheet2 is not None and ROW_LABEL_COL in df_sheet2.columns:
        row_labels = df_sheet2[ROW_LABEL_COL].dropna().unique()
    else:
        # Fallback to Sheet1's unique EBO names if Sheet2 is not present or doesn't have the column
        row_labels = df_sheet1[EBO_COLUMN_NAME].dropna().unique()

    # Process each row label
    processed_count = 0
    for label in row_labels:
        # Filter Sheet1 where EBO NAME matches the current Row Label
        filtered_df = df_sheet1[df_sheet1[EBO_COLUMN_NAME] == label]
        
        if not filtered_df.empty:
            # Create a new dataframe with just the columns we need, mapped to the correct names
            output_data = {}
            for out_col, in_col in COLUMN_MAPPING.items():
                if in_col in filtered_df.columns:
                    output_data[out_col] = filtered_df[in_col]
                else:
                    print(f"Warning: Column '{in_col}' not found in {SHEET1_NAME}. Filling '{out_col}' with blanks.")
                    output_data[out_col] = "" # Fill with empty string if column doesn't exist
            
            output_df = pd.DataFrame(output_data)
            
            # Pivot/Group by SKU because SKUs can be repeated
            if not output_df.empty and "clientSkuId" in output_df.columns:
                output_df["orderedQuantity"] = pd.to_numeric(output_df["orderedQuantity"], errors='coerce').fillna(0)
                output_df = output_df.groupby("clientSkuId", as_index=False).agg({
                    "orderedQuantity": "sum",
                    "sellingPrice": "first"
                })
                # Ensure column order matches the mapping
                output_df = output_df[["clientSkuId", "orderedQuantity", "sellingPrice"]]
            
            # Save to CSV
            output_filename = f"SAMPLE OUTWARD ORDER {label}.csv"
            
            # Clean up the filename to avoid invalid characters (like slashes)
            output_filename = "".join(c if c.isalnum() or c in (' ', '-', '_', '.') else '_' for c in output_filename)
            output_path = os.path.join("OUTPUTFILE", output_filename)
            
            try:
                output_df.to_csv(output_path, index=False)
                print(f"Created: {output_path}")
                processed_count += 1
            except Exception as e:
                print(f"Error writing to {output_filename}: {e}")

    print(f"\nDone! Generated {processed_count} CSV files.")

if __name__ == "__main__":
    process_replenishment_data()
