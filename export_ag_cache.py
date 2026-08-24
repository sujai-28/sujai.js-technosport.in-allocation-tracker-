"""
export_ag_cache.py
==================
Run this script locally AFTER running validate_styles.py.
It processes all Excel data and saves a lightweight JSON cache
that the cloud app reads instead of processing Excel files directly.

Usage:
    python export_ag_cache.py
"""
import os, sys, json, gc
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

print("="*55)
print("  AG Validation Data -> JSON Cache Exporter")
print("="*55)

# Bootstrap _state and load AG data using existing app logic
from app import _load_ag_data_internal, _state

print("\n[1/3] Loading and processing AG Validation Excel files...")
ok, err = _load_ag_data_internal()

if not ok:
    print(f"\n❌ Failed to load AG data: {err}")
    sys.exit(1)

print(f"[2/3] Loaded: {len(_state['ag_store_data'])} stores, "
      f"{len(_state['ag_wise_data'])} AGs, "
      f"{len(_state.get('ag_raw_data', []))} raw rows, "
      f"{len(_state.get('ag_style_data', []))} style rows")

# Build the cache payload
cache = {
    "status": _state["ag_status"],
    "store_data": _state["ag_store_data"],
    "ag_data": _state["ag_wise_data"],
    "raw_data": _state.get("ag_raw_data", []),
    "style_data": _state.get("ag_style_data", []),
    "style_cols": _state.get("ag_style_cols", []),
    "store_cols": [
        'store_code', 'Store Name', 'Sum of Final Options',
        'Sum of Final Options - M', 'Sum of Final Options - W', 'Sum of Final Options - K',
        'Only Stock Valid Options', 'Only Stock Valid Options - M',
        'Only Stock Valid Options - W', 'Only Stock Valid Options - K',
        'Stock and Transit Valid Options', 'Stock and Transit Valid Options - M',
        'Stock and Transit Valid Options - W', 'Stock and Transit Valid Options - K',
        'All Valid Options', 'All Valid Options - M',
        'All Valid Options - W', 'All Valid Options - K'
    ],
    "ag_cols": ['AG Name', 'Sum of Final Options', 'Only Stock Valid Options',
                'Stock and Transit Valid Options', 'All Valid Options']
}

out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ag_data_cache.json")

print(f"[3/3] Saving JSON cache to: {out_path}")
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(cache, f, ensure_ascii=False, default=str)

size_kb = os.path.getsize(out_path) / 1024
print(f"\n✅ Done! Cache saved: {size_kb:.1f} KB")
print(f"   Now run:  git add ag_data_cache.json && git commit -m 'Update AG cache' && git push")
print("   The cloud app will automatically use this lightweight JSON file.")
