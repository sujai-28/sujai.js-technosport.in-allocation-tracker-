"""Remove the stale old inwards functions that were left over in app.py"""
path = r"D:\INCREFF ORDER PUNCH\app.py"
with open(path, encoding='utf-8') as f:
    lines = f.readlines()

# Find the lines
start_marker = None
end_marker = None
for i, line in enumerate(lines):
    # Find the stale docstring that's now floating without a def
    if '"""Read all CSV/XLSX files under INWARDS_DIR' in line and start_marker is None:
        # Go back a couple lines to find the def (or floating lines)
        start_marker = max(0, i - 3)
    if start_marker and 'return jsonify({\'ok\': False, \'error\': str(e)})' in line and i > start_marker + 50:
        end_marker = i + 1
        break

if start_marker is None or end_marker is None:
    print(f"Markers: start={start_marker}, end={end_marker}")
    # Just print first match context
    for i, l in enumerate(lines):
        if 'INWARDS_DIR' in l or '_load_inwards_raw' in l or 'all subfolders' in l:
            print(i+1, repr(l[:80]))
else:
    print(f"Removing lines {start_marker+1}–{end_marker} ({end_marker - start_marker} lines)")
    # Preview what we're removing
    for i in range(start_marker, min(start_marker+5, end_marker)):
        print(f"  {i+1}: {lines[i][:80].rstrip()}")
    print("  ...")
    for i in range(max(start_marker, end_marker-3), end_marker):
        print(f"  {i+1}: {lines[i][:80].rstrip()}")

    new_lines = lines[:start_marker] + ['\n'] + lines[end_marker:]
    with open(path, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    print("Done.")
