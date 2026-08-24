import urllib.request, json, time

print("Testing /api/inwards_data ...")
start = time.time()
try:
    with urllib.request.urlopen('http://localhost:5050/api/inwards_data', timeout=30) as r:
        data = json.loads(r.read())
    elapsed = time.time() - start
    print(f"  Response time : {elapsed:.2f}s")
    print(f"  ok            : {data.get('ok')}")
    print(f"  records count : {len(data.get('records', []))}")
    print(f"  total_raw     : {data.get('total_raw')}")
    print(f"  file_time     : {data.get('file_time')}")
    print(f"  files_loaded  : {data.get('files_loaded')}")
    kpi = data.get('kpi', {})
    print(f"  KPI total_qty : {kpi.get('total_qty')}")
    print(f"  KPI styles    : {kpi.get('unique_styles')}")
    print(f"  KPI cats      : {kpi.get('unique_cats')}")
    print(f"  KPI date_from : {kpi.get('date_from')}")
    print(f"  KPI date_to   : {kpi.get('date_to')}")
    meta = data.get('filter_meta', {})
    print(f"  Categories    : {len(meta.get('categories', []))} items, first 5: {meta.get('categories', [])[:5]}")
    print(f"  Styles        : {len(meta.get('styles', []))} items")
    print(f"  Colors        : {len(meta.get('colors', []))} items")
    print(f"  Sizes         : {len(meta.get('sizes', []))} items")
    if data.get('records'):
        print(f"  Sample record : {json.dumps(data['records'][0], indent=4)}")
    if data.get('error'):
        print(f"  ERROR: {data['error']}")
        print(f"  TRACE: {data.get('trace','')[:500]}")
except Exception as e:
    print(f"  EXCEPTION: {e}")
