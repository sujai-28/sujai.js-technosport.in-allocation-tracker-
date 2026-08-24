"""
TechnoSport BOH Replenishment Dashboard - static site builder.

Reads the two daily CSV extracts and emits a single self-contained
public/index.html with the data embedded. No server, no database.

Usage (from the dashboard/ folder):
    python build_dashboard.py
    python build_dashboard.py --sales "..\\SALES.csv" --stock "..\\STOCK.csv"

Defaults look for the newest matching CSV in the parent folder.
"""

import argparse
import base64
import csv
import glob
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
PARENT = os.path.dirname(HERE)

# Brand artwork is inlined as data URIs rather than shipped as files: the page
# has to survive being served as a JS string from behind the password gate, and
# to keep working offline once a store saves it to a tablet home screen.
ASSETS = os.path.join(HERE, "assets")
LOGO_PNG = os.path.join(ASSETS, "technosport-logo.png")   # wordmark, transparent
ICON_PNG = os.path.join(ASSETS, "technosport-icon.png")   # square, home-screen


def data_uri(path):
    if not os.path.exists(path):
        sys.exit(
            f"ERROR: brand asset missing: {path}\n"
            f"       The page inlines it, so the build cannot continue without it."
        )
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode("ascii")

# --- size scales -----------------------------------------------------------
# Each department uses one of a few size runs. Rendering all 30 distinct sizes
# as columns is unusable on a phone, so rows are tabbed by scale. A style that
# spans two scales (e.g. boys tees sold in both 12Y and MED) appears in both.
ADULT_SIZES = ["SML", "MED", "LAR", "XLR", "2XL", "3XL", "4XL", "5XL"]
KIDS_SIZES = ["06Y", "08Y", "10Y", "12Y", "14Y"]
FOOTWEAR_SIZES = ["6", "7", "8", "9", "10", "11"]

SCALES = {
    "adult": ADULT_SIZES,
    "kids": KIDS_SIZES,
    "footwear": FOOTWEAR_SIZES,
}

# Packaging / consumables billed as line items. Not merchandise, never
# replenished from BOH - one bill line can carry 100 units and would swamp
# the "top sellers" sort.
EXCLUDE_SKU_PREFIXES = ("MAKRAFTBAG",)


def size_scale(size):
    for name, members in SCALES.items():
        if size in members:
            return name
    return "other"


def norm(v):
    return (v or "").strip()


def to_int(v):
    v = norm(v)
    if not v:
        return 0
    try:
        return int(float(v))
    except ValueError:
        return 0


def sku_stem(sku, size):
    """Strip the trailing size (and any variant suffix) off a SKU.

    MTP864IRGLAR -> MTP864IRG ; MTOR38TLE2XL004 -> MTOR38TLE ;
    MTOR13BLKLARNEW -> MTOR13BLK. Returns None when the pattern doesn't
    hold (mostly accessories with free-text sizes).
    """
    s = sku
    for suffix in ("NEW",):
        if s.endswith(suffix) and not size.endswith(suffix):
            s = s[: -len(suffix)]
    if not s.endswith(size):
        m = re.match(r"^(.*?)\d{3}$", s)
        if m:
            s = m.group(1)
    if s.endswith(size):
        return s[: -len(size)] or None
    return None


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def colour_name_map(sales):
    """One colour NAME per colour code, chain-wide, decided by majority vote.

    The stockboy reads "Ocean Waves", not "ONW", so the name matters more here
    than anywhere else on the page. Two problems in the source make it unsafe
    to just take the name off whichever line happens to come first:

      * 64 lines carry a real colour code but COLOR "0", and six codes
        (BEB, BLACK, DCO, DMT, GST, GYB) have no name anywhere in the file -
        those fall back to the code.
      * eight codes are spelled several ways ("Navy B" / "Navy Blue",
        "Chili Flakes" / "Chilli Flakes"). One code must print as one name or
        the same garment appears twice in a sorted list.

    Voting across the whole chain also fills a store's blanks from other
    stores' lines, which a per-store lookup could not do.
    """
    votes = defaultdict(Counter)
    for r in sales:
        code = norm(r.get("COLOR CODE"))
        name = norm(r.get("COLOR"))
        if code and name and name != "0":
            votes[code][name.title()] += 1
    return {code: names.most_common(1)[0][0] for code, names in votes.items()}


def resolve(explicit, patterns, label):
    if explicit:
        if not os.path.exists(explicit):
            sys.exit(f"ERROR: {label} file not found: {explicit}")
        return explicit
    hits = []
    for pat in patterns:
        hits.extend(glob.glob(os.path.join(PARENT, pat)))
    if not hits:
        sys.exit(f"ERROR: no {label} CSV found in {PARENT} (looked for {patterns})")
    hits.sort(key=os.path.getmtime, reverse=True)
    return hits[0]


# --- stock file schemas ----------------------------------------------------
# Stock arrives in two different layouts depending on which report was run.
# The trap is colour: the chain-wide report's "Colour" holds the CODE, while
# the per-store report's "COLOR" holds the full NAME and the code lives in
# "CCODE". Joining on the wrong one silently matches almost nothing.
STOCK_SCHEMAS = [
    {
        "id": "chain",
        "detect": "store_code",
        "store": "store_code", "name": "Store Name", "dept": "Department",
        "sku": "SKU", "style": "Style", "colour": "Colour", "size": "Size",
        "qty": "quantity", "status": None, "day": "day", "barcode": None,
    },
    {
        "id": "per-store",
        "detect": "SITE_CODE",
        "store": "SITE_CODE", "name": "OWNER SITE", "dept": "DEPARTMENT",
        "sku": "SKU", "style": "STYLE", "colour": "CCODE", "size": "SIZE",
        "qty": "STOCK QUANTITY", "status": "STOCK STATUS", "day": None,
        "barcode": "BARCODE",
    },
]

# Only stock physically in the store can be pulled to the floor today.
# "Transit" is on its way and is tracked separately so a zero on the shelf
# can be reported as "arriving" rather than "reorder".
AVAILABLE_STATUS = "store"


def load_stock(paths):
    """Read every stock export into one normalised list of rows."""
    rows, report = [], []
    for path in paths:
        raw = read_csv(path)
        if not raw:
            report.append((os.path.basename(path), "EMPTY", 0, 0, 0, 0))
            continue

        header = raw[0].keys()
        schema = next((s for s in STOCK_SCHEMAS if s["detect"] in header), None)
        if schema is None:
            sys.exit(
                f"ERROR: unrecognised stock layout in {os.path.basename(path)}.\n"
                f"       columns: {list(header)[:8]}...\n"
                f"       Expected a column named 'store_code' or 'SITE_CODE'."
            )

        # Per-store exports carry no date column; the file's own timestamp is
        # the best available statement of how fresh the count is.
        fallback_day = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d-%m-%Y")

        n_avail = n_transit = n_nonpos = 0
        for r in raw:
            qty = to_int(r.get(schema["qty"]))
            status = norm(r.get(schema["status"])) if schema["status"] else "Store"
            store = norm(r.get(schema["store"]))
            if not store:
                continue
            if qty <= 0:
                # Negative or zero on-hand is a stock-ledger artefact, not
                # something anyone can pull. Never let it net away real stock.
                n_nonpos += 1
                continue
            available = status.lower() == AVAILABLE_STATUS
            if available:
                n_avail += qty
            else:
                n_transit += qty
            rows.append({
                "store": store,
                "name": norm(r.get(schema["name"])),
                "dept": norm(r.get(schema["dept"])),
                "sku": norm(r.get(schema["sku"])),
                "style": norm(r.get(schema["style"])),
                "colour": norm(r.get(schema["colour"])) or "-",
                "size": norm(r.get(schema["size"])),
                "qty": qty,
                "available": available,
                "day": norm(r.get(schema["day"])) if schema["day"] else fallback_day,
            })
        stores = sorted({norm(r.get(schema["store"])) for r in raw if norm(r.get(schema["store"]))})
        report.append((os.path.basename(path), schema["id"], len(raw),
                       n_avail, n_transit, n_nonpos, stores))
    return rows, report


def build(sales_path, stock_paths, out_path):
    sales = read_csv(sales_path)
    stock, stock_report = load_stock(stock_paths)

    warnings = []
    colour_names = colour_name_map(sales)

    # --- stores ------------------------------------------------------------
    store_names = {}
    for r in stock:  # stock naming is the canonical "TSPL <CITY> EBO" form
        if r["store"] and r["name"]:
            store_names[r["store"]] = r["name"]
    for r in sales:
        code = norm(r["Store Code"])
        if code and code not in store_names:
            store_names[code] = norm(r["EBO NAME"])

    stock_store_codes = {r["store"] for r in stock}

    # --- aggregate sales ---------------------------------------------------
    # grain: store x style x colour-code ; per-size net quantity
    agg = {}
    returns_units = 0
    excluded_lines = 0
    skipped_lines = 0

    def slot(store, style, colour_code):
        key = (store, style, colour_code)
        if key not in agg:
            agg[key] = {
                "store": store,
                "style": style,
                "colour": colour_code,
                "colour_name": "",
                "dept": "",
                "stem": "",
                "sold": defaultdict(int),
                "stock": defaultdict(int),
                "transit": defaultdict(int),
                "sku": {},
                "barcode": {},
                "returns": 0,
                "rsp": 0,
            }
        return agg[key]

    for r in sales:
        sku = norm(r["SKU"])
        style = norm(r["STYLE"])
        store = norm(r["Store Code"])
        size = norm(r["SIZE"])
        qty = to_int(r["BILL_QUANTITY"])

        if not store or not sku or not style:
            skipped_lines += 1
            continue
        if sku.upper().startswith(EXCLUDE_SKU_PREFIXES):
            excluded_lines += 1
            continue
        if norm(r["VOID STATUS"]).lower() not in ("", "no"):
            skipped_lines += 1
            continue

        colour_code = norm(r["COLOR CODE"]) or "-"
        row = slot(store, style, colour_code)
        row["sold"][size] += qty
        if qty < 0:
            row["returns"] += -qty
            returns_units += -qty

        colour_name = norm(r["COLOR"])
        if colour_name and colour_name not in ("0",) and not row["colour_name"]:
            row["colour_name"] = colour_name.title()
        if not row["dept"]:
            row["dept"] = norm(r["DEPARTMENT"]) or "-"
        if size and sku:
            row["sku"][size] = sku
            row["barcode"][size] = norm(r["BARCODE"])
        rsp = to_int(r["RSP"])
        if rsp > row["rsp"]:
            row["rsp"] = rsp
        stem = sku_stem(sku, size)
        if stem and not row["stem"]:
            row["stem"] = stem

    # --- join stock --------------------------------------------------------
    # Full size run for every style+colour that sold, not just the sizes that
    # sold - the stockboy needs to see the gaps beside the movers.
    stock_index = defaultdict(int)
    transit_index = defaultdict(int)
    stock_sku = {}
    for r in stock:
        k = (r["store"], r["style"], r["colour"], r["size"])
        if r["available"]:
            stock_index[k] += r["qty"]
        else:
            transit_index[k] += r["qty"]
        if r["sku"]:
            stock_sku[k] = r["sku"]

    for (store, style, colour, size), qty in stock_index.items():
        key = (store, style, colour)
        if key in agg:
            agg[key]["stock"][size] += qty

    for (store, style, colour, size), qty in transit_index.items():
        key = (store, style, colour)
        if key in agg:
            agg[key]["transit"][size] += qty

    # --- emit rows ---------------------------------------------------------
    rows = []
    for row in agg.values():
        sold = {k: v for k, v in row["sold"].items() if v > 0}
        if not sold:
            continue  # pure-return or fully reversed line
        stock_map = {k: v for k, v in row["stock"].items() if v > 0}
        transit_map = {k: v for k, v in row["transit"].items() if v > 0}

        sizes = set(sold) | set(stock_map)
        scales = sorted({size_scale(s) for s in sizes})

        # fill in SKUs for stocked sizes that didn't sell, so every cell scans
        skus = dict(row["sku"])
        for s in stock_map:
            if s not in skus:
                got = stock_sku.get((row["store"], row["style"], row["colour"], s))
                if got:
                    skus[s] = got

        rows.append(
            {
                "s": row["store"],
                "y": row["style"],
                "c": row["colour"],
                "cn": colour_names.get(row["colour"]) or row["colour_name"] or row["colour"],
                "d": row["dept"],
                "k": row["stem"] or "",
                "g": scales,
                "p": row["rsp"],
                "r": row["returns"],
                "so": sold,
                "st": stock_map,
                "tr": transit_map,
                "sk": skus,
                "bc": {k: v for k, v in row["barcode"].items() if v},
                "ts": sum(sold.values()),
                "tk": sum(stock_map.values()),
                "tt": sum(transit_map.values()),
            }
        )

    rows.sort(key=lambda r: (-r["ts"], r["y"]))

    # --- per-store summary -------------------------------------------------
    stores = []
    for code, name in store_names.items():
        srows = [r for r in rows if r["s"] == code]
        if not srows:
            continue
        stores.append(
            {
                "c": code,
                "n": name,
                "hs": code in stock_store_codes,
                "u": sum(r["ts"] for r in srows),
                "l": len(srows),
                "z": sum(1 for r in srows if r["tk"] == 0),
                "tr": sum(r["tt"] for r in srows),
            }
        )
    stores.sort(key=lambda s: s["n"])

    sales_dates = {norm(r["BILL_DATE"]) for r in sales if norm(r["BILL_DATE"])}
    stock_dates = {r["day"] for r in stock if r["day"]}
    if len(sales_dates) > 1:
        warnings.append(f"Sales file spans {len(sales_dates)} dates: {sorted(sales_dates)}")

    no_stock = [s["n"] for s in stores if not s["hs"]]
    if no_stock:
        warnings.append(
            f"{len(no_stock)} of {len(stores)} stores have no rows in the stock file - "
            "their stock columns will read 'no data'."
        )

    has_transit = any(r["tt"] for r in rows)

    payload = {
        "salesDate": sorted(sales_dates)[0] if sales_dates else "",
        "stockDate": sorted(stock_dates)[0] if stock_dates else "",
        "hasTransit": has_transit,
        "builtAt": datetime.now().strftime("%d-%m-%Y %H:%M"),
        "scales": SCALES,
        "stores": stores,
        "rows": rows,
        "meta": {
            "salesLines": len(sales),
            "stockLines": len(stock),
            "returnUnits": returns_units,
            "excludedLines": excluded_lines,
            "skippedLines": skipped_lines,
            "storesWithStock": len(stock_store_codes),
        },
    }

    tpl_path = os.path.join(HERE, "template.html")
    with open(tpl_path, encoding="utf-8") as f:
        html = f.read()

    blob = json.dumps(payload, separators=(",", ":"), ensure_ascii=False)
    html = html.replace("/*__DATA__*/null", blob)
    html = html.replace("__LOGO__", data_uri(LOGO_PNG))
    html = html.replace("__ICON__", data_uri(ICON_PNG))

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    # The Vercel deployment serves the page from behind a password, so the HTML
    # ships as a JS module the function returns - never as a static file, which
    # would be readable without authenticating.
    page_js = os.path.join(HERE, "deploy", "api", "page.js")
    if os.path.isdir(os.path.dirname(page_js)):
        with open(page_js, "w", encoding="utf-8") as f:
            f.write("// GENERATED by build_dashboard.py - do not edit.\n")
            f.write("module.exports = " + json.dumps(html, ensure_ascii=False) + ";\n")

    # --- report ------------------------------------------------------------
    print(f"sales : {os.path.basename(sales_path)}")
    print("stock :")
    for entry in stock_report:
        fn, sid, n, avail, transit, nonpos = entry[:6]
        file_stores = entry[6] if len(entry) > 6 else []
        print(f"        {fn}")
        print(f"          layout={sid}  rows={n}  stores={','.join(file_stores)}")
        print(f"          in-store units={avail}  in-transit units={transit}  "
              f"non-positive rows skipped={nonpos}")
    print(f"output: {out_path}  ({os.path.getsize(out_path)/1024:.0f} KB)")
    print()
    print(f"  sales date        {payload['salesDate']}")
    print(f"  stock date        {payload['stockDate']}")
    print(f"  stores with sales {len(stores)}")
    print(f"  stores with stock {len(stock_store_codes)}")
    print(f"  style/colour rows {len(rows)}")
    print(f"  units sold (net)  {sum(r['ts'] for r in rows)}")
    print(f"  return units      {returns_units}")
    print(f"  packaging lines excluded {excluded_lines}")
    print(f"  malformed lines skipped  {skipped_lines}")
    print(f"  colour codes named       {len(colour_names)}")
    unnamed = sorted({r["c"] for r in rows if r["cn"] == r["c"]})
    if unnamed:
        print(f"  colour codes with no name, shown as the code: {', '.join(unnamed)}")
    for w in warnings:
        print(f"  ! {w}")


def resolve_stock(explicit):
    """Every stock export in the folder is used - one per store, or one for
    the whole chain, in either layout. More files simply means more coverage."""
    if explicit:
        missing = [p for p in explicit if not os.path.exists(p)]
        if missing:
            sys.exit(f"ERROR: stock file(s) not found: {missing}")
        return list(explicit)
    hits = set()
    for pat in ("*Stock*.csv", "*STOCK*.csv", "*stock*.csv"):
        hits.update(glob.glob(os.path.join(PARENT, pat)))
    if not hits:
        sys.exit(f"ERROR: no stock CSV found in {PARENT}")
    return sorted(hits)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--sales", help="path to the SKU-wise bill-wise sales CSV")
    ap.add_argument("--stock", nargs="+",
                    help="one or more stock report CSVs (default: every *Stock*.csv found)")
    ap.add_argument("--out", default=os.path.join(HERE, "public", "index.html"))
    a = ap.parse_args()

    build(
        resolve(a.sales, ["*SALES*.csv", "*sales*.csv"], "sales"),
        resolve_stock(a.stock),
        a.out,
    )
