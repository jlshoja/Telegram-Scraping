"""Build a WooCommerce import sheet from Product Database - filled.xlsx.

Format reference (never modified): woocommerce_import.xlsx
  - parent row:  post_type=product, title/prices/categories/gallery + all attributes
  - child rows:  post_type=product_variation, one per color, only color attribute
  - output adds ONE new attribute pair for جنس:
        attribute:material / attribute_name:material (= "جنس")
    every other column is kept exactly as the template (order included).

Source columns are looked up BY HEADER NAME (immune to column moves):
  کد محصول, نام محصول , اسامی محصولات, دسته بندی,
  قیمت فروش, قیمت فروش تخفیف خورده, ابعاد, جنس, کاربرد, رنگ

رنگ format per color:  <name>-<stock>  joined by "|"   (e.g. "آبی-2 | سبز-3")
  - stock accepts Persian digits; missing number -> stock "ناموجود" + warning
  - colors without a known English slug are SKIPPED + reported
    (slugs are learned from the template file; extend EXTRA_SLUGS below)

Usage:
  python build_woocommerce.py
  python build_woocommerce.py --codes FA-1959,TR-402   # test subset
  python build_woocommerce.py --out woocommerce_import_new.xlsx
"""

import argparse
import collections
import re
import sys
from pathlib import Path

import openpyxl

from tgcommon import normalize

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SRC = "Product Database - filled.xlsx"
TEMPLATE = "woocommerce_import.xlsx"
OUT = "woocommerce_import_new.xlsx"

# Extra / override Persian-color -> English-slug pairs.
# Slugs are otherwise learned from the template file (first occurrence wins).
# Add new colors here instead of inventing slugs inline, e.g. "فیروزه‌ای": "turquoise"
EXTRA_SLUGS = {
    "عسلی": "honey",  # template also uses "cream" once; "honey" is the newer pattern
}

SALE_TAG = "تخفیف خورده‌ها"  # exact template string (with ZWNJ)

CODE_RE = re.compile(r"^([A-Z]+)-(\d+)$")
MISSING_NAMES = {"", "ندارد", "نامشخص"}


def col_index(header):
    """Map header name -> 0-based column index (first occurrence)."""
    return {v: i for i, v in enumerate(header) if v not in (None, "")}


def cell(row, H, name):
    i = H.get(name)
    if i is None or i >= len(row):
        return None
    v = row[i]
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    return v


def as_number(value):
    """Return int/float for numeric cells (or numeric strings), else None."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        try:
            t = value.strip().replace(",", "").replace("٬", "")
            return float(t) if "." in t else int(t)
        except ValueError:
            return None
    return None


def parse_colors(value):
    """Parse رنگ into [(color_name, stock_int|None)]."""
    out = []
    if value in (None, ""):
        return out
    for part in str(value).split("|"):
        part = part.strip().strip("،,")
        if not part:
            continue
        norm = normalize(part)
        if "-" in norm:
            name, _, qty = norm.rpartition("-")
            name, qty = name.strip(), qty.strip()
            if name and qty.isdigit():
                out.append((name, int(qty)))
                continue
        out.append((part, None))
    return out


def load_slug_map(template_rows, H):
    """Persian color -> English slug, learned from template variation SKUs."""
    slugs = {}
    for r in template_rows:
        if r[H["post_type"]] != "product_variation":
            continue
        color = r[H["attribute:color"]]
        if color in (None, ""):
            continue
        color = str(color).strip()
        sku, parent = str(r[H["sku"]]), str(r[H["parent_sku"]])
        prefix = parent + "-"
        if sku.startswith(prefix):
            slugs.setdefault(color, sku[len(prefix):])
    slugs.update(EXTRA_SLUGS)  # explicit overrides win
    return slugs


def main(args):
    twb = openpyxl.load_workbook(args.template, read_only=True)
    tws = twb.active
    tmpl_headers = [c.value for c in tws[1]]
    tH = {v: i for i, v in enumerate(tmpl_headers)}
    tmpl_rows = list(tws.iter_rows(min_row=2, values_only=True))
    twb.close()
    slugs = load_slug_map(tmpl_rows, tH)

    out_headers = tmpl_headers + ["attribute:material", "attribute_name:material"]
    oH = {v: i for i, v in enumerate(out_headers)}

    swb = openpyxl.load_workbook(args.src, read_only=True)
    sws = swb.active
    src_headers = [c.value for c in sws[1]]
    H = col_index(src_headers)
    rows = list(sws.iter_rows(min_row=2, values_only=True))
    swb.close()

    if "کد محصول" not in H:
        raise SystemExit("missing column: کد محصول")
    for name in ["نام محصول ", "اسامی محصولات", "قیمت فروش",
                 "قیمت فروش تخفیف خورده", "دسته بندی",
                 "ابعاد", "جنس", "کاربرد", "رنگ"]:
        if name not in H:
            print(f"WARNING: column missing, treated as empty: {name}")

    wanted = None
    if args.codes:
        wanted = {c.strip().upper() for c in args.codes.split(",") if c.strip()}

    out_rows = []
    stats = collections.Counter()
    warnings = []
    seen_codes = set()

    for excel_row, r in enumerate(rows, 2):
        raw_code = cell(r, H, "کد محصول")
        if raw_code in (None, ""):
            stats["no_code"] += 1
            continue
        code = normalize(str(raw_code)).strip().upper()
        if not CODE_RE.match(code):
            stats["invalid"] += 1
            continue
        if wanted is not None and code not in wanted:
            continue
        if code in seen_codes:
            warnings.append(f"row {excel_row}: duplicate code {code}, skipped")
            stats["duplicate"] += 1
            continue
        seen_codes.add(code)

        name = cell(r, H, "نام محصول ")
        if name in (None, "") or str(name).strip() in MISSING_NAMES:
            name = cell(r, H, "اسامی محصولات")
        if name in (None, "") or str(name).strip() in MISSING_NAMES:
            warnings.append(f"row {excel_row}: {code} has no product name, skipped")
            stats["no_name"] += 1
            continue
        title = f"{str(name).strip()} کد {code}"

        price = as_number(cell(r, H, "قیمت فروش"))
        if price is None:
            stats["no_price"] += 1
            continue
        sale = as_number(cell(r, H, "قیمت فروش تخفیف خورده"))
        tag = SALE_TAG if sale is not None else None

        cats = cell(r, H, "دسته بندی")
        if cats in (None, ""):
            categories = None
        else:
            parts = [p.strip() for p in str(cats).split("|") if p.strip()]
            for p in parts:
                if not p.startswith("کیف "):
                    warnings.append(
                        f"row {excel_row}: {code} unusual category value: {p!r}")
            categories = " | ".join(f"کیف زنانه>{p}" for p in parts) or None

        # --- attributes (only non-empty ones are written) ---
        colors = parse_colors(cell(r, H, "رنگ"))
        color_names, stocks = [], {}
        for cname, qty in colors:
            if cname in stocks:
                warnings.append(
                    f"row {excel_row}: {code} duplicate color {cname!r}, kept first")
                continue
            stocks[cname] = qty
            color_names.append(cname)
        if colors and not color_names:
            warnings.append(f"row {excel_row}: {code} رنگ has no usable colors")

        usage = cell(r, H, "کاربرد")
        if usage in (None, ""):
            usage = None
        else:
            usage = "|".join(
                p.strip() for p in str(usage).split("|") if p.strip()) or None
        material = cell(r, H, "جنس")
        dimension = cell(r, H, "ابعاد")
        if isinstance(dimension, (int, float)):
            dimension = str(dimension)

        def new_row():
            return [None] * len(out_headers)

        parent = new_row()
        parent[oH["post_type"]] = "product"
        parent[oH["post_title"]] = title
        parent[oH["post_status"]] = "publish"
        parent[oH["sku"]] = code
        parent[oH["regular_price"]] = price
        if sale is not None:
            parent[oH["sale_price"]] = sale
            parent[oH["sale_tag"]] = tag
        parent[oH["manage_stock"]] = "yes"
        if categories is not None:
            parent[oH["categories"]] = categories
        if color_names:
            parent[oH["attribute:color"]] = "|".join(color_names)
            parent[oH["attribute_name:color"]] = "رنگ"
        if usage is not None:
            parent[oH["attribute:application"]] = usage
            parent[oH["attribute_name:application"]] = "کاربرد"
        if material is not None:
            parent[oH["attribute:material"]] = material
            parent[oH["attribute_name:material"]] = "جنس"
        if dimension is not None:
            parent[oH["attribute:dimention"]] = dimension
            parent[oH["attribute_name:dimention"]] = "ابعاد"
        out_rows.append(parent)
        stats["products"] += 1

        for cname in color_names:
            slug = slugs.get(cname)
            if not slug:
                warnings.append(
                    f"row {excel_row}: {code} unknown color {cname!r}, "
                    f"variation skipped (add slug to EXTRA_SLUGS)")
                stats["unknown_slug"] += 1
                continue
            qty = stocks[cname]
            if qty is None:
                warnings.append(
                    f"row {excel_row}: {code} color {cname!r} has no stock number, "
                    f"used 'ناموجود'")
                qty = "ناموجود"
            v = new_row()
            v[oH["post_type"]] = "product_variation"
            v[oH["post_status"]] = "publish"
            v[oH["sku"]] = f"{code}-{slug}"
            v[oH["parent_sku"]] = code
            v[oH["regular_price"]] = price
            if sale is not None:
                v[oH["sale_price"]] = sale
                v[oH["sale_tag"]] = tag
            v[oH["manage_stock"]] = "yes"
            v[oH["stock_quantity"]] = qty
            v[oH["attribute:color"]] = cname
            v[oH["attribute_name:color"]] = "رنگ"
            out_rows.append(v)
            stats["variations"] += 1

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(out_headers)
    for row in out_rows:
        ws.append(row)
    wb.save(args.out)

    print(f"saved {args.out}: {stats['products']} products, "
          f"{stats['variations']} variations")
    print(f"skipped: no_code={stats['no_code']} invalid={stats['invalid']} "
          f"no_name={stats['no_name']} no_price={stats['no_price']} "
          f"duplicate={stats['duplicate']} unknown_slug={stats['unknown_slug']}")
    for w in warnings:
        print("WARNING:", w)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--codes", default="",
                    help="only process these product codes, comma-separated (for testing)")
    ap.add_argument("--src", default=SRC)
    ap.add_argument("--template", default=TEMPLATE)
    ap.add_argument("--out", default=OUT)
    main(ap.parse_args())
