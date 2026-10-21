"""Build a WooCommerce import sheet from Product Database - filled.xlsx.

Format reference (template only, never modified):
  woocommerce_import_new.xlsx
  - parent row:  post_type=product, title/prices/categories/gallery + all attributes
  - child rows:  post_type=product_variation, one per color, only color attribute
  - if the template lacks the جنس pair, it is appended:
        attribute:material / attribute_name:material (= "جنس")
    every other column is kept exactly as the template (order included).

Source columns are looked up BY HEADER NAME (immune to column moves):
  کد محصول, نام محصول , اسامی محصولات, دسته بندی,
  قیمت فروش, قیمت فروش تخفیف خورده, ابعاد, جنس, کاربرد, رنگ

رنگ format per color:  <name> <stock>  joined by "|"
  e.g. "آبی ۱ | قرمز ۲"  or  "مشکی۳ | قهوه‌ای۳"
  - stock accepts Persian digits; missing number -> stock "ناموجود" + warning
  - colors without a known English slug are SKIPPED + reported
    (slugs are learned from the template file, then color_mapping.xlsx)

Images:
  Product Images/<product code>/<file>.webp
  - source names may contain the Persian or English color, e.g.
      main-آبی.webp, 1-آبی.webp, 2-آبی.webp   or   main-blue.webp, blue01.webp
  - files are renamed to the English convention:
      <code>-main-<color>.webp   (featured image)
      <code>-<color>01.webp      (gallery image 1, 2, ...)
  - every non-main file becomes a gallery image (gallery_images column)
  - a variation gets the first gallery image for its color; otherwise the first
    gallery image is used.
  - color_mapping.xlsx is re-read on every run, so new rows are picked up
    automatically (missing colors are reported in the console).

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
from urllib.parse import quote

import openpyxl

from tgcommon import normalize

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SRC = "Product Database - filled.xlsx"
TEMPLATE = "woocommerce_import_new.xlsx"
OUT = "woocommerce_import_new.xlsx"
COLOR_MAP_FILE = "color_mapping.xlsx"
IMAGES_DIR = "Product Images"
IMAGE_URL_BASE = "https://luxbaz.com/wp-content/uploads/2026/10/"
IMAGE_EXTS = {".webp", ".jpg", ".jpeg", ".png", ".gif", ".avif"}
IMAGE_STOPWORDS = {
    "the", "and", "with", "for", "light", "dark", "plain", "fabric", "stone",
    "charcoal", "design", "pattern", "hot", "stamped", "small", "large",
    "logo", "print", "shrink",
}
COLOR_ALIASES = {
    "سورمه ای": "سرمه ای",
    "سورمه": "سرمه ای",
    "کرمی": "کرم",
}

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


def color_key(value):
    """Normalize a Persian color name for lookups (ZWNJ/dash/spaces)."""
    key = normalize(str(value or "")).replace("\u200c", " ").replace("-", " ")
    return re.sub(r"\s+", " ", key).strip()


def parse_colors(value):
    """Parse رنگ into [(color_name, stock_int|None)].

    Accepted: "آبی ۱ | قرمز ۲", "مشکی۳ | قهوه‌ای۳", "آبی-2", "سبز".
    """
    out = []
    if value in (None, ""):
        return out
    for part in str(value).split("|"):
        part = part.strip().strip("،,")
        if not part:
            continue
        norm = normalize(part).strip()
        m = re.match(r"^(.*?)\s*-?\s*(\d+)\s*$", norm)
        if m:
            name = m.group(1).strip().strip("-").strip()
            if name:
                out.append((name, int(m.group(2))))
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
        sku, parent = str(r[H["sku"]]), str(r[H["parent_sku"]])
        prefix = parent + "-"
        if sku.startswith(prefix):
            slugs.setdefault(color_key(color), sku[len(prefix):])
    slugs.update({color_key(k): v for k, v in EXTRA_SLUGS.items()})
    return slugs


def load_color_map(path):
    """Load Persian -> English colors from color_mapping.xlsx (fresh each run)."""
    p = Path(path)
    if not p.exists():
        print(f"WARNING: color mapping file not found: {p}")
        return {}
    wb = openpyxl.load_workbook(p, read_only=True)
    ws = wb.active
    mapping = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or len(r) < 2:
            continue
        fa, en = r[0], r[1]
        if fa in (None, "") or en in (None, ""):
            continue
        key = color_key(fa)
        if key:
            mapping.setdefault(key, str(en).strip().strip('"').strip())
    wb.close()
    return mapping


def lookup_color(mapping, name):
    """Find the English name for a Persian color (exact, alias, prefix)."""
    key = color_key(name)
    if not key:
        return None
    if key in mapping:
        return mapping[key]

    alias = COLOR_ALIASES.get(key)
    if alias and alias in mapping:
        return mapping[alias]
    lookup_key = alias or key

    # Color is more specific than a mapping key: "سبز زمردی" -> "سبز".
    longer = [k for k in mapping if len(k) >= 3 and lookup_key.startswith(k)]
    if longer:
        return mapping[max(longer, key=len)]

    # Color is an abbreviation: "نسکافه" -> "نسکافه ای".
    shorter = [k for k in mapping if len(lookup_key) >= 3 and k.startswith(lookup_key)]
    if shorter:
        return mapping[min(shorter, key=len)]
    return None


def slugify(value):
    s = str(value or "").lower().replace("\u200c", " ")
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")


def image_file_slug(path):
    return (slugify(path.stem) or "image") + path.suffix.lower()


def image_url(path, base):
    return base.rstrip("/") + "/" + quote(image_file_slug(path))


def natural_key(value):
    parts = re.split(r"(\d+)", normalize(str(value)).lower())
    return [(0, int(p)) if p.isdigit() else (1, p) for p in parts]


def is_main_filename(path):
    stem = path.stem.lower()
    return bool(re.match(r"^main([-_\s]|$)", stem) or "-main-" in stem)


def detect_filename_color(mapping, stem):
    """Detect a color in a filename (Persian keys or English values).

    The earliest color in the name wins, so "black-leather-on-white" is black.
    """
    text = color_key(stem)
    slug = slugify(stem)
    matches = []

    for key, english in mapping.items():
        if not key or not english:
            continue
        index = text.find(key)
        if index >= 0:
            matches.append((index, -len(key), english))

    seen_values = set()
    for english in mapping.values():
        if not english:
            continue
        value_slug = slugify(english)
        if not value_slug or value_slug in seen_values:
            continue
        seen_values.add(value_slug)
        match = re.search(
            rf"(^|-){re.escape(value_slug)}(?=$|-|\d)", slug)
        if match:
            matches.append((match.start(), -len(value_slug), english))

    if matches:
        return min(matches, key=lambda item: (item[0], item[1]))[2]
    return None


def rename_product_images(files, code, mapping):
    """Rename to <code>-main-<color> / <code>-<color>NN; return new paths/count."""
    code_slug = slugify(code) or "product"
    ordered = sorted(files, key=lambda p: natural_key(p.name))
    entries = []
    for path in ordered:
        english = detect_filename_color(mapping, path.stem)
        entries.append({
            "path": path,
            "color": slugify(english or ""),
            "main": is_main_filename(path),
            "no_color": english is None,
        })

    featured = next((e for e in entries if e["main"]), None)
    source_paths = {p.resolve() for p in ordered}
    counters = {}
    taken = set()
    targets = {}
    for entry in entries:
        path = entry["path"]
        color = entry["color"]
        if entry is featured:
            stem = f"{code_slug}-main-{color}" if color else f"{code_slug}-main"
        else:
            group = color or "image"
            counters[group] = counters.get(group, 0) + 1
            stem = f"{code_slug}-{group}{counters[group]:02d}"
        target = path.with_name(stem + path.suffix.lower())
        candidate = target
        n = 2
        while candidate.name.lower() in taken or (
                candidate.exists() and candidate.resolve() not in source_paths):
            candidate = target.with_name(f"{target.stem}-{n}{target.suffix}")
            n += 1
        taken.add(candidate.name.lower())
        targets[path] = candidate
        if entry["no_color"]:
            print(f"WARNING: no known color in image name: {path.name}")

    final = {}
    pending = []
    for i, path in enumerate(ordered):
        target = targets[path]
        if path == target:
            final[path] = path
            continue
        tmp = path.with_name(f".rename-tmp-{i}{path.suffix.lower()}")
        n = 2
        while tmp.exists():
            tmp = path.with_name(f".rename-tmp-{i}-{n}{path.suffix.lower()}")
            n += 1
        try:
            path.rename(tmp)
            pending.append((tmp, target, path))
        except OSError as exc:
            print(f"WARNING: could not rename {path.name!r}: {exc}")
            final[path] = path

    renamed = 0
    for tmp, target, original in pending:
        try:
            tmp.rename(target)
            final[original] = target
            renamed += 1
        except OSError as exc:
            print(f"WARNING: could not rename {tmp.name!r} -> {target.name!r}: {exc}")
            try:
                tmp.rename(original)
            except OSError:
                pass
            final[original] = original
    return [final[p] for p in ordered], renamed


def find_product_images(images_dir, code, mapping):
    folder = Path(images_dir) / code
    if not folder.is_dir():
        return None, [], 0
    files = sorted(
        (p for p in folder.iterdir()
         if p.is_file() and p.suffix.lower() in IMAGE_EXTS),
        key=lambda p: natural_key(p.name),
    )
    files, renamed = rename_product_images(files, code, mapping)
    main = next((p for p in files if is_main_filename(p)), None)
    return main, [p for p in files if p != main], renamed


def image_keywords(mapping, color):
    english = lookup_color(mapping, color)
    if not english:
        return []
    words = re.split(r"[^a-z0-9]+", english.lower())
    return [
        w for w in words
        if len(w) >= 3 and not w.isdigit() and w not in IMAGE_STOPWORDS
    ]


def match_color_image(mapping, color, gallery, main=None):
    """Return (image, matched_by_color). First color-matching gallery image wins."""
    slug = slugify(lookup_color(mapping, color) or "")
    if slug:
        pattern = re.compile(rf"(^|-){re.escape(slug)}(?=$|-|\d)")
        for path in gallery:
            if pattern.search(path.stem.lower()):
                return path, True

    keywords = image_keywords(mapping, color)
    if keywords:
        for path in gallery:
            if any(word in path.stem.lower() for word in keywords):
                return path, True
    if gallery:
        return gallery[0], False
    return main, False


def main(args):
    twb = openpyxl.load_workbook(args.template, read_only=True)
    tws = twb.active
    tmpl_headers = [c.value for c in tws[1]]
    tH = {v: i for i, v in enumerate(tmpl_headers)}
    tmpl_rows = list(tws.iter_rows(min_row=2, values_only=True))
    twb.close()
    slugs = load_slug_map(tmpl_rows, tH)
    color_map = load_color_map(args.color_map)

    out_headers = list(tmpl_headers)
    for extra in ("attribute:material", "attribute_name:material"):
        if extra not in out_headers:
            out_headers.append(extra)
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
                 "دسته بندی", "ابعاد", "جنس", "کاربرد", "رنگ"]:
        if name not in H:
            print(f"WARNING: column missing, treated as empty: {name}")

    wanted = None
    if args.codes:
        wanted = {c.strip().upper() for c in args.codes.split(",") if c.strip()}

    out_rows = []
    stats = collections.Counter()
    warnings = []
    seen_codes = set()
    missing_colors = set()

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
        seen_colors = set()
        for cname, qty in colors:
            ckey = color_key(cname)
            if ckey in seen_colors:
                warnings.append(
                    f"row {excel_row}: {code} duplicate color {cname!r}, kept first")
                continue
            seen_colors.add(ckey)
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

        main_image, gallery_images, renamed_images = find_product_images(
            args.images_dir, code, color_map)
        stats["renamed_images"] += renamed_images
        if main_image or gallery_images:
            stats["products_with_images"] += 1

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
        if main_image:
            parent[oH["images"]] = image_url(main_image, args.image_base)
            parent[oH["image_alt"]] = title
            stats["featured_images"] += 1
        if gallery_images:
            parent[oH["gallery_images"]] = "|".join(
                image_url(p, args.image_base) for p in gallery_images)
        image_files = ([main_image] if main_image else []) + gallery_images
        if image_files:
            parent[oH["image_filename"]] = "|".join(
                image_file_slug(p) for p in image_files)
        out_rows.append(parent)
        stats["products"] += 1

        for cname in color_names:
            if lookup_color(color_map, cname) is None:
                missing_colors.add(cname)
            slug = slugs.get(color_key(cname))
            if not slug:
                slug = slugify(lookup_color(color_map, cname) or "")
            if not slug:
                warnings.append(
                    f"row {excel_row}: {code} unknown color {cname!r}, "
                    f"variation skipped (add it to color_mapping.xlsx)")
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
            var_image, color_matched = match_color_image(
                color_map, cname, gallery_images, main_image)
            if var_image:
                v[oH["images"]] = image_url(var_image, args.image_base)
                v[oH["image_filename"]] = image_file_slug(var_image)
                v[oH["image_titles"]] = f"{title} - {cname}"
                v[oH["image_alt"]] = f"{title} - {cname}"
                stats["variation_images_color" if color_matched
                      else "variation_images_fallback"] += 1
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
    print(f"images: {stats['products_with_images']} products with folders, "
          f"{stats['featured_images']} featured, "
          f"{stats['renamed_images']} files renamed; variations: "
          f"{stats['variation_images_color']} color-matched, "
          f"{stats['variation_images_fallback']} fallback")
    print(f"skipped: no_code={stats['no_code']} invalid={stats['invalid']} "
          f"no_name={stats['no_name']} no_price={stats['no_price']} "
          f"duplicate={stats['duplicate']} unknown_slug={stats['unknown_slug']}")
    if missing_colors:
        print("WARNING: colors missing from color_mapping.xlsx: "
              + ", ".join(sorted(missing_colors)))
    for w in warnings:
        print("WARNING:", w)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--codes", default="",
                    help="only process these product codes, comma-separated (for testing)")
    ap.add_argument("--src", default=SRC)
    ap.add_argument("--template", default=TEMPLATE)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--color-map", default=COLOR_MAP_FILE,
                    help="Persian->English color mapping workbook")
    ap.add_argument("--images-dir", default=IMAGES_DIR,
                    help="folder containing one subfolder per product code")
    ap.add_argument("--image-base", default=IMAGE_URL_BASE,
                    help="public URL prefix for image files")
    main(ap.parse_args())
