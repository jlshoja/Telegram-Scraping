"""Build a WooCommerce import sheet from the product database workbook.

Every fixed parameter (file paths, sheet names, source/output column headers,
the full output header row, attribute labels, category prefix, title/SKU
formats, image URL base, color aliases, ...) lives in settings.json. Edit that
file instead of this code when something changes; this script only holds logic.

settings.json sections:
  files            source/output/color-map paths and sheet names
  color_map        color-map columns/sheet/start row
  images           url_base ({year}/{month}/{day} = current date), extensions,
                  main-name regex, stopwords
  colors           aliases (Persian->Persian) and extra_slugs (Persian->slug)
  source_columns   logical name -> source workbook header
  output_columns   logical name -> WooCommerce header
  output_headers   the exact ordered WooCommerce header row (39 columns)
  values           attribute labels, category prefix, formats, regexes, ...
  options          require_images: skip products that have no picture

Only products with a name, a price and pictures end up in the output file
(options.require_images = false keeps products without pictures).

Usage:
  python build_woocommerce.py
  python build_woocommerce.py --settings other.json
  python build_woocommerce.py --codes FA-1959,TR-402   # test subset
  python build_woocommerce.py --out woocommerce_import_output.xlsx
"""

import argparse
import collections
import json
import re
import sys
from datetime import date
from pathlib import Path
from urllib.parse import quote

import openpyxl

from tgcommon import normalize

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

DEFAULT_SETTINGS_FILE = "settings.json"
BASE_DIR = Path(__file__).resolve().parent

# --- settings-driven state, populated by configure() -------------------------
SRC = ""
SOURCE_SHEET = None
OUT = ""
COLOR_MAP_FILE = ""
COLOR_MAP_OPTS = {}
IMAGES_DIR = ""
IMAGE_URL_BASE = ""
IMAGE_EXTS = frozenset()
IMAGE_STOPWORDS = frozenset()
MAIN_NAME_RE = re.compile(r"$^")
COLOR_ALIASES = {}
EXTRA_SLUGS = {}
SALE_TAG = ""
CODE_RE = re.compile(r"$^")
MISSING_NAMES = frozenset()
SRC_COLS = {}
OPTIONAL_SRC_COLS = frozenset()
OUT_COLS = {}
OUT_HEADERS = []
OPTIONS = {}
VALUES = {}


def load_settings(path):
    p = Path(path)
    if not p.exists():
        raise SystemExit(f"settings file not found: {p}")
    with p.open(encoding="utf-8") as fh:
        return json.load(fh)


def resolve_url_base(template, fixed_date=None):
    """Fill {year}/{month}/{day} in an uploads URL prefix.

    fixed_date: "YYYY-MM" or "YYYY-MM-DD" to pin the date, None = today.
    """
    if fixed_date:
        parts = [int(p) for p in str(fixed_date).split("-") if p]
        year = parts[0] if parts else date.today().year
        month = parts[1] if len(parts) > 1 else 1
        day = parts[2] if len(parts) > 2 else 1
    else:
        today = date.today()
        year, month, day = today.year, today.month, today.day
    return (str(template)
            .replace("{year}", f"{year:04d}")
            .replace("{month}", f"{month:02d}")
            .replace("{day}", f"{day:02d}"))


def configure(cfg):
    """Copy the settings file into module state (single source of truth)."""
    global SRC, SOURCE_SHEET, OUT
    global COLOR_MAP_FILE, COLOR_MAP_OPTS, IMAGES_DIR, IMAGE_URL_BASE
    global IMAGE_EXTS, IMAGE_STOPWORDS, MAIN_NAME_RE, COLOR_ALIASES, EXTRA_SLUGS
    global SALE_TAG, CODE_RE, MISSING_NAMES, SRC_COLS, OUT_COLS, OUT_HEADERS
    global OPTIONAL_SRC_COLS, VALUES, OPTIONS

    files = cfg["files"]
    SRC = files["source"]
    SOURCE_SHEET = files.get("source_sheet")
    OUT = files["output"]
    COLOR_MAP_FILE = files["color_map"]
    IMAGES_DIR = files["images_dir"]

    cm = cfg["color_map"]
    COLOR_MAP_OPTS = {
        "sheet": cm.get("sheet"),
        "persian_col": cm.get("persian_column", 1),
        "english_col": cm.get("english_column", 2),
        "first_row": cm.get("first_data_row", 2),
    }

    images = cfg["images"]
    IMAGE_URL_BASE = resolve_url_base(
        images["url_base"], images.get("url_base_date"))
    IMAGE_EXTS = frozenset(str(e).lower() for e in images["extensions"])
    IMAGE_STOPWORDS = frozenset(str(w).lower() for w in images["stopwords"])
    MAIN_NAME_RE = re.compile(images["main_name_pattern"], re.IGNORECASE)

    colors = cfg["colors"]
    COLOR_ALIASES = {color_key(k): v for k, v in colors["aliases"].items()}
    EXTRA_SLUGS = dict(colors["extra_slugs"])

    SRC_COLS = dict(cfg["source_columns"])
    OPTIONAL_SRC_COLS = frozenset(cfg.get("optional_source_columns", []))
    OUT_COLS = dict(cfg["output_columns"])
    OUT_HEADERS = list(cfg["output_headers"])

    VALUES = dict(cfg["values"])
    OPTIONS = dict(cfg.get("options", {}))
    SALE_TAG = VALUES["sale_tag"]
    CODE_RE = re.compile(VALUES["code_pattern"])
    MISSING_NAMES = frozenset(VALUES["missing_names"])


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


def load_color_map(path, sheet=None, persian_col=1, english_col=2, first_row=2):
    """Load Persian -> English colors from color_mapping.xlsx (fresh each run)."""
    p = Path(path)
    if not p.exists():
        print(f"WARNING: color mapping file not found: {p}")
        return {}
    wb = openpyxl.load_workbook(p, read_only=True)
    ws = wb[sheet] if sheet else wb.active
    pc, ec = int(persian_col) - 1, int(english_col) - 1
    mapping = {}
    for r in ws.iter_rows(min_row=int(first_row), values_only=True):
        if not r or len(r) <= max(pc, ec):
            continue
        fa, en = r[pc], r[ec]
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


def natural_key(value):
    parts = re.split(r"(\d+)", normalize(str(value)).lower())
    return [(0, int(p)) if p.isdigit() else (1, p) for p in parts]


def is_main_filename(path):
    return bool(MAIN_NAME_RE.search(path.stem))


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
        match = re.search(rf"(^|-){re.escape(value_slug)}(?=$|-|\d)", slug)
        if match:
            matches.append((match.start(), -len(value_slug), english))

    if matches:
        return min(matches, key=lambda item: (item[0], item[1]))[2]
    return None


def image_url(path, base):
    return base.rstrip("/") + "/" + quote(path.name)


def find_product_images(images_dir, code):
    folder = Path(images_dir) / code
    if not folder.is_dir():
        return None, []
    files = sorted(
        (p for p in folder.iterdir()
         if p.is_file() and p.suffix.lower() in IMAGE_EXTS),
        key=lambda p: natural_key(p.name),
    )
    main = next((p for p in files if is_main_filename(p)), None)
    return main, [p for p in files if p != main]


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
    wanted = slugify(lookup_color(mapping, color) or "")
    if wanted:
        for path in gallery:
            found = detect_filename_color(mapping, path.stem)
            if found and slugify(found) == wanted:
                return path, True

    keywords = image_keywords(mapping, color)
    if keywords:
        for path in gallery:
            if any(word in path.stem.lower() for word in keywords):
                return path, True
    if gallery:
        return gallery[0], False
    return main, False


def project_path(value):
    """Resolve a path: absolute stays as is, relative to the script folder."""
    p = Path(value)
    return p if p.is_absolute() else (BASE_DIR / p)


def main(args):
    configure(load_settings(project_path(args.settings)))

    src_path = project_path(args.src or SRC)
    out_path = project_path(args.out or OUT)
    color_map_path = project_path(args.color_map or COLOR_MAP_FILE)
    images_dir = project_path(args.images_dir or IMAGES_DIR)
    image_base = args.image_base or IMAGE_URL_BASE

    color_map = load_color_map(color_map_path, **COLOR_MAP_OPTS)
    slugs = {color_key(k): slugify(v)
             for k, v in EXTRA_SLUGS.items() if v}

    out_headers = list(OUT_HEADERS)
    index_of = {v: i for i, v in enumerate(out_headers)}
    oH = {}
    for logical, header in OUT_COLS.items():
        if header not in index_of:
            raise SystemExit(
                f"output_headers in settings is missing: {header}")
        oH[logical] = index_of[header]

    swb = openpyxl.load_workbook(src_path, read_only=True)
    sws = swb[SOURCE_SHEET] if SOURCE_SHEET else swb.active
    src_headers = [c.value for c in sws[1]]
    H = col_index(src_headers)
    rows = list(sws.iter_rows(min_row=2, values_only=True))
    swb.close()

    if SRC_COLS["code"] not in H:
        raise SystemExit(f"missing column: {SRC_COLS['code']}")
    for logical, header in SRC_COLS.items():
        if logical != "code" and logical not in OPTIONAL_SRC_COLS \
                and header not in H:
            print(f"WARNING: column missing, treated as empty: {header}")

    wanted = None
    if args.codes:
        wanted = {c.strip().upper() for c in args.codes.split(",") if c.strip()}

    out_rows = []
    stats = collections.Counter()
    warnings = []
    seen_codes = set()
    missing_colors = set()

    for excel_row, r in enumerate(rows, 2):
        raw_code = cell(r, H, SRC_COLS["code"])
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

        name = cell(r, H, SRC_COLS["name"])
        if name in (None, "") or str(name).strip() in MISSING_NAMES:
            name = cell(r, H, SRC_COLS["alt_name"])
        if name in (None, "") or str(name).strip() in MISSING_NAMES:
            warnings.append(f"row {excel_row}: {code} has no product name, skipped")
            stats["no_name"] += 1
            continue
        title = VALUES["title_format"].format(name=str(name).strip(), code=code)

        price = as_number(cell(r, H, SRC_COLS["price"]))
        if price is None:
            stats["no_price"] += 1
            continue
        sale = as_number(cell(r, H, SRC_COLS["sale_price"]))
        tag = SALE_TAG if sale is not None else None

        cats = cell(r, H, SRC_COLS["category"])
        if cats in (None, ""):
            categories = None
        else:
            parts = [p.strip() for p in str(cats).split("|") if p.strip()]
            required_prefix = VALUES["category_required_prefix"]
            for p in parts:
                if not p.startswith(required_prefix):
                    warnings.append(
                        f"row {excel_row}: {code} unusual category value: {p!r}")
            categories = " | ".join(
                f"{VALUES['category_parent']}>{p}" for p in parts) or None

        # --- attributes (only non-empty ones are written) ---
        colors = parse_colors(cell(r, H, SRC_COLS["color"]))
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
            warnings.append(f"row {excel_row}: {code} has رنگ but no usable colors")

        usage = cell(r, H, SRC_COLS["usage"])
        if usage in (None, ""):
            usage = None
        else:
            usage = "|".join(
                p.strip() for p in str(usage).split("|") if p.strip()) or None
        material = cell(r, H, SRC_COLS["material"])
        dimension = cell(r, H, SRC_COLS["dimension"])
        if isinstance(dimension, (int, float)):
            dimension = str(dimension)

        main_image, gallery_images = find_product_images(images_dir, code)
        if not main_image and not gallery_images:
            stats["no_images"] += 1
            if OPTIONS.get("require_images", True):
                continue
        else:
            stats["products_with_images"] += 1

        def new_row():
            return [None] * len(out_headers)

        parent = new_row()
        parent[oH["post_type"]] = VALUES["product_type"]
        parent[oH["post_title"]] = title
        parent[oH["post_status"]] = VALUES["post_status"]
        parent[oH["sku"]] = code
        parent[oH["regular_price"]] = price
        if sale is not None:
            parent[oH["sale_price"]] = sale
            parent[oH["sale_tag"]] = tag
        parent[oH["manage_stock"]] = VALUES["manage_stock"]
        if categories is not None:
            parent[oH["categories"]] = categories
        if color_names:
            parent[oH["attribute_color"]] = "|".join(color_names)
            parent[oH["attribute_name_color"]] = VALUES["color_attribute_name"]
        if usage is not None:
            parent[oH["attribute_application"]] = usage
            parent[oH["attribute_name_application"]] = \
                VALUES["application_attribute_name"]
        if material is not None:
            parent[oH["attribute_material"]] = material
            parent[oH["attribute_name_material"]] = VALUES["material_attribute_name"]
        if dimension is not None:
            parent[oH["attribute_dimention"]] = dimension
            parent[oH["attribute_name_dimention"]] = \
                VALUES["dimention_attribute_name"]
        if main_image:
            parent[oH["images"]] = image_url(main_image, image_base)
            parent[oH["image_alt"]] = title
            stats["featured_images"] += 1
        if gallery_images:
            parent[oH["gallery_images"]] = "|".join(
                image_url(p, image_base) for p in gallery_images)
        image_files = ([main_image] if main_image else []) + gallery_images
        if image_files:
            parent[oH["image_filename"]] = "|".join(p.name for p in image_files)
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
                    f"variation skipped (add it to {color_map_path})")
                stats["unknown_slug"] += 1
                continue
            qty = stocks[cname]
            if qty is None:
                warnings.append(
                    f"row {excel_row}: {code} color {cname!r} has no stock number, "
                    f"used {VALUES['out_of_stock']!r}")
                qty = VALUES["out_of_stock"]
            v = new_row()
            v[oH["post_type"]] = VALUES["variation_type"]
            v[oH["post_status"]] = VALUES["post_status"]
            v[oH["sku"]] = VALUES["variation_sku_format"].format(
                code=code, slug=slug)
            v[oH["parent_sku"]] = code
            v[oH["regular_price"]] = price
            if sale is not None:
                v[oH["sale_price"]] = sale
                v[oH["sale_tag"]] = tag
            v[oH["manage_stock"]] = VALUES["manage_stock"]
            v[oH["stock_quantity"]] = qty
            v[oH["attribute_color"]] = cname
            v[oH["attribute_name_color"]] = VALUES["color_attribute_name"]
            var_image, color_matched = match_color_image(
                color_map, cname, gallery_images, main_image)
            if var_image:
                v[oH["images"]] = image_url(var_image, image_base)
                v[oH["image_filename"]] = var_image.name
                var_title = VALUES["variation_image_title_format"].format(
                    title=title, color=cname)
                v[oH["image_titles"]] = var_title
                v[oH["image_alt"]] = var_title
                stats["variation_images_color" if color_matched
                      else "variation_images_fallback"] += 1
            out_rows.append(v)
            stats["variations"] += 1

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = VALUES["sheet_name"]
    ws.append(out_headers)
    for row in out_rows:
        ws.append(row)
    wb.save(out_path)

    print(f"saved {out_path}: {stats['products']} products, "
          f"{stats['variations']} variations")
    print(f"images: {stats['products_with_images']} products with folders, "
          f"{stats['featured_images']} featured; variations: "
          f"{stats['variation_images_color']} color-matched, "
          f"{stats['variation_images_fallback']} fallback")
    print(f"image url base: {image_base}")
    print(f"skipped: no_code={stats['no_code']} invalid={stats['invalid']} "
          f"no_name={stats['no_name']} no_price={stats['no_price']} "
          f"no_images={stats['no_images']} "
          f"duplicate={stats['duplicate']} unknown_slug={stats['unknown_slug']}")
    if missing_colors:
        print(f"WARNING: colors missing from {color_map_path}: "
              + ", ".join(sorted(missing_colors)))
    for w in warnings:
        print("WARNING:", w)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--settings", default=DEFAULT_SETTINGS_FILE,
                    help="JSON file holding every fixed parameter")
    ap.add_argument("--codes", default="",
                    help="only process these product codes, comma-separated (for testing)")
    ap.add_argument("--src", help="override settings.files.source")
    ap.add_argument("--out", help="override settings.files.output")
    ap.add_argument("--color-map", help="override settings.files.color_map")
    ap.add_argument("--images-dir", help="override settings.files.images_dir")
    ap.add_argument("--image-base", help="override settings.images.url_base")
    main(ap.parse_args())
