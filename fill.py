import argparse
import json
import re
import sys
from pathlib import Path

import openpyxl

from tgcommon import extract_codes, extract_fields, normalize

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE_DIR = Path(__file__).resolve().parent
POSTS_FILE = BASE_DIR / "cache/posts.json"
SETTINGS = BASE_DIR / "settings.json"
_SETTINGS = json.loads(SETTINGS.read_text(encoding="utf-8"))
SRC = BASE_DIR / _SETTINGS["files"]["source"]
STD_CATS = BASE_DIR / _SETTINGS["files"]["standard_categories"]
CATEGORY_ALIASES = _SETTINGS.get("category_aliases", {})
REPORT = BASE_DIR / "report.txt"

FIELDS = ["ابعاد", "جنس", "نوع بند", "کاربرد"]

CATEGORY_SEP = re.compile(r"[|\-–—،,]")


def strip_affixes(s):
    """کیف دوشی زنانه -> دوشی"""
    s = str(s).strip()
    if s.startswith("کیف "):
        s = s[4:].strip()
    elif s == "کیف":
        s = ""
    for suf in (" زنانه", "زنانه"):
        if s.endswith(suf):
            s = s[:-len(suf)].strip()
            break
    return re.sub(r"\s+", " ", s).strip()


def load_standard_categories(path):
    """Read standard_categories.xlsx -> {keyword: standard_subcategory}."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    mapping = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        sub = row[1] if len(row) > 1 else None
        if sub is None or str(sub).strip() in ("", "—", "-"):
            continue
        sub = str(sub).strip()
        kw = strip_affixes(sub)
        if kw:
            mapping[kw] = sub
    return mapping


def standardize_category(raw, keyword_map, aliases):
    """کیف دستی | کیف دوشی -> کیف دستی زنانه | کیف دوشی زنانه"""
    parts = [p.strip() for p in CATEGORY_SEP.split(str(raw)) if p.strip()]
    matched, unmatched = [], []
    for p in parts:
        kw = strip_affixes(p)
        kw = aliases.get(kw, kw)
        if kw in keyword_map:
            matched.append(keyword_map[kw])
        else:
            unmatched.append(p)
    seen, out = set(), []
    for m in matched:
        if m not in seen:
            seen.add(m)
            out.append(m)
    return (" | ".join(out) if out else None), unmatched


def standardize_column(ws, header, std_map):
    """Standardize the دسته بندی column in-place; return category stats."""
    stats = {"category": [], "category_cleared": [], "category_empty": 0}
    for row in range(2, ws.max_row + 1):
        cat_cell = ws.cell(row, header["دسته بندی"])
        raw = cat_cell.value
        if raw in (None, ""):
            stats["category_empty"] += 1
            continue
        standardized, unmatched = standardize_category(raw, std_map, CATEGORY_ALIASES)
        if standardized is None:
            cat_cell.value = None
            stats["category_cleared"].append((row, raw, unmatched))
        else:
            cat_cell.value = standardized
            stats["category"].append((row, raw, standardized, unmatched))
    return stats


def categories_only():
    """Only standardize the دسته بندی column (no field filling)."""
    wb = openpyxl.load_workbook(SRC)
    ws = wb.worksheets[0]
    header = {c.value: c.column for c in ws[1] if c.value}
    if "دسته بندی" not in header:
        raise SystemExit("missing column: دسته بندی")
    std_map = load_standard_categories(STD_CATS)
    stats = standardize_column(ws, header, std_map)
    wb.save(SRC)
    print(f"saved {SRC}")

    lines = []
    lines.append(f"== دسته بندی standardized ({len(stats['category'])}) ==")
    for r, raw, std, un in stats["category"]:
        warn = f"  [unmatched: {', '.join(un)}]" if un else ""
        lines.append(f"  row {r}: {raw}  ->  {std}{warn}")
    lines.append(f"== دسته بندی cleared ({len(stats['category_cleared'])}) ==")
    for r, raw, un in stats["category_cleared"]:
        lines.append(f"  row {r}: {raw}  ->  (empty)  [unmatched: {', '.join(un)}]")
    lines.append(f"rows with empty category: {stats['category_empty']}")

    Path(REPORT).write_text("\n".join(lines), encoding="utf-8")
    print(f"report -> {REPORT}")
    print(f"category_standardized={len(stats['category'])} "
          f"category_cleared={len(stats['category_cleared'])}")


def build_index(posts):
    exact = {}
    nums = {}
    for p in posts:
        fields = extract_fields(p["text"])
        ek, nk = extract_codes(p["text"])
        for key, mtype in ek.items():
            exact.setdefault(key, []).append((mtype, fields, p))
        for key, mtype in nk.items():
            nums.setdefault(key, []).append((mtype, fields, p))
    return exact, nums


def build_seller_affinity(rows, nums):
    """How often a seller's product codes show up in each channel."""
    counts = {}
    for _row, seller, num in rows:
        if not seller:
            continue
        for mtype, _fields, p in nums.get(num, ()):
            w = 3 if mtype == "code_kw" else 1
            counts[(seller, p["channel"])] = counts.get((seller, p["channel"]), 0) + w
    return counts


def score(mtype, fields):
    base = {"code_kw": 60, "full": 40}.get(mtype, 10)
    return base + 2 * len(fields)


def best_candidate(cands, affinity):
    def key(c):
        mtype, fields, post = c
        return (
            affinity.get(post["channel"], 0),
            score(mtype, fields),
            len(fields),
            len(post["text"]),
        )

    ranked = sorted(cands, key=key, reverse=True)
    top = key(ranked[0])
    ties = [c for c in ranked if key(c) == top]
    return ranked[0], len(ties)


def main(args):
    if args.categories_only:
        categories_only()
        return

    posts = json.loads(POSTS_FILE.read_text(encoding="utf-8"))
    print(f"loaded {len(posts)} posts")
    exact, nums = build_index(posts)

    wb = openpyxl.load_workbook(SRC)
    ws = wb.worksheets[0]
    header = {c.value: c.column for c in ws[1] if c.value}
    for name in FIELDS + ["کد محصول", "فروشنده", "دسته بندی"]:
        if name not in header:
            raise SystemExit(f"missing column: {name}")

    parsed = []
    for row in range(2, ws.max_row + 1):
        raw = ws.cell(row, header["کد محصول"]).value
        seller = ws.cell(row, header["فروشنده"]).value
        code = normalize(str(raw)).strip().upper() if raw not in (None, "") else ""
        m = re.match(r"^([A-Z]+)-(\d+)$", code)
        parsed.append((row, code, m, seller))

    if args.codes:
        wanted = {c.strip().upper() for c in args.codes.split(",") if c.strip()}
        parsed = [p for p in parsed if p[1] in wanted]
        print(f"testing with codes: {sorted(wanted)}")

    code_rows = [(r, s, m.group(2)) for r, c, m, s in parsed if m]
    seller_affinity = build_seller_affinity(code_rows, nums)

    stats = {"filled_rows": 0, "no_code": 0, "invalid": [], "not_found": [],
             "found_empty": [], "kept_existing": [], "ambiguous": [], "filled": [],
             "category": [], "category_cleared": [], "category_empty": 0}

    for row, code, m, seller in parsed:
        if not code:
            stats["no_code"] += 1
            continue
        if not m:
            stats["invalid"].append((row, code))
            continue

        cands = exact.get(code) or nums.get(m.group(2))
        if not cands:
            hint = [
                f"{p['channel']} #{p['id']}"
                for p in posts
                if re.search(rf"(?<!\d){re.escape(m.group(2))}(?!\d)", normalize(p["text"]))
            ][:3]
            stats["not_found"].append((row, code, hint))
            continue

        affinity = {}
        if seller:
            for (sel, chan), w in seller_affinity.items():
                if sel == seller:
                    affinity[chan] = w

        (mtype, fields, post), ties = best_candidate(cands, affinity)
        if ties > 1:
            stats["ambiguous"].append((row, code, ties))

        if not fields:
            stats["found_empty"].append(
                (row, code, post["channel"], post["id"], post["text"][:200])
            )
            continue

        written = {}
        for name in FIELDS:
            if name not in fields:
                continue
            cell = ws.cell(row, header[name])
            if cell.value not in (None, "") and not args.refresh:
                stats["kept_existing"].append((row, code, name, cell.value))
                continue
            cell.value = fields[name]
            written[name] = fields[name]
        if written:
            stats["filled_rows"] += 1
            stats["filled"].append((row, code, post["channel"], post["id"], written))

    std_map = load_standard_categories(STD_CATS)
    stats.update(standardize_column(ws, header, std_map))

    wb.save(SRC)
    print(f"saved {SRC}")

    lines = []
    lines.append(f"posts: {len(posts)}")
    lines.append(f"rows filled: {stats['filled_rows']}")
    lines.append(f"rows without code: {stats['no_code']}")
    lines.append("")
    lines.append("== seller -> channel used ==")
    for (sel, chan), w in sorted(seller_affinity.items(), key=lambda x: -x[1]):
        lines.append(f"  {sel} -> {chan} ({w})")
    lines.append("")
    lines.append(f"== filled ({len(stats['filled'])}) ==")
    for r, c, ch, mid, f in stats["filled"]:
        vals = "; ".join(f"{k}={v}" for k, v in f.items())
        lines.append(f"  row {r}: {c}  [{ch} #{mid}]  {vals}")
    lines.append(f"== دسته بندی standardized ({len(stats['category'])}) ==")
    for r, raw, std, un in stats["category"]:
        warn = f"  [unmatched: {', '.join(un)}]" if un else ""
        lines.append(f"  row {r}: {raw}  ->  {std}{warn}")
    lines.append(f"== دسته بندی cleared ({len(stats['category_cleared'])}) ==")
    for r, raw, un in stats["category_cleared"]:
        lines.append(f"  row {r}: {raw}  ->  (empty)  [unmatched: {', '.join(un)}]")
    lines.append(f"rows with empty category: {stats['category_empty']}")
    lines.append(f"== invalid code ({len(stats['invalid'])}) ==")
    for r, c in stats["invalid"]:
        lines.append(f"  row {r}: {c}")
    lines.append(f"== code not found in channels ({len(stats['not_found'])}) ==")
    for r, c, hint in stats["not_found"]:
        extra = f"  (digits seen in: {', '.join(hint)})" if hint else ""
        lines.append(f"  row {r}: {c}{extra}")
    lines.append(f"== found post but no field extracted ({len(stats['found_empty'])}) ==")
    for r, c, ch, mid, txt in stats["found_empty"]:
        lines.append(f"  row {r}: {c}  [{ch} #{mid}] {txt}")
    lines.append(f"== ambiguous candidates ({len(stats['ambiguous'])}) ==")
    for r, c, n in stats["ambiguous"]:
        lines.append(f"  row {r}: {c} ({n} tied)")
    lines.append(f"== existing values kept ({len(stats['kept_existing'])}) ==")
    for r, c, n, v in stats["kept_existing"]:
        lines.append(f"  row {r}: {c} {n}={v}")

    Path(REPORT).write_text("\n".join(lines), encoding="utf-8")
    print(f"report -> {REPORT}")
    print(f"filled={stats['filled_rows']} not_found={len(stats['not_found'])} "
          f"empty_post={len(stats['found_empty'])} invalid={len(stats['invalid'])} "
          f"category_standardized={len(stats['category'])} "
          f"category_cleared={len(stats['category_cleared'])}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--refresh", action="store_true",
                    help="overwrite already-filled cells with fresh values")
    ap.add_argument("--categories-only", action="store_true",
                    help="only standardize the دسته بندی column (skip field filling)")
    ap.add_argument("--codes", default="",
                    help="only process these product codes, comma-separated (for testing)")
    args = ap.parse_args()
    main(args)
