"""Standardize `دسته بندی` values against `standard_categories.xlsx`.

Shared by fill.py (`--categories-only` writes the source workbook in place) and
build_woocommerce.py (standardizes in memory while building the import sheet),
so the database and the generated sheet always use identical matching rules.

A category cell may hold several parts separated by `|`, `-`, `–`, `—`, `،`
or `,`; each part is stripped of the `کیف` prefix / `زنانه` suffix, run through
the alias map, then matched against the standard subcategory keywords.

    "کیف دستی | کیف مجسلی"  ->  "کیف دستی زنانه | کیف مجلسی زنانه"
"""

import re

import openpyxl

CATEGORY_SEP = re.compile(r"[|\-–—،,]")


def strip_affixes(s):
    """`کیف دوشی زنانه` -> `دوشی`"""
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
    wb.close()
    return mapping


def standardize_category(raw, keyword_map, aliases):
    """`کیف دستی | کیف دوشی` -> (`کیف دستی زنانه | کیف دوشی زنانه`, unmatched).

    Returns (None, unmatched) when no part matches a standard subcategory.
    """
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
