import re

_DIGIT_MAP = str.maketrans(
    "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
    "01234567890123456789",
)

_PERSIAN_LETTER = r"[\u0600-\u06FF]"
_JUNK = r"[^\w\n]{0,15}?"  # spaces, emoji, punctuation (no letters/digits/newline)

CODE_AFTER = re.compile(
    r"(?:کد(?:\s*(?:محصول|کالا|کلیدی|شناسه))?|code)" + _JUNK +
    r"([A-Za-z]{0,6}[-]?\d{2,7})",
    re.I,
)
FULL_CODE = re.compile(r"\b([A-Za-z]{2,6})\s*[-–]\s*(\d{2,7})\b")

_SEG = r"\d{1,4}(?:/\d)?"
DIM = _SEG + r"\s*[*×xX]\s*" + _SEG + r"(?:\s*[*×xX]\s*" + _SEG + r")?"
DIM_SLASH = r"\d{1,4}/\d{1,2}"
DIM_AFTER_WORD = re.compile(
    r"(?:ابعاد|سایز)" + _JUNK + r"(" + DIM + r"|" + DIM_SLASH + r")", re.I
)
DIM_ANY = re.compile(DIM)

GENS_AFTER = re.compile(r"جنس" + _JUNK + r"([^\n،,;|]+)")
BAND_AFTER = re.compile(r"نوع\s*بند" + _JUNK + r"([^\n،,;|]+)")
BAND_PHRASE = re.compile(
    r"(?<!نوع )(?<!\d)(?<!" + _PERSIAN_LETTER + r")بند(?!ی)" + _JUNK + r"[^\n،,;|]*"
)
CAT_AFTER = re.compile(r"دسته\s*بندی" + _JUNK + r"([^\n،,;|]+)")

_BAND_STOP = [" هم داره", "داره", " میشه", "بصورت", "ازش", " استفاده", " که", " کرد"]

CAT_KEYWORDS = ["دوشی", "روزمره", "اداری", "مجلسی", "کوله", "پاسپورتی"]

_NEXT_KEYWORDS = [
    "نوع بند",
    "دسته بندی",
    "ابعاد",
    "سایز",
    "کد",
    "رنگ",
    "قیمت",
    "تعداد",
    "مدل",
]


def normalize(text):
    if not text:
        return ""
    return (
        text.translate(_DIGIT_MAP)
        .replace("ك", "ک")
        .replace("ي", "ی")
        .replace("ى", "ی")
        .replace("–", "-")
        .replace("—", "-")
    )


def _cut(value):
    cut = len(value)
    for i, ch in enumerate(value):
        if ch in "\n\r،,;|":
            cut = min(cut, i)
    for kw in _NEXT_KEYWORDS:
        i = value.find(kw)
        if i > 0:
            cut = min(cut, i)
    return value[:cut].strip().strip(":：-").strip()


def _band_ok(value, text, start):
    if not value or value.startswith("بندی"):
        return False
    if any(ch.isdigit() for ch in value):
        return False
    line_start = text.rfind("\n", 0, start) + 1
    if "قیمت" in text[line_start:start]:
        return False
    return True


def _trim_band(value):
    cut = len(value)
    for stop in _BAND_STOP:
        i = value.find(stop)
        if i > 0:
            cut = min(cut, i)
    value = value[:cut].strip().strip(":：-،,").strip()
    while value.endswith(")") and value.count(")") > value.count("("):
        value = value[:-1].strip()
    while value.endswith("(") and value.count("(") > value.count(")"):
        value = value[:-1].strip()
    words = value.split()
    if len(words) > 8:
        value = " ".join(words[:8])
    return value


def extract_fields(text):
    text = normalize(text)
    out = {}

    m = DIM_AFTER_WORD.search(text) or DIM_ANY.search(text)
    if m:
        val = m.group(1) if m.lastindex else m.group(0)
        out["ابعاد"] = val.replace(" ", "")

    m = GENS_AFTER.search(text)
    if m:
        val = _cut(m.group(1))
        if val and val != "جنس":
            out["جنس"] = val

    m = BAND_AFTER.search(text)
    val = m.group(1).strip() if m else ""
    if val:
        val = _trim_band(val)
    if not val:
        m = BAND_PHRASE.search(text)
        if m and _band_ok(m.group(0), text, m.start()):
            val = _trim_band(m.group(0))
    if val and val != "بند":
        out["نوع بند"] = val

    matched = [k for k in CAT_KEYWORDS if k in text]
    if matched:
        out["کاربرد"] = " | ".join(dict.fromkeys(matched))
    else:
        m = CAT_AFTER.search(text)
        if m:
            val = _cut(m.group(1))
            if val and "دسته بندی" not in val:
                out["کاربرد"] = val

    return out


def extract_codes(text):
    """Return (exact_keys, num_keys).

    exact_keys: {"FA-1959": "code_kw"|"full"}  prefix + number found in text
    num_keys:   {"1959": "code_kw"|"full"}     bare number, higher trust when
                                               written next to a کد keyword
    """
    text = normalize(text)
    exact = {}
    nums = {}

    for m in CODE_AFTER.finditer(text):
        frag = m.group(1).replace(" ", "")
        fm = re.match(r"^([A-Za-z]{1,6})[-]?(\d{2,7})$", frag)
        if fm:
            exact[f"{fm.group(1).upper()}-{fm.group(2)}"] = "code_kw"
            nums[fm.group(2)] = "code_kw"
        elif frag.isdigit():
            nums[frag] = "code_kw"

    for m in FULL_CODE.finditer(text):
        key = f"{m.group(1).upper()}-{m.group(2)}"
        exact.setdefault(key, "full")
        nums.setdefault(m.group(2), "full")

    return exact, nums
