import argparse
import asyncio
import json
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl
from PIL import Image
from telethon import TelegramClient, errors

from scrape import API_HASH, API_ID, SESSION, link_to_target, parse_proxy, read_channels, resolve
from tgcommon import extract_codes, extract_fields, normalize

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

POSTS_FILE = Path("cache/posts.json")
SRC = "Product Database.xlsx"
DST = "Product Database - filled.xlsx"
REPORT = "report.txt"

FIELDS = ["ابعاد", "جنس", "نوع بند", "کاربرد"]
COL = {  # 1-based excel columns
    "ابعاد": 11,
    "جنس": 12,
    "نوع بند": 13,
    "کاربرد": 14,
    "دسته بندی": 6,
}

MEDIA_DIR = Path("media")
MAX_WEBP_SIZE_KB = 100


def usage_to_category(usage):
    """دوشی | پاسپورتی  ->  کیف دوشی زنانه | کیف پاسپورتی زنانه"""
    parts = [p.strip() for p in str(usage).split("|") if p.strip()]
    out = []
    for p in parts:
        out.append(p if p.startswith("کیف") else f"کیف {p} زنانه")
    return " | ".join(dict.fromkeys(out))


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


def convert_to_webp(src, dst, max_kb=MAX_WEBP_SIZE_KB):
    """Convert image to WebP under max_kb. Returns (ok, size_kb)."""
    img = Image.open(src)
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGB")
    for q in (85, 75, 65, 55, 45):
        img.save(dst, "WEBP", quality=q, method=6)
        size_kb = dst.stat().st_size / 1024
        if size_kb <= max_kb:
            return True, size_kb
    w, h = img.size
    while w > 200 and h > 200:
        w, h = int(w * 0.8), int(h * 0.8)
        small = img.resize((w, h), Image.LANCZOS)
        small.save(dst, "WEBP", quality=60, method=6)
        size_kb = dst.stat().st_size / 1024
        if size_kb <= max_kb:
            return True, size_kb
    return False, dst.stat().st_size / 1024


def media_kind(msg):
    if getattr(msg, "photo", None):
        return "photo"
    doc = getattr(msg, "document", None)
    if doc is not None:
        mime = getattr(doc, "mime_type", "") or ""
        if mime.startswith("video/"):
            return "video"
        if mime.startswith("image/"):
            return "photo"
    if getattr(msg, "video", None) or getattr(msg, "gif", None):
        return "video"
    return None


def video_ext(msg):
    doc = getattr(msg, "document", None)
    if doc is not None:
        for a in getattr(doc, "attributes", []):
            name = getattr(a, "file_name", None)
            if name and "." in name:
                return Path(name).suffix.lower()
    return ".mp4"


async def fetch_post_messages(client, entity, msg_id, timeout=60):
    """Fetch a message plus its album mates (same grouped_id)."""
    try:
        msg = await asyncio.wait_for(client.get_messages(entity, ids=msg_id), timeout)
    except asyncio.TimeoutError:
        return [], "fetch timed out"
    except Exception as exc:
        return [], f"fetch failed: {type(exc).__name__} {exc}"
    if msg is None:
        return [], "message not found (deleted?)"
    msgs = [msg]
    gid = getattr(msg, "grouped_id", None)
    if gid:
        try:
            batch = await asyncio.wait_for(
                client.get_messages(entity, ids=list(range(msg_id - 10, msg_id + 11))),
                timeout,
            )
            if batch:
                for m in batch:
                    if m is not None and m.id != msg_id and getattr(m, "grouped_id", None) == gid:
                        msgs.append(m)
                msgs.sort(key=lambda m: m.id)
        except Exception:
            pass
    return msgs, ""


async def download_all(jobs):
    """jobs: list of (code, channel_title, msg_id). Returns dict code -> result."""
    from scrape import PROXY_URL

    results = {}
    proxy = parse_proxy(PROXY_URL)
    client = TelegramClient(str(SESSION), API_ID, API_HASH, proxy=proxy)
    await client.connect()
    if not await client.is_user_authorized():
        await client.disconnect()
        raise SystemExit("Telegram session expired - run: python scrape.py --login")

    entities = {}
    for link in read_channels():
        kind, value = link_to_target(link)
        try:
            ent = await resolve(client, kind, value)
        except Exception:
            ent = None
        if ent is None:
            print(f"SKIP cannot access: {link}")
            continue
        title = getattr(ent, "title", None) or value
        entities[title] = ent
        print(f"MEDIA channel ok: {title}")

    tmpdir = Path(tempfile.mkdtemp(prefix="tgmedia_", dir=str(MEDIA_DIR)))
    for stale in MEDIA_DIR.glob("tgmedia_*"):
        if stale != tmpdir:
            shutil.rmtree(stale, ignore_errors=True)
    try:
        total = len(jobs)
        for ji, (code, channel, msg_id) in enumerate(jobs, 1):
            saved_photos, saved_videos, notes = [], [], []
            print(f"[{ji}/{total}] {code} ({channel} #{msg_id})...", flush=True)
            ent = entities.get(channel)
            if ent is None:
                results[code] = ([], [], [f"channel not accessible: {channel}"])
                print(f"[{ji}/{total}] {code} done: no channel", flush=True)
                continue
            msgs, err = await fetch_post_messages(client, ent, msg_id)
            if err:
                results[code] = ([], [], [err])
                print(f"[{ji}/{total}] {code} done: {err}", flush=True)
                continue
            photo_n, video_n = 0, 0
            for m in msgs:
                kind = media_kind(m)
                if kind is None:
                    continue
                if kind == "photo":
                    photo_n += 1
                    dst = MEDIA_DIR / f"{code}-{photo_n}.webp"
                else:
                    video_n += 1
                    ext = video_ext(m)
                    name = f"{code}-V{ext}" if video_n == 1 else f"{code}-V-{video_n}{ext}"
                    dst = MEDIA_DIR / name
                if dst.exists():
                    notes.append(f"{dst.name} exists, skipped")
                    (saved_photos if kind == "photo" else saved_videos).append(dst.name)
                    continue
                tmp = tmpdir / f"{code}_{m.id}"
                try:
                    got = await asyncio.wait_for(
                        client.download_media(m, file=str(tmp)), 180
                    )
                except asyncio.TimeoutError:
                    notes.append(f"msg {m.id} download timed out")
                    if kind == "photo":
                        photo_n -= 1
                    else:
                        video_n -= 1
                    continue
                except errors.FloodWaitError as exc:
                    await asyncio.sleep(min(exc.seconds + 1, 60))
                    try:
                        got = await asyncio.wait_for(
                            client.download_media(m, file=str(tmp)), 180
                        )
                    except asyncio.TimeoutError:
                        notes.append(f"msg {m.id} download timed out")
                        if kind == "photo":
                            photo_n -= 1
                        else:
                            video_n -= 1
                        continue
                    except Exception as exc2:
                        notes.append(f"msg {m.id} download failed: {exc2}")
                        if kind == "photo":
                            photo_n -= 1
                        else:
                            video_n -= 1
                        continue
                except Exception as exc:
                    notes.append(f"msg {m.id} download failed: {exc}")
                    if kind == "photo":
                        photo_n -= 1
                    else:
                        video_n -= 1
                    continue
                if got is None:
                    if kind == "photo":
                        photo_n -= 1
                    else:
                        video_n -= 1
                    continue
                tmp = Path(got)
                if kind == "photo":
                    try:
                        ok, size_kb = convert_to_webp(tmp, dst)
                    except Exception as exc:
                        notes.append(f"{dst.name} convert failed: {exc}")
                        tmp.unlink(missing_ok=True)
                        continue
                    tmp.unlink(missing_ok=True)
                    saved_photos.append(dst.name)
                    if not ok:
                        notes.append(f"{dst.name} is {size_kb:.0f}KB (over 100KB)")
                else:
                    shutil.move(str(tmp), str(dst))
                    saved_videos.append(dst.name)
            if not saved_photos and not saved_videos and not notes:
                notes.append("post has no photo/video")
            results[code] = (saved_photos, saved_videos, notes)
            print(f"[{ji}/{total}] {code} done: "
                  f"{len(saved_photos)} photos, {len(saved_videos)} videos", flush=True)
            await asyncio.sleep(0.5)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
        await client.disconnect()
    return results


def main(args):
    posts = json.loads(POSTS_FILE.read_text(encoding="utf-8"))
    print(f"loaded {len(posts)} posts")
    exact, nums = build_index(posts)

    wb = openpyxl.load_workbook(SRC)
    ws = wb.worksheets[0]
    header = {c.value: c.column for c in ws[1] if c.value}
    for name in FIELDS + ["کد محصول", "فروشنده", "دسته بندی"]:
        if name not in header:
            raise SystemExit(f"missing column: {name}")
    photo_col = header.get("نام عکس محصول ")

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
             "category": [], "media": []}

    matched = {}  # code -> (row, channel, msg_id)
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

        matched.setdefault(code, (row, post["channel"], post["id"]))

        if not fields:
            stats["found_empty"].append(
                (row, code, post["channel"], post["id"], post["text"][:200])
            )
            continue

        written = {}
        for name in FIELDS:
            if name not in fields:
                continue
            cell = ws.cell(row, COL[name])
            if cell.value not in (None, ""):
                stats["kept_existing"].append((row, code, name, cell.value))
                continue
            cell.value = fields[name]
            written[name] = fields[name]
        if written:
            stats["filled_rows"] += 1
            stats["filled"].append((row, code, post["channel"], post["id"], written))

    for row in range(2, ws.max_row + 1):
        cat_cell = ws.cell(row, COL["دسته بندی"])
        if cat_cell.value not in (None, ""):
            continue
        usage = ws.cell(row, COL["کاربرد"]).value
        if usage in (None, ""):
            continue
        cat_cell.value = usage_to_category(usage)
        stats["category"].append((row, usage, cat_cell.value))

    if args.download_media and matched:
        MEDIA_DIR.mkdir(exist_ok=True)
        jobs = [(code, ch, mid) for code, (_row, ch, mid) in matched.items()]
        print(f"downloading media for {len(jobs)} products...")
        results = asyncio.run(download_all(jobs))
        for code, (photos, videos, notes) in results.items():
            row = matched[code][0]
            stats["media"].append((row, code, photos, videos, notes))
            # Reconcile with disk: flaky album fetches may miss files that
            # a previous run already saved, so list everything on disk.
            disk_photos = sorted(p.name for p in MEDIA_DIR.glob(f"{code}-*.webp"))
            disk_videos = sorted(p.name for p in MEDIA_DIR.glob(f"{code}-V*"))
            if photo_col and disk_photos:
                cell = ws.cell(row, photo_col)
                if cell.value in (None, ""):
                    cell.value = " | ".join(disk_photos)

    wb.save(DST)
    print(f"saved {DST}")

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
    lines.append(f"== دسته بندی derived from کاربرد ({len(stats['category'])}) ==")
    for r, u, c in stats["category"]:
        lines.append(f"  row {r}: {u}  ->  {c}")
    if args.download_media:
        lines.append(f"== media ({len(stats['media'])}) ==")
        for r, c, photos, videos, notes in stats["media"]:
            files = photos + videos
            msg = ", ".join(files) if files else "nothing downloaded"
            extra = f" ({'; '.join(notes)})" if notes else ""
            lines.append(f"  row {r}: {c} -> {msg}{extra}")
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
          f"empty_post={len(stats['found_empty'])} invalid={len(stats['invalid'])}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--download-media", action="store_true",
                    help="download product photos (WebP <100KB) and videos from Telegram")
    ap.add_argument("--codes", default="",
                    help="only process these product codes, comma-separated (for testing)")
    args = ap.parse_args()
    main(args)
