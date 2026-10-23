"""Download product photos from Telegram into Product Images/<CODE>/.

For every coded product in the product database (Product Database.xlsx,
codes like FA-1863) the matching Telegram post is located in the cached
posts (cache/posts.json); the match is by full code, or by the bare number
1863 like the channels post it. Then the photos of that post are downloaded:

  - photo 1  -> <CODE>-main.webp     (WooCommerce featured image)
  - photo 2+ -> <CODE>-02.webp, <CODE>-03.webp, ...

Photos are converted to WebP (max_kb from settings.json, images.max_webp_kb).
Videos are skipped. Codes whose folder already contains files are skipped
unless --refresh. The last-scraped date per channel is printed from the
posts cache; scrape.py (run.bat option 2) adds only posts newer than the
last scan, so rerunning this script only processes what is new.

Usage:
  python download_images.py
  python download_images.py --limit 5     # only the first 5 pending codes
  python download_images.py --refresh     # redownload existing folders too
"""

import argparse
import asyncio
import json
import re
import shutil
import sys
import tempfile
from pathlib import Path

from PIL import Image
import openpyxl
from telethon import TelegramClient, errors

from fill import build_index, build_seller_affinity, best_candidate
from scrape import (API_HASH, API_ID, PROXY_URL, SESSION, link_to_target,
                    parse_proxy, read_channels, resolve)
from tgcommon import normalize

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE_DIR = Path(__file__).resolve().parent
SETTINGS_FILE = "settings.json"
MAX_WEBP_SIZE_KB = 100


def project_path(value):
    p = Path(value)
    return p if p.is_absolute() else (BASE_DIR / p)


def convert_to_webp(src, dst, max_kb=MAX_WEBP_SIZE_KB):
    """Convert an image to WebP under max_kb. Returns (ok, size_kb)."""
    img = Image.open(src).convert("RGB")
    tmp = dst.with_suffix(".tmp.webp")
    quality, scale = 80, 1.0
    for _ in range(14):
        cur = img
        if scale < 1.0:
            cur = img.resize(
                (max(1, int(img.width * scale)),
                 max(1, int(img.height * scale))),
                Image.LANCZOS)
        cur.save(tmp, "WEBP", quality=quality, method=6)
        size_kb = tmp.stat().st_size / 1024
        if size_kb <= max_kb:
            tmp.replace(dst)
            return True, size_kb
        if quality > 32:
            quality -= 12
        else:
            scale *= 0.8
            quality = 80
    tmp.replace(dst)
    return False, size_kb


def last_scrape_dates(posts):
    per_channel = {}
    for p in posts:
        d = str(p.get("date", ""))
        ch = per_channel.setdefault(p["channel"], "")
        if d > ch:
            per_channel[p["channel"]] = d
    overall = max((str(p.get("date", "")) for p in posts), default="")
    return per_channel, overall


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
                timeout)
            for m in batch or []:
                if m is not None and m.id != msg_id \
                        and getattr(m, "grouped_id", None) == gid:
                    msgs.append(m)
            msgs.sort(key=lambda m: m.id)
        except Exception:
            pass
    return msgs, ""


async def download_all(jobs, images_dir, max_kb):
    """jobs: list of (code, channel_title, msg_id). Returns dict code -> (files, notes)."""
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
        print(f"channel ok: {title}")
    tmpdir = Path(tempfile.mkdtemp(prefix="tgimg_", dir=str(images_dir)))
    momentum = {"timeouts": 0, "ok": 0}
    try:
        total = len(jobs)
        for ji, (code, channel, msg_id) in enumerate(jobs, 1):
            saved, notes = [], []
            ent = entities.get(channel)
            if ent is None:
                results[code] = ([], [f"channel not accessible: {channel}"])
                continue
            msgs, err = await fetch_post_messages(client, ent, msg_id)
            if err:
                results[code] = ([], [err])
                continue
            folder = Path(images_dir) / code
            folder.mkdir(exist_ok=True)
            for photo_n, m in enumerate((m for m in msgs if getattr(m, "photo", None)), 1):
                stem = f"{code}-main" if photo_n == 1 else f"{code}-{photo_n:02d}"
                dst = folder / f"{stem}.webp"
                if dst.exists():
                    saved.append(dst.name)
                    continue
                tmp = tmpdir / f"{code}_{m.id}"
                try:
                    got = await asyncio.wait_for(
                        client.download_media(m, file=str(tmp)), 180)
                except asyncio.CancelledError:
                    print("\ninterrupted by Ctrl+C, exiting", flush=True)
                    raise
                except errors.FloodWaitError as exc:
                    momentum["timeouts"] = 0
                    await asyncio.sleep(min(exc.seconds + 1, 60))
                    try:
                        got = await asyncio.wait_for(
                            client.download_media(m, file=str(tmp)), 180)
                    except asyncio.CancelledError:
                        print("\ninterrupted by Ctrl+C, exiting", flush=True)
                        raise
                    except (asyncio.TimeoutError, Exception) as exc:
                        notes.append(f"msg {m.id} retry failed: {exc}")
                        continue
                except asyncio.TimeoutError:
                    momentum["timeouts"] += 1
                    notes.append(f"msg {m.id} download timed out")
                    if momentum["timeouts"] >= 3 and momentum["ok"] == 0:
                        print(
                            "3 downloads timed out with nothing saved "
                            "- stopping. Check VPN/proxy and run again",
                            flush=True,
                        )
                        raise SystemExit(3)
                    continue
                except Exception as exc:
                    notes.append(f"msg {m.id} download failed: {exc}")
                    continue
                momentum["timeouts"] = 0
                momentum["ok"] += 1
                tmp = Path(got)
                try:
                    ok, size_kb = convert_to_webp(tmp, dst, max_kb)
                except Exception as exc:
                    notes.append(f"{dst.name} convert failed: {exc}")
                    continue
                tmp.unlink(missing_ok=True)
                saved.append(dst.name)
                if not ok:
                    notes.append(f"{dst.name} is {size_kb:.0f}KB (over {max_kb}KB)")
            if not msgs:
                notes.append("post not found")
            elif not saved and not notes:
                notes.append("post has no photos")
            results[code] = (saved, notes)
            print(f"[{ji}/{total}] {code} done: {len(saved)} photos", flush=True)
            await asyncio.sleep(0.5)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
        await client.disconnect()
    return results


def main(args):
    settings = json.loads(project_path(args.settings).read_text(encoding="utf-8"))
    src = project_path(settings["files"]["source"])
    images_dir = project_path(settings["files"]["images_dir"])
    src_cols = settings["source_columns"]
    code_header = src_cols["code"]
    seller_header = src_cols.get("seller", "فروشنده")
    max_kb = int(settings["images"].get("max_webp_kb", MAX_WEBP_SIZE_KB))
    code_re = re.compile(settings["values"]["code_pattern"])

    posts_file = project_path("cache/posts.json")
    if not posts_file.exists():
        raise SystemExit("cache/posts.json not found - run run.bat option 2 first")
    posts = json.loads(posts_file.read_text(encoding="utf-8-sig"))
    print(f"loaded {len(posts)} posts")

    per_channel, overall = last_scrape_dates(posts)
    for ch, d in sorted(per_channel.items()):
        print(f"last scraped: {ch} -> {d}")
    print(f"last scraped (all): {overall}")

    wb = openpyxl.load_workbook(src)
    ws = wb.worksheets[0]
    header = {c.value: c.column for c in ws[1] if c.value}
    if code_header not in header:
        raise SystemExit(f"missing column: {code_header}")
    code_col = header[code_header]
    seller_col = header.get(seller_header)

    entries = []
    seen = set()
    for row in range(2, ws.max_row + 1):
        raw = ws.cell(row, code_col).value
        code = normalize(str(raw)).strip().upper() \
            if raw not in (None, "") else ""
        m = code_re.match(code) if code else None
        if not m:
            continue
        if code in seen:
            continue
        seen.add(code)
        seller = ""
        if seller_col:
            v = ws.cell(row, seller_col).value
            seller = normalize(str(v)).strip() if v not in (None, "") else ""
        entries.append((row, code, m.group(2), seller))
    wb.close()
    print(f"codes in database: {len(entries)}")

    exact, nums = build_index(posts)
    seller_affinity = build_seller_affinity(
        [(row, seller, num) for row, _c, num, seller in entries], nums)

    jobs = []
    skipped_existing = not_found = ambiguous = 0
    for row, code, num, seller in entries:
        folder = images_dir / code
        if not args.refresh and (folder.exists() and any(folder.iterdir())):
            skipped_existing += 1
            continue
        cands = exact.get(code) or nums.get(num)
        if not cands:
            not_found += 1
            print(f"- {code}: no matching post found")
            continue
        affinity = {}
        if seller:
            for (sel, chan), w in seller_affinity.items():
                if sel == seller:
                    affinity[chan] = w
        (mtype, fields, post), ties = best_candidate(cands, affinity)
        if ties > 1:
            ambiguous += 1
        jobs.append((code, post["channel"], post["id"]))

    if args.limit:
        jobs = jobs[:args.limit]
    if not jobs:
        print("nothing to download "
              f"(up to date: {skipped_existing} already have images)")
        return

    print(f"download images for {len(jobs)} products "
          f"(skipped existing: {skipped_existing}, no post: {not_found}, "
          f"ambiguous: {ambiguous})", flush=True)
    images_dir.mkdir(exist_ok=True)
    try:
        results = asyncio.run(download_all(jobs, images_dir, max_kb))
    except SystemExit as exc:
        print(exc)
        return
    except Exception as exc:
        print(f"download failed: {type(exc).__name__} {exc}")
        print("(is the proxy/VPN running? check TG_PROXY in .env)")
        return

    got = sum(len(files) for files, _n in results.values())
    problems = [(c, n) for c, (f, n) in results.items() if n]
    print(f"saved {got} photos for {len(results)} codes")
    for code, notes in problems:
        print(f"NOTE {code}: {'; '.join(notes)}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--settings", default=SETTINGS_FILE)
    ap.add_argument("--limit", type=int, default=0,
                    help="only process the first N pending codes (test)")
    ap.add_argument("--refresh", action="store_true",
                    help="redownload existing folders too")
    main(ap.parse_args())
