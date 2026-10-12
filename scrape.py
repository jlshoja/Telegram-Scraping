import argparse
import asyncio
import json
import os
import sys
from pathlib import Path
from urllib.parse import urlparse

import openpyxl
from dotenv import load_dotenv
from telethon import TelegramClient, errors, functions, types

from tgcommon import normalize

load_dotenv()

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

API_ID = int(os.environ.get("TELEGRAM_API_ID", "0"))
API_HASH = os.environ.get("TELEGRAM_API_HASH", "")
PROXY_URL = os.environ.get("TG_PROXY", "").strip()
CHANNELS_FILE = "Telegram channels.XLSX"
CACHE_DIR = Path("cache")
POSTS_FILE = CACHE_DIR / "posts.json"
SESSION = CACHE_DIR / "tg"


def parse_proxy(url):
    if not url:
        return None
    u = urlparse(url)
    if not u.scheme or not u.hostname:
        return None
    return (u.scheme.lower(), u.hostname, u.port or 1080, True, u.username, u.password)


def read_channels():
    wb = openpyxl.load_workbook(CHANNELS_FILE, read_only=True)
    ws = wb.worksheets[0]
    links = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        link = (row[0] or "").strip() if row else ""
        if link:
            links.append(link)
    wb.close()
    return links


def link_to_target(link):
    link = link.strip()
    if "t.me/" in link:
        link = link.split("t.me/", 1)[1]
    link = link.split("?")[0].rstrip("/")
    if link.startswith("+"):
        return ("invite", link[1:])
    return ("username", link)


async def resolve(client, kind, value):
    if kind == "username":
        try:
            return await client.get_entity(value)
        except errors.UsernameNotOccupiedError:
            return None
        except errors.ChannelPrivateError:
            return None

    try:
        r = await client(functions.messages.ImportChatInviteRequest(value))
        return r.chats[0]
    except errors.UserAlreadyParticipantError:
        pass
    except (errors.InviteHashInvalidError, errors.InviteHashExpiredError):
        return None
    except errors.InviteRequestSentError:
        return None

    try:
        r = await client(functions.messages.CheckChatInviteRequest(value))
        if isinstance(r, types.ChatInviteAlready):
            return r.chat
        return None
    except Exception:
        return None


async def ensure_auth(client, args):
    if await client.is_user_authorized():
        return
    phone = args.phone or os.environ.get("TG_PHONE") or ""
    if not phone:
        phone = input("Phone number (+98...): ").strip()

    if args.code:
        try:
            await client.sign_in(phone=phone, code=args.code)
            print("LOGIN_OK")
            return
        except errors.SessionPasswordNeededError:
            pw = args.password or os.environ.get("TG_PASSWORD") or ""
            if not pw:
                pw = input("2FA password: ").strip()
            await client.sign_in(password=pw)
            print("LOGIN_OK")
            return
        except (ValueError, errors.PhoneCodeInvalidError):
            pass

    await client.send_code_request(phone)
    print("CODE_SENT", flush=True)
    if args.code:
        try:
            await client.sign_in(phone=phone, code=args.code)
            print("LOGIN_OK")
            return
        except errors.SessionPasswordNeededError:
            pw = args.password or os.environ.get("TG_PASSWORD") or ""
            if not pw:
                pw = input("2FA password: ").strip()
            await client.sign_in(password=pw)
            print("LOGIN_OK")
            return
        except errors.PhoneCodeInvalidError:
            print("OLD_CODE_INVALID, a new code was sent", flush=True)
    code = ""
    if sys.stdin.isatty() and not args.wait_file:
        code = input("Enter the code you received: ").strip()
    else:
        pending = CACHE_DIR / "pending_code.txt"
        print(f"waiting for code file: {pending}", flush=True)
        for _ in range(600):
            await asyncio.sleep(1)
            if pending.exists():
                code = pending.read_text(encoding="utf-8").strip()
                pending.unlink()
                break
        if not code:
            print("TIMEOUT waiting for code")
            sys.exit(2)
    try:
        await client.sign_in(phone=phone, code=code)
    except errors.SessionPasswordNeededError:
        pw = args.password or os.environ.get("TG_PASSWORD") or ""
        if not pw and sys.stdin.isatty():
            pw = input("2FA password: ").strip()
        if not pw:
            print("2FA_PASSWORD_NEEDED")
            sys.exit(3)
        await client.sign_in(password=pw)
    print("LOGIN_OK")


def load_cache():
    if POSTS_FILE.exists():
        return json.loads(POSTS_FILE.read_text(encoding="utf-8"))
    return []


def save_cache(posts):
    CACHE_DIR.mkdir(exist_ok=True)
    POSTS_FILE.write_text(
        json.dumps(posts, ensure_ascii=False, indent=1), encoding="utf-8"
    )


async def fetch(args):
    links = read_channels()
    CACHE_DIR.mkdir(exist_ok=True)
    proxy = parse_proxy(PROXY_URL)
    if proxy:
        print(f"proxy {PROXY_URL}")
    client = TelegramClient(str(SESSION), API_ID, API_HASH, proxy=proxy)
    await client.connect()
    await ensure_auth(client, args)
    if args.login:
        await client.disconnect()
        return

    posts = load_cache()
    seen = {(p["channel"], p["id"]) for p in posts}
    fresh = 0

    for link in links:
        kind, value = link_to_target(link)
        entity = await resolve(client, kind, value)
        if entity is None:
            print(f"SKIP cannot access: {link}")
            continue
        try:
            title = entity.title or value
        except AttributeError:
            title = value
        print(f"SCAN {title} ({link})")

        count = 0
        try:
            async for msg in client.iter_messages(entity, limit=args.limit):
                text = msg.text or ""
                if not text.strip():
                    continue
                key = (title, msg.id)
                if key in seen:
                    continue
                seen.add(key)
                posts.append(
                    {
                        "channel": title,
                        "id": msg.id,
                        "date": str(msg.date),
                        "text": normalize(text),
                    }
                )
                count += 1
                fresh += 1
                if fresh % 200 == 0:
                    save_cache(posts)
                    print(f"  ... {fresh} new posts")
        except errors.ChannelPrivateError:
            print(f"SKIP private/no access: {title}")
            continue
        except errors.ChatAdminRequiredError:
            print(f"SKIP admin required: {title}")
            continue
        print(f"  +{count} posts")
        save_cache(posts)

    save_cache(posts)
    print(f"DONE total={len(posts)} new={fresh}")
    await client.disconnect()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--login", action="store_true", help="only authenticate")
    ap.add_argument("--code", help="login code")
    ap.add_argument("--password", help="2FA password")
    ap.add_argument("--phone", help="phone number")
    ap.add_argument("--limit", type=int, default=None, help="max messages per channel")
    ap.add_argument("--wait-file", action="store_true",
                    help="send the code, then wait for it in cache/pending_code.txt")
    args = ap.parse_args()
    asyncio.run(fetch(args))


if __name__ == "__main__":
    main()
