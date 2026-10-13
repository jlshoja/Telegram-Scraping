# Telegram Scraper + Excel Filler — Command Guide

## 📁 Project folder
`E:\Luxbaz\All Codes\Projects\Telegram Scraping`

---

## 🖱 Easiest: double-click `run.bat`

A menu with all running options — just type the number. You can also pass the
option directly: `run.bat 2` (fill + media), `run.bat 4 FA-1959,TR-402`
(specific codes).

| # | Action |
|---|--------|
| 1 | Fill Excel (cached posts) |
| 2 | Fill Excel + download media |
| 3 | Rescan channels, then fill + media |
| 4 | Fill specific codes |
| 5 | Refresh old values with fresh data |
| 6 | Telegram login (first time only) |

---

## 🚀 Quick start

### 1️⃣ First‑time login (only once)

```powershell
python scrape.py --login --phone "+989143463191" --wait-file
```

- A login code is sent to your Telegram app.
- The script waits for you to **write the code to `cache\pending_code.txt`** (open that file and paste the 5‑digit code you received).
- After the code is placed, the script signs in, scans the 5 channels, saves the post cache to `cache/posts.json`, and exits.

> After the first successful run the script remembers the session in `cache\tg.session`.  
> Subsequent runs skip the login step and just scan.

### 2️⃣ Fill / update the Excel file

```powershell
python fill.py
```

- Reads the cached posts (`cache/posts.json`)
- Extracts **ابعاد**, **جنس**, **نوع بند**, **کاربرد** from each post
- Writes the values into `Product Database - filled.xlsx` (the original `Product Database.xlsx` stays untouched)
- Produces `report.txt` with statistics and any problems (missing codes, empty posts, etc.)

> **Result:** 25 of the 46 coded rows are filled; the remaining 21 either have no matching post or the post only contains price info.

### 3️⃣ Add / change channels

Edit `Telegram channels.XLSX` – it has a single column named **link** (the 5 URLs you already have, plus any new ones you want to scan). Then re‑run:

```powershell
python scrape.py      # (or python scrape.py --wait-file for a fresh login)
python fill.py
```

> `scrape.py` is incremental: it only downloads posts it hasn't seen before (tracked in `cache/posts.json`).

### 4️⃣ New products? Just rerun — old ones are skipped

Reruns only touch what's new:
- `python fill.py` never overwrites filled cells — new Excel rows get matched and filled, old rows are left alone
- `python fill.py --download-media` skips photos/videos already in `media/`
- Each product's media source post is tracked in `media/manifest.json` — if a product gets reposted, old files are deleted and fresh ones downloaded automatically

```powershell
python scrape.py                  # pick up new channel posts (incremental)
python fill.py --download-media   # fill new rows + download new media
```

Only specific new products? Filter by code:

```powershell
python fill.py --download-media --codes FA-9999,TR-123
```

Want to overwrite old values with fresh data (e.g. a post was corrected)?

```powershell
python fill.py --refresh                       # refresh text fields
python fill.py --download-media --refresh      # refresh text fields + photo names
```

### 5️⃣ Download product photos & videos (optional)

```powershell
python fill.py --download-media
```

- Re-fetches each matched product post from Telegram and downloads its media into `media/`
- Photos → converted to **WebP under 100KB**, named `{code}-{n}.webp` (e.g. `FA-2354-1.webp`, `FA-2354-2.webp`)
- Videos → downloaded as-is, named `{code}-V.mp4` (e.g. `FA-2354-V.mp4`)
- Filenames are written into the `نام عکس محصول` column (only if empty)
- Details per product go into the `== media ==` section of `report.txt`
- Reruns are safe: existing files are skipped, only missing ones are fetched
- For a quick test on specific products: `python fill.py --download-media --codes FA-1959,TR-402`

> **Result:** 127 photos + 4 videos for all 32 matched products. Connection via Telegram can be flaky — if some products show `fetch timed out` in the report, just rerun the command.

### 5️⃣ Quick checklist

| Action | Command |
|--------|---------|
| **First login** (one‑off) | `python scrape.py --login --phone "+989143463191" --wait-file` |
| **Regular fill** (cached) | `python fill.py` |
| **Fill + download media** | `python fill.py --download-media` |
| **Only new products** | `python scrape.py` then `python fill.py --download-media` (old rows/files skipped) |
| **Refresh old values** | `python fill.py --refresh` |
| **Re‑scan channels** | `python scrape.py` then `python fill.py` |
| **Check the report** | `type report.txt` or open `report.txt` |

### 6️⃣ File overview

| File | Purpose |
|------|---------|
| `scrape.py` | Authenticates via Telethon, reads the 5 channels, caches all posts to `cache/posts.json` |
| `fill.py` | Extracts **ابعاد / جنس / نوع بند / کاربرد**, writes `Product Database - filled.xlsx`, downloads media with `--download-media` |
| `tgcommon.py` | Shared helpers: normalize Persian digits, extract codes, extract the four fields |
| `media/` | Downloaded product photos (WebP <100KB) and videos (not committed to git) |
| `cache/posts.json` | 1634 raw posts (downloaded once, reused every run) |
| `Telegram channels.XLSX` | The 5 channel URLs (add/remove as needed) |
| `Product Database.xlsx` | Original data – **never modified** |
| `Product Database - filled.xlsx` | Filled output (25 rows populated) |
| `report.txt` | Summary of filled / missing / empty rows |

---

## 🛠 One‑time setup (already done)

- The scripts run locally via your existing SOCKS5 proxy (`127.0.0.1:10808`). No Google OAuth is required.
- The first auth flow may require entering a code; after that `cache\tg.session` persists the login.

--- 

**That’s it** – just two commands and you’re all set. Happy scraping!