# Telegram Scraper + Excel Filler — Command Guide

## 📁 Project folder
`E:\Luxbaz\All Codes\Projects\Telegram Scraping`

---

## 🖱 Easiest: double-click `run.bat`

A menu with all running options — just type the number. You can also pass the
option directly: `run.bat 2` (rescan + fill), `run.bat 3 FA-1959,TR-402`
(specific codes).

| # | Action |
|---|--------|
| 1 | Fill Excel (cached posts) |
| 2 | Rescan channels, then fill |
| 3 | Fill specific codes |
| 4 | Refresh old values with fresh data |
| 5 | Telegram login (first time only) |

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

```powershell
python scrape.py   # pick up new channel posts (incremental)
python fill.py     # fill new rows
```

Only specific new products? Filter by code:

```powershell
python fill.py --codes FA-9999,TR-123
```

Want to overwrite old values with fresh data (e.g. a post was corrected)?

```powershell
python fill.py --refresh   # refresh text fields
```

### 5️⃣ Quick checklist

| Action | Command |
|--------|---------|
| **First login** (one‑off) | `python scrape.py --login --phone "+989143463191" --wait-file` |
| **Regular fill** (cached) | `python fill.py` |
| **Only new products** | `python scrape.py` then `python fill.py` (old rows skipped) |
| **Test specific codes** | `python fill.py --codes FA-1959,TR-402` |
| **Refresh old values** | `python fill.py --refresh` |
| **Re‑scan channels** | `python scrape.py` then `python fill.py` |
| **Build WooCommerce import** | `python build_woocommerce.py` (needs `قیمت فروش` + `رنگ` filled first) |
| **Check the report** | `type report.txt` or open `report.txt` |

### 6️⃣ File overview

| File | Purpose |
|------|---------|
| `scrape.py` | Authenticates via Telethon, reads the 5 channels, caches all posts to `cache/posts.json` |
| `fill.py` | Extracts **ابعاد / جنس / نوع بند / کاربرد**, writes `Product Database - filled.xlsx` |
| `build_woocommerce.py` | Builds `woocommerce_import_output.xlsx` from the filled database, using `woocommerce_import_new.xlsx` as the read-only input template (paths/settings in `settings.json`) |
| `woocommerce_import_new.xlsx` | **Input template** (1515 rows, 39 columns) — never modified |
| `woocommerce_import_output.xlsx` | **Generated output** to upload to WooCommerce |
| `tgcommon.py` | Shared helpers: normalize Persian digits, extract codes, extract the four fields |
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
