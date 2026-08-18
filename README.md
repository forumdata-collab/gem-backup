# Gem Backup

Backup your **Gemini Gems** (custom AI instructions) to Excel before Google
sunset them in October 2026. Paste a public Google Drive folder URL → the
server downloads every Gem, parses the instruction bytes, and packages a
`.xlsx` (plus a `.zip` of any attachments) for download. Files are kept on the
server for **30 minutes** and then deleted.

**Live demo:** https://gembackup.hermesagent.de5.net

## How it works (the method)

No OAuth, no Google API keys, no AI — a fully deterministic pipeline using
only public Drive endpoints:

| Step | Mechanism |
|---|---|
| 1. List folder | `https://drive.google.com/embeddedfolderview?id=<FOLDER_ID>#list` renders a server-side HTML listing for any folder shared as **"Anyone with the link"**. Parse `flip-entry` blocks → file ID, name, type (gem / file / subfolder), modified date. Subfolders are recursed (depth ≤ 3, cycle-guarded). |
| 2. Download | `https://drive.google.com/uc?export=download&id=<FILE_ID>` streams binaries without auth. Files > 100 MB behind Google's virus-scan get a `confirm=` token — the parser extracts it from the HTML and retries with the token. |
| 3. Parse Gem | A Gem file (`application/vnd.google-gemini.gem`) is a tiny **protobuf**. Top-level field **2** wraps a message with: field 1 = name, field 2 = description, field 3 = instructions (the system prompt). A ~40-line minimal protobuf walker extracts all three. |
| 4. Build xlsx | `openpyxl`. **Excel hard-caps cells at 32,767 chars** — any description/instructions longer than 30,000 chars are split across rows with a 【續上】 continuation marker, so no content is silently truncated. |
| 5. Zip | Non-gem files (PDFs, images…) are zipped with sanitized, de-duplicated names (`a b.jpg` + `a-b.jpg` → `a_b.jpg` + `a_b (1).jpg`). |
| 6. Handout | Job ID → `GET /d/<job>/<file>` serves both files; a sweep (per-request + 15-min daemon) deletes jobs older than 30 minutes from `/tmp`. |

## Infrastructure

One self-contained file, zero frameworks:

```
app.py          # HTTP server (stdlib http.server) + all pipeline logic (~430 lines)
test_sanity.py  # 59-assert sanity suite (unit + optional live-folder checks)
LICENSE         # MIT
docs/index.html # this project page (GitHub Pages)
```

**Deploy (reference layout — adapt to your host):**

```bash
# 1. systemd unit (gembackup-web.service)
[Service]
WorkingDirectory=/srv/gembackup
ExecStart=/usr/bin/python3 /srv/gembackup/app.py
Restart=always

# 2. nginx reverse proxy (port 80 → 127.0.0.1:8301)
server {
    listen 80;
    server_name YOUR_DOMAIN;
    location / {
        proxy_pass http://127.0.0.1:8301;
        proxy_set_header Host $host;
        proxy_read_timeout 300s;
        client_max_body_size 10k;   # API body is one small JSON only
    }
}

# 3. Cloudflare (optional): orange-cloud A record + SSL mode "Flexible"
#    (origin is plain HTTP; HTTPS terminates at the CF edge)
```

Run tests:

```bash
pip install requests openpyxl
python3 test_sanity.py          # pure unit tests
TEST_FOLDER_ID=<your_folder_id> python3 test_sanity.py   # + live end-to-end
```

## Operational limits

- **Folder must be shared as "Anyone with the link"** (public listing is required — no auth flows).
- Single attachment ≤ 200 MB; total job ≤ 1 GB; 4 concurrent jobs (HTTP 429 beyond).
- Subfolder recursion depth cap: 3 levels.
- Files auto-delete 30 minutes after creation (TTL sweep on every request + daemon).
- No AI, no analytics, no persistence: the server never stores anything beyond `/tmp` job output.

## Security notes

- Downloads are pinned to `drive.google.com` URLs built from parsed file IDs.
- Job IDs are UUID hex, strictly validated; the download route resolves `realpath` and rejects any path escaping the job directory.
- Nothing user-supplied is ever executed or rendered (JSON only; filenames sanitized for zip).

## License

MIT — see [LICENSE](LICENSE).