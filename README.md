# Gem Backup

Save your Gemini Gems before they disappear.

Export your Gemini Gems' custom instructions from a public Google Drive folder
into a searchable Excel backup — with attachments included. No Google login ·
No API key · No AI · No data stored permanently.

30-second backup → XLSX + ZIP.

> Google is sunsetting **Gems** (custom Gemini instructions) in **October 2026**.
> This tool lets you own your Gems as plain, portable data before they go.

## How to use

1. **Paste** your public Google Drive folder link (`drive.google.com/drive/folders/…`). The folder must be shared as **"Anyone with the link"** (Viewer).
2. **Backup** — the server lists the folder, downloads every Gem, parses the instruction bytes.
3. **Download** — an Excel workbook (all Gems) + a ZIP (machine-readable backup + attachments). Both are deleted from the server after **30 minutes**.

## Output

### Gems_Backup.xlsx

| Sheet | Content |
|---|---|
| Manifest | Backup time, tool version, source folder name/ID, gems found, attachments, total size, failed files |
| Gems | Name, description, full instructions (system prompt), gem link, **Drive file ID**, last modified, status |

Excel hard-caps cells at 32,767 chars — any text longer than 30,000 chars is
split across rows with a 【續上】 continuation marker, so nothing is silently
truncated.

### Gems_Backup.zip — machine-readable canonical backup

```
gem-backup/
├── manifest.json            # tool/version/time/source folder + per-file sha256
├── gems/
│   ├── gem-001/
│   │   ├── metadata.json    # index, name, Drive ID, modified, status
│   │   ├── instructions.txt # the full system prompt
│   │   └── description.txt
│   └── gem-002/ …
└── attachments/             # original files (sanitized, de-duplicated names)
```

Every file carries its **sha256** checksum in `manifest.json`, so a restore
tool (or a human) can verify integrity later. The Excel is the human-readable
view; the ZIP is the disaster-recovery archive.

## How it works (the method)

No OAuth, no Google API keys, no AI — a fully deterministic pipeline using
only public Drive endpoints:

| Step | Mechanism |
|---|---|
| 1. List folder | `https://drive.google.com/embeddedfolderview?id=<FOLDER_ID>#list` renders a server-side HTML listing for any folder shared as "Anyone with the link". Parse `flip-entry` blocks → file ID, name, type (gem / file / subfolder), modified date. Subfolders are recursed (depth ≤ 3, cycle-guarded). Folder name comes from the page `<title>`. |
| 2. Download | `https://drive.google.com/uc?export=download&id=<FILE_ID>` streams binaries without auth, with connect + read timeouts. Files > 100 MB behind Google's virus-scan get a `confirm=` token — the parser extracts it and retries with the token. |
| 3. Parse Gem | A Gem file (`application/vnd.google-gemini.gem`) is a tiny **protobuf**. Top-level field **2** wraps a message with: field 1 = name, field 2 = description, field 3 = instructions (the system prompt). A minimal protobuf walker extracts all three. |
| 4. Build xlsx | `openpyxl` — Manifest sheet + chunked Gems sheet. |
| 5. Build zip | `gem-backup/` tree above, sha256 hashed while streaming. |
| 6. Handout | Job ID → `GET /d/<job>/<file>` serves both files; a sweep (per-request + 15-min daemon) deletes jobs older than 30 minutes from `/tmp`. |

## Infrastructure

One self-contained file, zero frameworks:

```
app.py          # HTTP server (stdlib http.server) + all pipeline logic
test_sanity.py  # 68-assert sanity suite (unit + optional live-folder checks)
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
python3 test_sanity.py          # pure unit tests (68 asserts)
TEST_FOLDER_ID=<your_folder_id> python3 test_sanity.py   # + live end-to-end
```

## Operational limits

- **Folder must be shared as "Anyone with the link"** (public listing is required — no auth flows).
- Single attachment ≤ 200 MB; total job ≤ 1 GB; **global job storage quota 5 GB** (new jobs rejected with 507 when full); 4 concurrent jobs (HTTP 429 beyond).
- Subfolder recursion depth cap: 3 levels.
- Files auto-delete 30 minutes after creation (TTL sweep on every request + daemon).
- No AI, no analytics, no persistence: the server never stores anything beyond `/tmp` job output.

## Security notes

- **User input → parsed Drive folder ID only.** URLs are accepted only in the
  `drive.google.com/drive/folders/<id>` shape; every download URL is built
  from parsed file IDs and pinned to `drive.google.com`. No arbitrary URLs,
  no SSRF surface.
- Job IDs are UUID hex, strictly validated; the download route resolves
  `realpath` and rejects any path escaping the job directory.
- Downloads use explicit connect + read timeouts; per-job size caps are
  enforced while streaming (a slow/oversized remote cannot hold a worker
  indefinitely).
- Nothing user-supplied is ever executed or rendered (JSON only; filenames
  sanitized for zip).

## License

MIT — see [LICENSE](LICENSE).