# Daily Work Activity

A small web app that appends each day's work activities to **one permanent master Excel workbook**
(`GOKUL 26-27.xlsx`). Open the site, pick the date and shift, paste the activities (one per line),
and click **Submit Activities**. The app updates the same workbook every time. Nobody uploads files,
and no new workbook is created per submission.

```
Browser ──HTTPS──► Caddy ──► Next.js (UI + /api proxy) ──► FastAPI + openpyxl ──► /data/GOKUL 26-27.xlsx
                                                                                    (persistent volume)
```

## How the workbook is updated

Only the **Work activity** sheet is changed. `work 2026 -27` and `ESD Chiar retrofit` are kept as they are.

The sheet is a series of daily blocks: a header row, then 7 numbered rows, then one blank row. In each
block, Shift, Assigned CDL person, Work status, Shift DRI and Remarks are merged down the rows. For each
submission the backend:

1. finds the last row with data and leaves one blank row after it (it does **not** use "next empty
   physical row", so the pre-formatted empty block at the end of the sheet is reused and its empty merges
   are replaced cleanly);
2. writes a new block with the header, Sl.no 1..7 (more rows when there are more than 7 activities),
   the date on every row (`mm-dd-yy`, as in recent blocks), the shift, and these fixed values:
   - Assigned CDL person: the 8 names from the workbook, one per line
   - Work status: `Completed`
   - Shift DRI: `NAGENDRA.C      GOKUL.R  ` (exact spacing kept)
   - Remarks: `CLOSED`
3. copies every cell style (fonts, fills, borders, alignment, number formats) from the most recent
   well-formed block, so new blocks look the same as the hand-typed ones;
4. refuses to write over any cell that already has a value.

If the selected **date already has activities**, the server returns a warning. If the user confirms,
the activities go into the free rows of that date and shift's block. Any that don't fit go into a
follow-on block whose Sl.no numbering continues from the first block.

Pasted text is stored as text, never as a formula (for example `=SUM(...)`). Blank lines are ignored.
Leading and trailing spaces are trimmed. The wording is otherwise unchanged.

## Safety guarantees

| Concern | How it's handled |
|---|---|
| Concurrent submissions | Exclusive OS file lock (`flock`) around load → modify → save. Works across threads, uvicorn workers and containers sharing the volume. |
| Partial or corrupt writes | The new file is written to a temp file in the same folder, fsynced, **re-opened and checked** (the sheet exists, the activities are in the expected rows, the submission is recorded), then moved over the master with an atomic rename. If anything fails the old master stays, and the user sees *"The activities could not be saved…"*. |
| Double-click, retry, refresh | The browser sends a UUID with each submission and keeps it in local storage until the save is confirmed. Processed ids are stored as custom document properties **inside the workbook**, in the same atomic save. A retry of a request that was already saved returns success without adding rows again. |
| Accidental data loss | A timestamped copy of the previous version is kept in `/data/backups` before every save (the last 60 by default). |
| Lost volume | The seed workbook is copied only on the very first start. After that, a missing master is reported as an error and is never silently replaced by the old template. |
| Missing sheet | Reported. No replacement workbook is created. |

## Persistent storage choice

**Option A, a persistent server volume**, is used: a Docker named volume (`master-data`) mounted at
`/data`. Restarts, redeploys and image rebuilds keep it. This is the simplest reliable option, and it
works with any VM or VPS (or any host with a persistent disk, such as Render/Railway/Fly volumes mounted
at `/data`). Don't deploy the backend on platforms with only ephemeral disks (for example Vercel
functions or Heroku dynos) without attaching a persistent disk.

Back up the volume the same way as any other important file. The built-in `/data/backups` folder
protects against bad writes, not against losing the disk.

## Deploying (Docker Compose)

```bash
cp .env.example .env            # set SITE_ADDRESS to your domain (DNS must point at the server)
cp "/path/to/GOKUL 26-27.xlsx" seed/
docker compose up -d --build
```

Caddy gets and renews a TLS certificate automatically. Only Caddy's ports 80 and 443 are published.
The FastAPI backend and the workbook are reachable only from the internal Docker network, so the
browser can never write to the file directly. No credentials ship to the frontend.

To copy the live master off the server: use **Download Latest Excel** in the app, or run
`docker compose cp backend:/data/"GOKUL 26-27.xlsx" .`.

## Running locally

```bash
# backend
cd backend
pip install -r requirements-dev.txt
DATA_DIR=../data SEED_WORKBOOK_PATH="/path/to/GOKUL 26-27.xlsx" uvicorn app.main:app --port 8000

# frontend (another terminal)
cd frontend
npm ci
npm run dev        # http://localhost:3000, proxies /api to http://127.0.0.1:8000
```

## Configuration (backend environment variables)

| Variable | Default | Purpose |
|---|---|---|
| `DATA_DIR` | `/data` | Persistent folder holding the master, lock file, backups and audit log |
| `MASTER_WORKBOOK_NAME` | `GOKUL 26-27.xlsx` | File name of the master inside `DATA_DIR` |
| `SEED_WORKBOOK_PATH` | – | Workbook copied in on first start only |
| `CDL_PERSONS` | the 8 names in the workbook | `;`-separated override for the Assigned CDL person list |
| `SHIFT_DRI` | `NAGENDRA.C      GOKUL.R  ` | Shift DRI text |
| `BACKUPS_TO_KEEP` | `60` | Number of previous versions kept in `DATA_DIR/backups` |
| `LOCK_TIMEOUT_SECONDS` | `30` | How long a submission waits for another to finish |
| `MAX_ACTIVITIES` / `MAX_ACTIVITY_LENGTH` | `100` / `1000` | Input limits |

The frontend needs `BACKEND_URL` at **build** time (the compose file sets `http://backend:8000`).

## API

| Method | Path | |
|---|---|---|
| `POST` | `/api/submissions` | `{submission_id, date, shift: "A"\|"B", activities: string[], confirm_existing}` |
| `GET` | `/api/dates/{yyyy-mm-dd}` | Activities already recorded for a date, per shift |
| `GET` | `/api/workbook/status` | Whether the master is reachable, and its latest entry |
| `GET` | `/api/workbook/download` | Snapshot of the current master (read-only) |

## Tests

```bash
cd backend
python -m pytest -q                                       # synthetic workbook with the same layout
REAL_WORKBOOK="/path/to/GOKUL 26-27.xlsx" python -m pytest -q   # also runs every test on the real file
```

The tests cover single and multiple activities, next-day appends, A and B shifts, blocks with more than
7 activities, existing-date confirmation and overflow, unchanged history and other sheets (values,
styles, merges, column widths and row heights compared cell by cell), copied formatting, formula
injection, idempotent retries (including 8 concurrent identical requests), concurrent threads and
separate processes with no lost updates, survival across app restarts, missing master and missing sheet
handling, save failures leaving the master untouched, backups, input validation, and downloads.

> The workbook itself is **not** in this repository (it is git-ignored) because it contains internal
> company data and this repository is public.
