# BUSINESS LOGIC — PDF Print Utility

How the service behaves, end to end. This document describes **what happens to a PDF after it enters the API and why** — independent of the specific code paths. Read `print_automation.py` alongside Section 4 for the line-level reference.

---

## 1. Purpose

A modern HTTP service is required to "print" PDFs through Adobe Acrobat 9 Pro's **Adobe PDF virtual printer** — i.e. flatten / re-render the document via Acrobat's print pipeline so that downstream systems receive a re-stamped PDF that matches what a human user would have produced from `File → Print`.

There is no headless API for Acrobat 9 Pro. The only deterministic way to produce the same output is to drive the desktop UI. This service automates that interaction and exposes it as a queued REST API.

---

## 2. High-Level Flow

```
            ┌────────────────┐
 client ──▶ │ POST           │
            │ /print-queue   │
            └──────┬─────────┘
                   │ saves PDF to inputs/<job_id>/, enqueues JobData
                   ▼
            ┌────────────────────────┐
            │ asyncio.Queue (FIFO)   │  bounded by MAX_QUEUE_SIZE
            └──────┬─────────────────┘
                   │ background worker (single concurrency)
                   ▼
            ┌────────────────────────────────────┐
            │ PDFPrintAutomation.process_pdf_job │
            │   PHASE 1 → PHASE 9                │
            └──────┬─────────────────────────────┘
                   │ writes outputs/<job_id>/<name>_printed.pdf
                   ▼
            ┌────────────────┐
 client ──▶ │ GET /job-status│   reads file from disk on demand,
            │   /{job_id}    │   returns base64 result
            └────────────────┘

  (separate task) every CLEANUP_INTERVAL_SECONDS, jobs older than
  JOB_TTL_HOURS are purged (metadata + input + output folders).
```

---

## 3. Lifecycle of a Job

A job moves through exactly four states. Transitions are recorded in two places: the in-memory `JobData` dict and the daily `logs/.../states.json` file.

| State | Set By | Meaning |
|-------|--------|---------|
| `queued` | `add_job()` | PDF saved to disk, awaiting worker |
| `processing` | worker loop | Worker has popped the job from the queue |
| `completed` | worker loop on success | Output file exists & passes validation |
| `failed` | worker loop on exception | One of the typed errors below was raised |

### Failure taxonomy (`error_type`)

| `error_type` | Raised When | Recoverable? |
|--------------|-------------|--------------|
| `acrobat_not_found` | Acrobat window can't be located after the document-load wait | No — usually a config error: Acrobat not default handler, no display, or popup blocking |
| `timeout` | Save-As dialog or file write doesn't complete in budget | Sometimes — heavy machine load can cause it; retry the job |
| `ui_automation_failed` | Generic UI interaction failure (Save dialog appeared but couldn't be activated, etc.) | Sometimes — usually transient |
| `file_validation_failed` | Input missing, output dir not writable, or output PDF < 1 KB | No — fix the underlying file/disk issue |
| `unknown` | Any uncaught exception | Investigate logs |

Each terminal failure is followed by `_force_kill_acrobat()` so the next job starts from a clean slate.

---

## 4. The 9-Phase Print Automation

Each phase is logged with its name in `logs/<job_id>.log`. When a job fails, the last successful phase tells you exactly where to look.

### PHASE 0 — Pre-job diagnostics & cleanup
Before touching Acrobat, the automation logs disk space on the output drive, screen resolution, current Windows session (must be interactive — `Console`/`Services` is flagged as a warning), Adobe PDF printer presence, and currently visible windows. Then it runs `taskkill /F /IM Acrobat.exe AcroRd32.exe` to remove any zombie instance from a previous job.

### PHASE 1 — Open the PDF
`os.startfile(input_path)` hands the file to whatever Windows declares is the default `.pdf` handler. This MUST be **Adobe Acrobat 9 Pro**; if it isn't, PHASE 2 fails. The script then sleeps `DOCUMENT_LOAD_WAIT` seconds (default 20) to let Acrobat initialize, render, and become responsive.

### PHASE 2 — Locate & activate the Acrobat window
Searches `pygetwindow` for windows matching, in order:
1. `Adobe Acrobat 9 Pro`
2. `Adobe Acrobat 9`
3. `Adobe Acrobat`
4. The input PDF's filename
5. `Adobe`

The first match is force-activated using a Win32 `AttachThreadInput` trick that defeats Windows' foreground-lock policy (necessary because the API runs as a background process). Failure here raises `AcrobatWindowError`.

### PHASE 3 — Open the Print dialog
`Ctrl+P`, wait `PRINT_DIALOG_WAIT` (10 s). The script logs the foreground window after the wait — if it's still the Acrobat main window, the keystroke didn't take and the job will fail in PHASE 5.

### PHASE 4 — Submit the print job
Press `Enter` to accept the dialog. This relies on **Adobe PDF being the selected/default printer**. Acrobat now begins rendering through the Adobe PDF printer driver, which will eventually pop a Save As dialog asking where to write the resulting PDF.

### PHASE 5 — Wait for Save As dialog (poll)
Polls every second for up to 30 s for any window titled `Save As`, `Save PDF`, `Save As PDF`, or `Save PDF File`. Every 5 s, dumps the foreground title and any "interesting" visible window to the log. Timeout raises `AutomationTimeoutError` with a complete window snapshot — invaluable for debugging.

### PHASE 6 — Activate Save As dialog
Same `_force_activate_window` trick as PHASE 2. Verifies the dialog actually became foreground.

### PHASE 7 — Type the output path & save
1. `Alt+N` → focus the filename field.
2. `Ctrl+A` → select existing path.
3. Set the OS clipboard to the desired output path (using Win32 clipboard APIs — no PowerShell, no `pyperclip` — Windows-Server-safe).
4. `Ctrl+V` → paste.
5. Verify clipboard content matches what was intended (defensive against clipboard managers).
6. `Enter` → submit.
7. If a `Confirm Save As` / `Replace?` dialog appears, press `Enter` again to accept.

### PHASE 8 — Wait for file to appear on disk
Polls for up to `SAVE_CHECK_SECONDS` (10 s) per attempt, twice (`MAX_SAVE_RETRIES = 2`), checking that:

- `output_path` exists.
- File size > 1 KB (rules out a half-written file).

Between attempts, PHASE 7 is repeated **from step 1** — re-find the dialog, re-paste, re-Enter. This robustness is critical: Acrobat 9 occasionally drops focus to the desktop after pressing Enter, especially on slower hardware.

### PHASE 9 — Close all Acrobat windows
Find every `Adobe Acrobat`/`Acrobat` window, **verify each is owned by `Acrobat.exe`/`AcroRd32.exe`** (preventing accidental closure of a different app that happens to mention "Acrobat"), and send `Alt+F4`. If Acrobat shows a "save changes?" prompt during close, send `n`. As a last resort, the automation `taskkill`s any survivor.

---

## 5. Concurrency Model

The system is intentionally **single-concurrency**:

- `MAX_PRINT_WORKERS = 1`.
- A single `process_queue` task pops jobs one at a time.
- Acrobat 9 Pro cannot reliably handle two concurrent UI sessions on the same desktop; queueing serializes them.

Throughput is therefore bounded by `DOCUMENT_LOAD_WAIT + PRINT_DIALOG_WAIT + SAVE_DIALOG_WAIT + per-PDF render time`. On typical hardware: **30–60 seconds per job**.

`MAX_QUEUE_SIZE = 50` protects against unbounded memory growth — additional `POST /print-queue` calls beyond this return HTTP `429`.

---

## 6. Storage Model

```
inputs/
  job_3_a1b2c3d4/
    input_invoice.pdf            ← raw upload, base64-decoded
outputs/
  job_3_a1b2c3d4/
    invoice_printed.pdf          ← Acrobat-rendered output
logs/
  2026/05/05/
    job_3_a1b2c3d4.log           ← per-job phase log (DEBUG)
    states.json                  ← daily roll-up of all jobs (status, durations, errors)
```

- **No base64 in memory after upload.** The API decodes once, writes to `inputs/`, and tracks only the disk path on `JobData`. This matters for large PDFs and high job counts.
- **Output is read on-demand at status time.** `GET /job-status/{job_id}` reads the output file from disk, base64-encodes, and returns. The result is **not** cached server-side.
- **TTL cleanup.** Every `CLEANUP_INTERVAL_SECONDS` (30 min), jobs older than `JOB_TTL_HOURS` (3 h) are purged — folder + metadata. On startup, an additional sweep removes orphaned folders from a previous process.

---

## 7. API Contract

### `POST /print-queue` — Bearer-protected
- **Body**: `multipart/form-data`, single `file` field with the PDF.
- **Success**: `200` `{ job_id, filename, message, status: "queued" }`.
- **Errors**:
  - `401` invalid/missing API key.
  - `429` `MAX_QUEUE_SIZE` reached.
  - `500` save-to-disk failure or other server error.

### `GET /job-status/{job_id}` — Bearer-protected
- **Success**: `200` with the full `JobData` (status, error fields, timestamps, and — when `status == completed` — `result` containing base64 of the printed PDF).
- **Errors**: `401` invalid token, `404` unknown `job_id`.

### `GET /health` — public
Returns `{ status, queue_size, queue_limit, queue_full, processing, current_job }`. Used by external monitors.

---

## 8. Why the Architecture Looks This Way

| Decision | Reason |
|----------|--------|
| Sequential queue, one worker | Acrobat 9 Pro can't be driven concurrently on a single desktop; serializing avoids stuck dialogs. |
| Disk-backed inputs, not in-memory base64 | Memory bound regardless of job size or queue depth. |
| `JOB_TTL_HOURS = 3` | Long enough for client polling cycles, short enough to bound disk usage. |
| Force-kill on every error | UI automation failures often leave Acrobat in a bad state; restart is faster and safer than recovery. |
| Process-verified window closes | "Adobe Acrobat" can appear in unrelated window titles; verifying via `OpenProcess + GetModuleFileNameEx` prevents collateral damage. |
| Win32 clipboard, not `pyperclip` | `pyperclip` shells out to PowerShell on Windows; on Server SKUs that can fail or be slow. Direct Win32 calls are deterministic. |
| `AttachThreadInput` activation | A background process cannot normally call `SetForegroundWindow`; this is the documented workaround. |

---

## 9. Operational Invariants

These must hold for the service to function. Violations are logged at `WARNING` or `ERROR` and almost always lead to job failures.

1. The default `.pdf` handler is **Adobe Acrobat 9 Pro**.
2. Adobe PDF printer is installed and (ideally) the default printer.
3. The host is in an interactive desktop session, screen unlocked.
4. No Acrobat popups (welcome, EULA, updater, online-storage) are pending.
5. The `outputs/` directory is writable and has > 500 MB free.
6. No other process is moving the mouse / sending keystrokes during a job.
7. The PyAutoGUI failsafe corner is not hit (mouse to top-left aborts).
