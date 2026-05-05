# HANDOVER — PDF Print Utility

A condensed brief for the next engineer / operator taking over the service. If you read only one document in this repo, read this one.

---

## TL;DR

- FastAPI service that receives PDFs and "prints" them through **Adobe Acrobat 9 Pro** via UI automation (`pyautogui` + Win32).
- Single worker, queued jobs, on-disk storage, Bearer-token auth.
- Lives at `127.0.0.1:8001`, repo `https://github.com/IT-Ops-1706/pdf_printing.git`.
- Runs only on Windows, only on a logged-in interactive desktop, only when Acrobat 9 Pro is the default PDF handler with all popups disabled.

For deep dives, see:
- `SETUP.md` — install + environment hardening.
- `BUSINESS_LOGIC.md` — request flow, phases, error taxonomy.
- `README.md` — original short summary.

---

## 1. Recent Changes Operators Must Know

These are the configuration / deployment changes baked into the current production stance. They are environment-level, not code-level.

| Change | Why | Where it lives |
|--------|-----|----------------|
| **Adobe Acrobat 9 Pro is now the default PDF opener.** Any other handler (Edge, Chrome, Reader DC, Foxit) will break PHASE 2. | `os.startfile()` in `print_automation.py` relies on Windows' file association. | OS-level: `Settings → Default Apps → .pdf`. See `SETUP.md` §2. |
| **All Acrobat startup popups are disabled by default.** Welcome screen, EULA, Updater, online-storage prompts, JavaScript alerts, Protected View — all turned off. | Any popup steals focus and the automation hangs at the offending phase until manually closed. | Acrobat preferences (per-user). See `SETUP.md` §3. |
| **Adobe PDF set as the default Windows printer.** | PHASE 4 presses Enter to accept the selected printer; the default printer must be Adobe PDF. | OS-level: `Devices and Printers`. See `SETUP.md` §4. |
| **`Show online storage when saving files` is unchecked.** | Otherwise the Save As dialog opens to the Adobe Cloud picker and the path-paste step lands in the wrong field. | Acrobat → Edit → Preferences → General. |
| **`View Adobe PDF results` (printer pref) is unchecked.** | Stops Acrobat from re-opening the freshly-saved PDF in a new window mid-job. | `Devices and Printers → Adobe PDF → Printing Preferences`. |

If a fresh box is being commissioned, work through `SETUP.md` §1–§5 before doing anything else. Skipping any of those steps will produce a service that "almost works" — first job succeeds, subsequent jobs hang.

---

## 2. What's Where

```
pdf_printing/
├── main.py                  FastAPI app + lifespan + routes
├── config.py                All tunables (API key, timeouts, paths, queue size)
├── models.py                Pydantic models (JobData is the canonical record)
├── queue_manager.py         asyncio.Queue + worker loop + cleanup tasks
├── print_automation.py      The 9-phase Acrobat driver (the hard part)
├── errors.py                Typed exceptions → error_type in API
├── logging_setup.py         Pipe-delimited terminal logger + per-job FileHandler
├── job_state_logger.py      Atomic daily states.json roll-up
├── requirements.txt
├── README.md                Short overview
├── SETUP.md                 Full install + env hardening
├── BUSINESS_LOGIC.md        Per-phase flow, failure taxonomy
└── HANDOVER.md              You are here
```

Runtime-only directories (gitignored, created on demand): `inputs/`, `outputs/`, `logs/`, `venv/`.

---

## 3. Day-to-Day Operations

### Start the service
```powershell
.\venv\Scripts\Activate.ps1
python main.py
```
Bind: `127.0.0.1:8001`. Change `HOST`/`PORT` in `config.py` for LAN exposure.

### Auth
Header: `Authorization: Bearer BabajiShivram@1706` (rotate before going to prod).

### Queue a job
```powershell
$h = @{ Authorization = "Bearer BabajiShivram@1706" }
Invoke-RestMethod -Uri "http://127.0.0.1:8001/print-queue" `
                  -Method Post -Headers $h `
                  -Form @{ file = Get-Item "C:\samples\test.pdf" }
```

### Check a job
```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:8001/job-status/job_1_xxxxxxxx" -Headers $h
```

### Health
```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:8001/health"
```

---

## 4. Where the Logs Are

```
logs/<YYYY>/<MM>/<DD>/
    job_<id>.log     — per-job DEBUG log, every phase, every dialog poll
    states.json      — daily roll-up: counts, per-job timestamps, durations, errors
```

When a job fails, open `job_<id>.log` and search for the last `PHASE` line — that's where the failure happened. The log dumps all visible windows on every relevant timeout, so most diagnoses are possible without reproducing.

`states.json` is the single observability surface for ops dashboards — it's atomic-written and safe to tail.

---

## 5. Common Failures & Fixes

| Symptom | Likely Cause | Fix |
|---------|--------------|-----|
| All jobs fail at PHASE 2 with `acrobat_not_found` | Edge/Reader/Chrome stole the `.pdf` association after Windows update | `Settings → Default Apps → .pdf` → re-set to Acrobat 9 Pro |
| All jobs fail at PHASE 5 (no Save As dialog) | Adobe PDF is not the default printer, **or** the print dialog's previous selection is still a real printer | Set Adobe PDF as default, or reopen Acrobat once and choose Adobe PDF in the print dropdown so the selection is remembered |
| Jobs hang at PHASE 5; logs show a window like "Adobe Acrobat — Welcome" | A popup re-appeared (often after an Acrobat update) | Open Acrobat manually, dismiss + uncheck "show again". Re-disable Updater preference. |
| First job after boot fails, subsequent jobs work | Acrobat cold-start exceeds `DOCUMENT_LOAD_WAIT` | Increase `DOCUMENT_LOAD_WAIT` to 30, or open Acrobat once before starting the API |
| Jobs sometimes fail at PHASE 8 with `timeout` | Disk slow, output file written in chunks > 10 s | Bump `SAVE_CHECK_SECONDS` in `print_automation.py` (currently 10) |
| `file_validation_failed` with output < 1 KB | Acrobat saved an empty/corrupt file (almost always after a popup intervened mid-save) | Check the popup; the size threshold is the safety net catching the corruption |
| API returns 429 | `MAX_QUEUE_SIZE` reached | Wait, or increase the limit + scale up the host |
| RDP disconnect → all jobs fail | Disconnecting RDP locks the screen → no interactive desktop | Use `tscon` to leave the session active, or run from a console session |

---

## 6. Things You Should NOT Do

- **Do not** run this as a session-0 Windows Service. UI automation requires an interactive desktop.
- **Do not** install Acrobat Reader DC alongside Acrobat 9 Pro and let it become the default — Acrobat Reader uses different window titles and the script will misroute.
- **Do not** raise `MAX_PRINT_WORKERS` above 1. Two concurrent automations on the same desktop fight for the foreground and corrupt each other.
- **Do not** click anything on the host while a job is processing. The script controls the keyboard and mouse for ~30–60 s; user input during that window will derail the job.
- **Do not** rely on `result` being cached. `GET /job-status` reads from disk on each call; if the TTL cleanup ran, the file is gone and the response will flip to `failed`.
- **Do not** log the API key in client code or config-management diffs. Rotate it via `config.py` and restart.

---

## 7. Repo & Branching

- Remote: `origin = https://github.com/IT-Ops-1706/pdf_printing.git`
- Default branch: `main` — direct commits land here today; introduce PR review before opening up access.
- `.gitignore` excludes `inputs/`, `outputs/`, `logs/`, `venv/`, IDE folders. Don't add them.

---

## 8. Contact / Ownership

This service was built for an internal Babaji Shivram automation pipeline. Code under: `Backend/Python Projects/pdf_merge_utility/pdf_printing/`. When in doubt, read the per-job log for the failing `job_id` first — it captures more context than any external monitoring tool.
