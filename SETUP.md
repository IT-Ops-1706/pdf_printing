# SETUP — PDF Print Utility

End-to-end setup for the PDF Print API on a Windows machine. Follow every step in order on a **fresh deployment**; the UI automation is sensitive to the host environment, so the prerequisites are not optional.

---

## 1. System Prerequisites

| Component | Required Version / State |
|-----------|--------------------------|
| OS | Windows 10 / 11 / Server (interactive desktop session — **not** a "Services" session) |
| Python | 3.9 or higher (64-bit) |
| Adobe Acrobat | **Adobe Acrobat 9 Pro** (full version, not Reader) |
| Adobe PDF Printer | Installed automatically with Acrobat 9 Pro — verify in `Control Panel → Devices and Printers` |
| Screen | A real / virtual display must be active — RDP session is OK, but a locked screen is **NOT** |

> The automation script types into the active desktop. If the machine is locked, has no monitor, or runs the script as a Windows Service in session 0, every job will fail at PHASE 2 (`AcrobatWindowError`).

---

## 2. Set Adobe Acrobat 9 Pro as the Default PDF Handler

This is the **single most important configuration step**. The automation calls `os.startfile(input_path)`; Windows then routes the file to whichever app owns the `.pdf` extension. If anything other than Acrobat 9 Pro opens (Edge, Chrome, Reader DC, Foxit, etc.) the script will not find the expected window and will fail.

### Steps

1. Open **Settings → Apps → Default Apps** (Windows 10/11).
2. Scroll to **Choose defaults by file type**.
3. Find `.pdf` in the list.
4. Click the current handler and choose **Adobe Acrobat 9 Pro** from the list.
5. Repeat for `.fdf` and `.xfdf` if present (optional but recommended).

### Verify

```powershell
# Should print Acrobat.exe (NOT AcroRd32.exe, NOT msedge.exe, NOT chrome.exe)
cmd /c assoc .pdf
cmd /c ftype Acrobat.Document.9
```

You should see something resembling:

```
Acrobat.Document.9="C:\Program Files (x86)\Adobe\Acrobat 9.0\Acrobat\Acrobat.exe" "%1"
```

If a different handler is shown, repeat step 4 — Windows occasionally restores Edge as the default after updates.

---

## 3. Disable All Adobe Acrobat Startup Popups & Dialogs

Acrobat 9 Pro shows several modal popups that **block UI automation** if not pre-dismissed: the Welcome Screen, EULA acceptance, "Updater" prompt, "Send usage data" prompt, and the protected-mode warning. These must be turned off **before** the API is started.

### One-time manual launch

1. Open Adobe Acrobat 9 Pro **manually** once on the machine that will run the API.
2. Accept the EULA if prompted (this only needs to happen once per Windows user).
3. Inside Acrobat go to **Edit → Preferences** and apply the following:

| Preferences Page | Setting | Value |
|------------------|---------|-------|
| General | Show splash screen at startup | **Unchecked** |
| General | Show messages at launch | **Unchecked** |
| Updater | How would you like to install updates | **Do not download or install updates automatically** |
| Documents | Show online storage when opening files | **Unchecked** |
| Documents | Show online storage when saving files | **Unchecked** |
| Security (Enhanced) | Enable Protected View | **Off** |
| Security (Enhanced) | Enable Enhanced Security | **Off** |
| Trust Manager | Allow opening of non-PDF file attachments… | **Unchecked** |
| JavaScript | Enable Acrobat JavaScript | **Unchecked** *(prevents form-script popups)* |

4. Close Acrobat. Re-open it once more — confirm no popup, no welcome dialog, and no update prompt appears. The **Document Window** should be the only thing on screen.

### Disable "Save in Adobe Cloud" prompt (Save As dialog)

In **Edit → Preferences → General**, ensure **"Show online storage when saving files"** is unchecked. If it is checked, the `Save As` dialog opens to the Adobe Cloud picker first and the path-paste step (PHASE 7) will land in the wrong field.

---

## 4. Set "Adobe PDF" as the Default Printer (recommended)

The automation presses **Enter** in the print dialog to accept the currently selected printer — which works only if Adobe PDF is the default. To make this deterministic:

1. `Control Panel → Devices and Printers`.
2. Right-click **Adobe PDF** → **Set as default printer**.
3. Open Acrobat once, press `Ctrl+P`, and confirm the dropdown shows **Adobe PDF**. Cancel.

If a different printer must remain default for another workflow, edit `print_automation.py` PHASE 4 to navigate the printer dropdown explicitly — but the simpler operational stance is to make Adobe PDF the default on this machine.

---

## 5. Disable "Adobe PDF Printer — Save As" Behavior Quirks

The Adobe PDF printer's Save As dialog can show two extra confirmations that the script tolerates but does not require:

- **Replace existing file?** — handled in `print_automation.py` PHASE 7 by re-pressing Enter.
- **Confirm Save As** — same handler.

No setup action required, but if you change the Adobe PDF printer's defaults via `Devices and Printers → Adobe PDF → Printing Preferences`:

- **Adobe PDF Settings tab**: uncheck **View Adobe PDF results** so Acrobat doesn't reopen the just-saved file in a new window.
- **Adobe PDF Settings tab**: uncheck **Ask to Replace existing PDF file** if you want fully unattended overwrite.
- **Adobe PDF Settings tab**: uncheck **Add document information**.

---

## 6. Project Installation

```powershell
# 1. Clone (or pull) the repository
git clone https://github.com/IT-Ops-1706/pdf_printing.git
cd pdf_printing

# 2. Create and activate a virtual environment
python -m venv venv
.\venv\Scripts\Activate.ps1

# 3. Install dependencies
pip install --upgrade pip
pip install -r requirements.txt
```

If PowerShell blocks the activation script:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

---

## 7. Configuration

All runtime settings live in `config.py`. The defaults are production-sane; review before first run:

| Setting | Default | Description |
|---------|---------|-------------|
| `API_KEY` | `BabajiShivram@1706` | Bearer token required on `/print-queue` and `/job-status/*` |
| `HOST` | `127.0.0.1` | Bind address — change to `0.0.0.0` for LAN exposure |
| `PORT` | `8001` | API port |
| `MAX_QUEUE_SIZE` | `50` | Pending-job ceiling; `429` returned when full |
| `JOB_TTL_HOURS` | `3` | Auto-cleanup of input/output folders & job metadata |
| `DOCUMENT_LOAD_WAIT` | `20` s | Time waited for Acrobat to render the PDF |
| `PRINT_DIALOG_WAIT` | `10` s | Time waited for Ctrl+P dialog |
| `SAVE_DIALOG_WAIT` | `10` s | Time waited for Save As dialog (before polling) |
| `PYAUTOGUI_FAILSAFE` | `True` | Mouse to a screen corner aborts automation |
| `PYAUTOGUI_PAUSE` | `0.8` s | Delay between every keystroke |

Treat `API_KEY` as a secret — rotate it before going to production and store it in an environment variable rather than the source file.

---

## 8. Run the Service

```powershell
# Foreground
python main.py
```

You should see:

```
... | INFO     | __main__ | Starting PDF Print API...
... | INFO     | __main__ | Queue processor and cleanup task started
INFO:     Uvicorn running on http://127.0.0.1:8001
```

Visit `http://127.0.0.1:8001/docs` for interactive Swagger UI.

### Smoke test

```powershell
$headers = @{ "Authorization" = "Bearer BabajiShivram@1706" }

# Health
Invoke-RestMethod -Uri "http://127.0.0.1:8001/health" -Headers $headers

# Queue a job
$form = @{ file = Get-Item "C:\path\to\sample.pdf" }
Invoke-RestMethod -Uri "http://127.0.0.1:8001/print-queue" `
                  -Method Post -Headers $headers -Form $form

# Poll status
Invoke-RestMethod -Uri "http://127.0.0.1:8001/job-status/job_1_xxxxxxxx" -Headers $headers
```

---

## 9. Run as a Background Service (optional)

The automation **must** run inside an interactive desktop session. Do **not** install it as a classic Windows Service in session 0 — UI automation will fail. Recommended options:

- **Task Scheduler** with **"Run only when user is logged on"** + **"Start in"** = project root.
- **NSSM** with the `AppExit Default Restart` action and the user account auto-logged on.
- A pinned RDP session that stays open on the build server.

---

## 10. First-Run Checklist

Before declaring the deployment "done", confirm in order:

- [ ] Double-clicking any `.pdf` opens **Acrobat 9 Pro** (not Reader, Edge, or Chrome).
- [ ] Launching Acrobat manually shows **no popups** of any kind.
- [ ] `Adobe PDF` appears in `Devices and Printers` and is the default printer.
- [ ] `python main.py` starts without errors and `/health` returns `200`.
- [ ] A test PDF queued via `/print-queue` produces a file under `outputs/<job_id>/`.
- [ ] `logs/<year>/<month>/<day>/states.json` records the job and its `completed` status.

If any item fails, fix it before queuing real traffic — once a popup appears mid-job, the automation will hang on that machine until the dialog is closed manually.
