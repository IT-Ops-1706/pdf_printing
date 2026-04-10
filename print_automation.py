"""
PDF Print Automation Service
"""
import os
import asyncio
import logging
import shutil
import subprocess
import ctypes
import time
import pyautogui
import pygetwindow as gw
import win32clipboard
import win32gui
import win32process
import win32api
import win32con

from config import config
from models import JobData
from errors import (
    AcrobatWindowError,
    AutomationTimeoutError,
    UIAutomationError,
    FileValidationError,
)

logger = logging.getLogger(__name__)


class PDFPrintAutomation:
    """Service for handling PDF print automation using UI automation"""

    def __init__(self):
        # Configure PyAutoGUI
        pyautogui.FAILSAFE = config.PYAUTOGUI_FAILSAFE
        pyautogui.PAUSE = config.PYAUTOGUI_PAUSE

    def _clipboard_set(self, text: str) -> None:
        """Set clipboard text using Win32 API (Windows Server compatible)"""
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()

    def _is_acrobat_process(self, hwnd: int) -> bool:
        """Verify a window handle belongs to an Adobe Acrobat process.
        Uses Win32 API to check the executable name, preventing
        accidental closure of unrelated windows.
        """
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            handle = win32api.OpenProcess(
                win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,
                False, pid
            )
            try:
                exe_path = win32process.GetModuleFileNameEx(handle, 0)
                exe_name = os.path.basename(exe_path).lower()
                is_acrobat = exe_name in ("acrobat.exe", "acrord32.exe")
                logger.debug(f"Window hwnd={hwnd} pid={pid} exe={exe_name} is_acrobat={is_acrobat}")
                return is_acrobat
            finally:
                win32api.CloseHandle(handle)
        except Exception as e:
            logger.debug(f"Could not verify process for hwnd={hwnd}: {e}")
            return False

    def _force_activate_window(self, hwnd: int) -> None:
        """Force a window to the foreground using Win32 AttachThreadInput trick.
        Bypasses Windows restrictions that prevent background processes
        from calling SetForegroundWindow.
        """
        try:
            # Restore if minimized
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)

            # Get thread IDs
            foreground_hwnd = win32gui.GetForegroundWindow()
            current_thread_id = ctypes.windll.kernel32.GetCurrentThreadId()
            target_thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]

            # Attach our thread input to the target thread to bypass foreground lock
            if current_thread_id != target_thread_id:
                ctypes.windll.user32.AttachThreadInput(current_thread_id, target_thread_id, True)
                try:
                    win32gui.SetForegroundWindow(hwnd)
                finally:
                    ctypes.windll.user32.AttachThreadInput(current_thread_id, target_thread_id, False)
            else:
                win32gui.SetForegroundWindow(hwnd)

            time.sleep(0.5)

            # Verify activation — if still not foreground, try BringWindowToTop
            if win32gui.GetForegroundWindow() != hwnd:
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.3)

            logger.debug(f"Force-activated window hwnd={hwnd}")
        except Exception as e:
            logger.warning(f"Force activate failed for hwnd={hwnd}: {e}")
            raise AcrobatWindowError(f"Could not activate window: {e}")

    def _force_kill_acrobat(self) -> None:
        """Force-kill all Acrobat processes using taskkill.
        More reliable than Alt+F4 which requires foreground access.
        """
        for exe in ("Acrobat.exe", "AcroRd32.exe"):
            try:
                result = subprocess.run(
                    ["taskkill", "/F", "/IM", exe],
                    capture_output=True, text=True, timeout=10
                )
                if result.returncode == 0:
                    logger.info(f"Force-killed all {exe} processes")
                else:
                    logger.debug(f"No {exe} processes to kill")
            except Exception as e:
                logger.debug(f"Error killing {exe}: {e}")
        time.sleep(2)  # Wait for processes to fully terminate

    async def process_pdf_job(self, job: JobData) -> str:
        """Process a PDF print job and save result to persistent storage.
        Reads directly from inputs/job_xxx/, writes directly to outputs/job_xxx/.

        Raises:
            FileValidationError: Input/output file validation failed
            AcrobatWindowError: Could not find or activate Acrobat
            AutomationTimeoutError: Save dialog or file save timed out
            UIAutomationError: General UI interaction failure
        """
        # Verify input file exists
        if not job.input_path or not os.path.exists(job.input_path):
            raise FileValidationError(f"Input PDF not found at: {job.input_path}")

        # --- DEBUG: Input file details ---
        try:
            input_size = os.path.getsize(job.input_path)
            logger.info(f"[{job.id}] Input validation passed. Path: {job.input_path}, Size: {input_size} bytes")
        except Exception as e:
            logger.warning(f"[{job.id}] Could not get input file size: {e}")

        # Create persistent output directory for this job
        job_output_dir = os.path.join(config.OUTPUT_DIR, job.id)
        os.makedirs(job_output_dir, exist_ok=True)

        # Generate output path directly in the output directory
        output_filename = f"{os.path.splitext(job.filename)[0]}_printed.pdf"
        output_path = os.path.join(job_output_dir, output_filename)

        # --- DEBUG: Verify output directory is writable ---
        try:
            test_file = os.path.join(job_output_dir, ".write_test")
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            logger.debug(f"[{job.id}] Output directory is writable: {job_output_dir}")
        except Exception as e:
            logger.error(f"[{job.id}] Output directory NOT writable: {job_output_dir} — {e}")
            raise FileValidationError(f"Output directory not writable: {job_output_dir} — {e}")

        # --- DEBUG: Check if output file already exists (could cause Replace dialog) ---
        if os.path.exists(output_path):
            existing_size = os.path.getsize(output_path)
            logger.warning(f"[{job.id}] Output file ALREADY EXISTS: {output_path}, size={existing_size} bytes. May trigger 'Replace?' dialog.")
        else:
            logger.debug(f"[{job.id}] Output target set (does not exist yet): {output_path}")

        # Execute print automation (raises on failure)
        logger.info(f"[{job.id}] Starting print automation...")
        await self._print_single_pdf(job.input_path, output_path)

        # Verify output file exists and has content
        if not os.path.exists(output_path):
            raise FileValidationError(f"Output file missing after automation: {output_path}")

        final_size = os.path.getsize(output_path)
        if final_size < 1000:
            raise FileValidationError(
                f"Output file too small after automation: {output_path} (size={final_size} bytes)"
            )

        logger.info(f"[{job.id}] Job completed. Output: {output_path}, Size: {final_size} bytes")
        return output_path

    async def _print_single_pdf(self, input_path: str, output_path: str) -> None:
        """Execute Adobe Acrobat 9 Pro print automation for a single PDF.

        Raises:
            AcrobatWindowError: Acrobat window not found
            AutomationTimeoutError: Dialog or file save timed out
            UIAutomationError: General UI interaction failure
        """
        try:
            # --- PRE-JOB: Environment diagnostics ---
            logger.info("="*60)
            logger.info("PRE-JOB: Starting environment diagnostics...")

            # Disk space check on output drive
            try:
                output_drive = os.path.splitdrive(output_path)[0] or 'C:'
                total, used, free = shutil.disk_usage(output_drive + '\\')
                free_mb = free // (1024 * 1024)
                logger.info(f"PRE-JOB: Disk space on {output_drive} — Free: {free_mb} MB, Total: {total // (1024*1024)} MB")
                if free_mb < 500:
                    logger.warning(f"PRE-JOB: LOW DISK SPACE on {output_drive}: only {free_mb} MB free!")
            except Exception as e:
                logger.warning(f"PRE-JOB: Could not check disk space: {e}")

            # Screen resolution and interactive session check
            try:
                screen_w, screen_h = pyautogui.size()
                logger.info(f"PRE-JOB: Screen resolution: {screen_w}x{screen_h}")
                if screen_w == 0 or screen_h == 0:
                    logger.error("PRE-JOB: Screen resolution is 0x0 — no interactive desktop session!")
            except Exception as e:
                logger.error(f"PRE-JOB: Could not get screen resolution (no interactive session?): {e}")

            # Current user and session info
            try:
                current_user = os.environ.get('USERNAME', 'unknown')
                session_name = os.environ.get('SESSIONNAME', 'unknown')
                logger.info(f"PRE-JOB: User: {current_user}, Session: {session_name}")
                if session_name.lower() in ('services', 'console'):
                    logger.warning(f"PRE-JOB: Running in '{session_name}' session — UI automation may not work!")
            except Exception as e:
                logger.debug(f"PRE-JOB: Could not get user/session info: {e}")

            # Check if Adobe PDF printer is installed
            try:
                result = subprocess.run(
                    ['wmic', 'printer', 'get', 'name'],
                    capture_output=True, text=True, timeout=10
                )
                printers = result.stdout.strip()
                has_adobe_pdf = 'adobe pdf' in printers.lower()
                logger.info(f"PRE-JOB: Adobe PDF printer installed: {has_adobe_pdf}")
                if not has_adobe_pdf:
                    logger.error(f"PRE-JOB: 'Adobe PDF' printer NOT found! Available printers:\n{printers}")
            except Exception as e:
                logger.debug(f"PRE-JOB: Could not check printers: {e}")

            # Log all currently visible windows
            try:
                all_wins_pre = gw.getAllWindows()
                visible_pre = [w.title for w in all_wins_pre if w.title.strip()]
                logger.info(f"PRE-JOB: Currently open windows ({len(visible_pre)}): {visible_pre[:15]}")
            except Exception as e:
                logger.debug(f"PRE-JOB: Could not list windows: {e}")

            logger.info("PRE-JOB: Environment diagnostics complete.")
            logger.info("="*60)

            # --- PRE-JOB CLEANUP: Kill any stale Acrobat instances ---
            logger.info("PRE-JOB: Cleaning up any stale Acrobat processes...")
            self._force_kill_acrobat()

            # --- PHASE 1: Open PDF in Acrobat ---
            logger.info(f"PHASE 1: Opening PDF: {os.path.basename(input_path)}")
            logger.debug(f"PHASE 1: Full input path: {input_path}")

            # Verify input file is not locked
            try:
                with open(input_path, 'rb') as f:
                    f.read(1)
                logger.debug(f"PHASE 1: Input file is readable (not locked)")
            except Exception as e:
                logger.error(f"PHASE 1: Input file may be LOCKED: {e}")

            os.startfile(input_path)
            logger.debug(f"PHASE 1: os.startfile() called. Waiting {config.DOCUMENT_LOAD_WAIT}s for Acrobat to load...")
            await asyncio.sleep(config.DOCUMENT_LOAD_WAIT)

            # --- DEBUG: Check what happened after DOCUMENT_LOAD_WAIT ---
            try:
                fg_hwnd_p1 = win32gui.GetForegroundWindow()
                fg_title_p1 = win32gui.GetWindowText(fg_hwnd_p1)
                logger.info(f"PHASE 1: After {config.DOCUMENT_LOAD_WAIT}s wait — Foreground: '{fg_title_p1}' (hwnd={fg_hwnd_p1})")
            except Exception as e:
                logger.debug(f"PHASE 1: Could not check foreground: {e}")

            # --- PHASE 2: Find and activate Acrobat window ---
            logger.info("PHASE 2: Searching for Acrobat window...")
            patterns = ["Adobe Acrobat 9 Pro", "Adobe Acrobat 9", "Adobe Acrobat", os.path.basename(input_path), "Adobe"]
            acrobat_window = None

            for pattern in patterns:
                windows = gw.getWindowsWithTitle(pattern)
                logger.debug(f"PHASE 2: Pattern '{pattern}' matched {len(windows)} window(s)")
                if windows:
                    acrobat_window = windows[0]
                    logger.debug(f"PHASE 2: Using window: '{acrobat_window.title}' (hwnd={acrobat_window._hWnd})")
                    self._force_activate_window(acrobat_window._hWnd)
                    await asyncio.sleep(1)
                    logger.info(f"PHASE 2: Acrobat window activated via pattern '{pattern}'")
                    break

            if not acrobat_window:
                # --- DEBUG: Dump ALL windows to help diagnose why Acrobat wasn't found ---
                try:
                    all_wins_p2 = gw.getAllWindows()
                    all_titles_p2 = [f"'{w.title}' (hwnd={w._hWnd})" for w in all_wins_p2 if w.title.strip()]
                    logger.error(f"PHASE 2: Acrobat NOT found. All {len(all_titles_p2)} windows: {all_titles_p2}")
                except Exception:
                    pass
                raise AcrobatWindowError(
                    "Could not find Adobe Acrobat window. Is Acrobat installed and configured as the default PDF handler?"
                )

            # --- PHASE 3: Open Print dialog ---
            logger.info("PHASE 3: Opening Print dialog (Ctrl+P)...")
            pyautogui.hotkey('ctrl', 'p')
            await asyncio.sleep(config.PRINT_DIALOG_WAIT)

            # --- DEBUG: Verify print dialog opened ---
            try:
                fg_hwnd_p3 = win32gui.GetForegroundWindow()
                fg_title_p3 = win32gui.GetWindowText(fg_hwnd_p3)
                logger.info(f"PHASE 3: After Ctrl+P wait — Foreground: '{fg_title_p3}' (hwnd={fg_hwnd_p3})")
                if 'print' not in fg_title_p3.lower() and fg_hwnd_p3 == acrobat_window._hWnd:
                    logger.warning("PHASE 3: Print dialog may NOT have opened — foreground is still the Acrobat main window")
            except Exception as e:
                logger.debug(f"PHASE 3: Could not verify print dialog: {e}")

            # --- PHASE 4: Start printing ---
            logger.info("PHASE 4: Pressing Enter to start Adobe PDF printing...")
            pyautogui.press('enter')
            await asyncio.sleep(3)

            # --- DEBUG: Check state after pressing Enter ---
            try:
                fg_hwnd_p4 = win32gui.GetForegroundWindow()
                fg_title_p4 = win32gui.GetWindowText(fg_hwnd_p4)
                logger.info(f"PHASE 4: After Enter — Foreground: '{fg_title_p4}' (hwnd={fg_hwnd_p4})")
            except Exception as e:
                logger.debug(f"PHASE 4: Could not check foreground: {e}")

            # --- PHASE 5: Wait for Save As dialog ---
            logger.info("PHASE 5: Waiting for Save As dialog...")
            await asyncio.sleep(5)  # Wait for progress bar to appear

            def check_save_dialog():
                possible_titles = ["Save As", "Save PDF", "Save As PDF", "Save PDF File"]
                for title in possible_titles:
                    dialogs = gw.getWindowsWithTitle(title)
                    if dialogs:
                        return True
                return False

            save_dialog_found = False
            for i in range(30):  # Wait up to 30 seconds
                if check_save_dialog():
                    save_dialog_found = True
                    logger.info(f"PHASE 5: Save dialog appeared after {i+1}s")
                    break

                # Log diagnostic info every 5 seconds to help debug failures
                if i % 5 == 0:
                    try:
                        fg_hwnd = win32gui.GetForegroundWindow()
                        fg_title = win32gui.GetWindowText(fg_hwnd)
                        logger.info(f"PHASE 5: [poll {i+1}s] Foreground window: '{fg_title}' (hwnd={fg_hwnd})")

                        # Log all visible windows with relevant keywords
                        all_windows = gw.getAllWindows()
                        relevant = [w for w in all_windows if w.title and any(
                            kw in w.title.lower() for kw in ["adobe", "acrobat", "save", "print", "pdf", "error", "warning"]
                        )]
                        for w in relevant:
                            logger.info(f"PHASE 5: [poll {i+1}s]   Visible: '{w.title}' (hwnd={w._hWnd}, visible={w.visible})")
                        if not relevant:
                            logger.info(f"PHASE 5: [poll {i+1}s]   No Adobe/Save/Print windows found")
                    except Exception as diag_err:
                        logger.debug(f"PHASE 5: Diagnostic logging error: {diag_err}")

                await asyncio.sleep(1)

            if not save_dialog_found:
                # Final diagnostic dump before raising error
                try:
                    fg_hwnd = win32gui.GetForegroundWindow()
                    fg_title = win32gui.GetWindowText(fg_hwnd)
                    all_windows = gw.getAllWindows()
                    all_titles = [w.title for w in all_windows if w.title.strip()]
                    logger.error(f"PHASE 5: TIMEOUT. Foreground: '{fg_title}'. All window titles: {all_titles}")
                except Exception:
                    pass
                raise AutomationTimeoutError(
                    "Save As PDF dialog did not appear within 30 seconds. "
                    "Acrobat may have encountered an error during printing."
                )

            # --- PHASE 6: Activate Save As dialog ---
            logger.info("PHASE 6: Activating Save As dialog...")
            save_dialog = None
            possible_titles = ["Save As", "Save PDF", "Save As PDF", "Save PDF File"]
            for title in possible_titles:
                dialogs = gw.getWindowsWithTitle(title)
                if dialogs:
                    save_dialog = dialogs[0]
                    logger.debug(f"PHASE 6: Found dialog with title '{title}' (hwnd={save_dialog._hWnd})")
                    break

            if not save_dialog:
                # Log all visible windows for diagnosis
                try:
                    all_wins = gw.getAllWindows()
                    all_titles = [w.title for w in all_wins if w.title.strip()]
                    logger.error(f"PHASE 6: No Save As dialog found. All window titles: {all_titles}")
                except Exception:
                    pass
                raise UIAutomationError(
                    "Save As dialog was detected but could not be activated."
                )

            self._force_activate_window(save_dialog._hWnd)
            await asyncio.sleep(2)

            # Verify Save As dialog is actually in foreground
            fg_hwnd = win32gui.GetForegroundWindow()
            fg_title = win32gui.GetWindowText(fg_hwnd)
            logger.info(f"PHASE 6: Save As dialog activated. Foreground window: '{fg_title}' (hwnd={fg_hwnd})")
            if fg_hwnd != save_dialog._hWnd:
                logger.warning(f"PHASE 6: Foreground hwnd ({fg_hwnd}) does not match Save dialog hwnd ({save_dialog._hWnd})")

            # --- PHASE 7 + 8: Enter output path, save, and verify (with retry) ---
            MAX_SAVE_RETRIES = 2
            SAVE_CHECK_SECONDS = 10  # seconds to wait per attempt before retrying

            def check_file_saved():
                if os.path.exists(output_path):
                    try:
                        size = os.path.getsize(output_path)
                        logger.debug(f"PHASE 8: File exists, size={size} bytes")
                        return size > 1000  # At least 1KB
                    except Exception as e:
                        logger.debug(f"PHASE 8: File exists but size check failed: {e}")
                        return False
                return False

            file_saved = False

            for attempt in range(1, MAX_SAVE_RETRIES + 1):
                logger.info(f"PHASE 7: Save attempt {attempt}/{MAX_SAVE_RETRIES}")

                # Step 1: Re-find and activate the Save As dialog (in case focus was lost)
                if attempt > 1:
                    logger.info(f"PHASE 7: Retry — re-finding Save As dialog...")
                    save_dialog = None
                    for title in possible_titles:
                        dialogs = gw.getWindowsWithTitle(title)
                        if dialogs:
                            save_dialog = dialogs[0]
                            logger.debug(f"PHASE 7: Re-found dialog '{title}' (hwnd={save_dialog._hWnd})")
                            break

                    if not save_dialog:
                        logger.warning(f"PHASE 7: Save As dialog no longer exists on retry {attempt}. Checking if file was already saved...")
                        if check_file_saved():
                            file_saved = True
                            logger.info(f"PHASE 8: File was saved (detected on retry). Size: {os.path.getsize(output_path)} bytes")
                            break
                        logger.error(f"PHASE 7: No Save As dialog and file not saved. Cannot retry.")
                        break

                    self._force_activate_window(save_dialog._hWnd)
                    await asyncio.sleep(1)

                # Step 2: Focus the filename input field using Alt+N
                logger.debug(f"PHASE 7: Pressing Alt+N to focus filename field...")
                pyautogui.hotkey('alt', 'n')
                await asyncio.sleep(0.5)

                # Log current foreground window to verify focus
                try:
                    fg_hwnd_now = win32gui.GetForegroundWindow()
                    fg_title_now = win32gui.GetWindowText(fg_hwnd_now)
                    logger.debug(f"PHASE 7: After Alt+N — Foreground: '{fg_title_now}' (hwnd={fg_hwnd_now})")
                except Exception as e:
                    logger.debug(f"PHASE 7: Could not check foreground after Alt+N: {e}")

                # Step 3: Select all text and paste the output path
                logger.debug(f"PHASE 7: Selecting all text in filename field (Ctrl+A)...")
                pyautogui.hotkey('ctrl', 'a')
                await asyncio.sleep(0.3)

                self._clipboard_set(output_path)
                pyautogui.hotkey('ctrl', 'v')
                logger.info(f"PHASE 7: Pasted output path via clipboard: {output_path}")
                await asyncio.sleep(0.5)

                # Step 4: Verify clipboard content matches what we set
                try:
                    win32clipboard.OpenClipboard()
                    clipboard_content = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                    if clipboard_content == output_path:
                        logger.debug(f"PHASE 7: Clipboard verification OK")
                    else:
                        logger.warning(f"PHASE 7: Clipboard mismatch! Expected: '{output_path}', Got: '{clipboard_content}'")
                except Exception as e:
                    logger.debug(f"PHASE 7: Clipboard verification failed: {e}")

                # Step 5: Press Enter to initiate save
                logger.info(f"PHASE 7: Pressing Enter to save PDF (attempt {attempt})...")
                pyautogui.press('enter')
                await asyncio.sleep(1)

                # Step 6: Handle possible "Replace file?" / "Confirm Save As" dialog
                try:
                    fg_hwnd_after = win32gui.GetForegroundWindow()
                    fg_title_after = win32gui.GetWindowText(fg_hwnd_after)
                    logger.debug(f"PHASE 7: After Enter — Foreground: '{fg_title_after}' (hwnd={fg_hwnd_after})")

                    confirm_keywords = ["confirm", "replace", "overwrite", "already exists", "save as"]
                    if any(kw in fg_title_after.lower() for kw in confirm_keywords):
                        logger.info(f"PHASE 7: Detected confirmation dialog: '{fg_title_after}'. Pressing Enter to confirm...")
                        pyautogui.press('enter')
                        await asyncio.sleep(1)
                except Exception as e:
                    logger.debug(f"PHASE 7: Error checking for confirmation dialog: {e}")

                # --- PHASE 8: Wait for file to be saved ---
                logger.info(f"PHASE 8: Waiting for file save (attempt {attempt}, timeout={SAVE_CHECK_SECONDS}s)...")

                for i in range(SAVE_CHECK_SECONDS):
                    if check_file_saved():
                        file_saved = True
                        logger.info(f"PHASE 8: File saved successfully after {i+1}s on attempt {attempt}. Size: {os.path.getsize(output_path)} bytes")
                        break

                    # Log diagnostic info every 3 seconds
                    if (i + 1) % 3 == 0:
                        try:
                            fg_hwnd_wait = win32gui.GetForegroundWindow()
                            fg_title_wait = win32gui.GetWindowText(fg_hwnd_wait)
                            file_exists = os.path.exists(output_path)
                            file_size = os.path.getsize(output_path) if file_exists else 0
                            logger.debug(
                                f"PHASE 8: [poll {i+1}s] file_exists={file_exists}, size={file_size}, "
                                f"foreground='{fg_title_wait}' (hwnd={fg_hwnd_wait})"
                            )
                        except Exception as diag_err:
                            logger.debug(f"PHASE 8: Diagnostic error: {diag_err}")

                    await asyncio.sleep(1)

                if file_saved:
                    break

                # File not saved on this attempt — log details before retry
                logger.warning(
                    f"PHASE 8: File NOT saved after {SAVE_CHECK_SECONDS}s on attempt {attempt}. "
                    f"file_exists={os.path.exists(output_path)}"
                )
                try:
                    fg_hwnd_fail = win32gui.GetForegroundWindow()
                    fg_title_fail = win32gui.GetWindowText(fg_hwnd_fail)
                    all_wins = gw.getAllWindows()
                    relevant = [w.title for w in all_wins if w.title and any(
                        kw in w.title.lower() for kw in ["adobe", "acrobat", "save", "print", "pdf", "error", "warning"]
                    )]
                    logger.warning(
                        f"PHASE 8: Foreground: '{fg_title_fail}'. Relevant windows: {relevant}"
                    )
                except Exception:
                    pass

            if not file_saved:
                total_wait = MAX_SAVE_RETRIES * SAVE_CHECK_SECONDS
                raise AutomationTimeoutError(
                    f"File save timed out after {MAX_SAVE_RETRIES} attempts ({total_wait}s total). "
                    f"Output path: {output_path}"
                )

            logger.info(f"PHASE 8: File saved successfully: {os.path.basename(output_path)}")

            # --- PHASE 9: Close all Acrobat windows ---
            logger.info("PHASE 9: Closing all Acrobat windows...")
            await self._close_all_acrobat_windows()

            logger.info("Print automation completed successfully")

        except (AcrobatWindowError, AutomationTimeoutError, UIAutomationError):
            # Custom exceptions — re-raise as-is after force-killing Acrobat
            try:
                self._force_kill_acrobat()
            except Exception:
                pass
            raise

        except Exception as e:
            # Unexpected errors — wrap in UIAutomationError
            logger.error(f"Print automation failed with unexpected error: {str(e)}")
            try:
                self._force_kill_acrobat()
            except Exception:
                pass
            raise UIAutomationError(f"Unexpected automation failure: {str(e)}")

    async def _safe_close_acrobat(self):
        """Legacy fallback — delegates to force kill"""
        self._force_kill_acrobat()

    async def _close_all_acrobat_windows(self):
        """Close all Adobe Acrobat windows, verified by process executable.
        Only closes windows confirmed to belong to Acrobat.exe/AcroRd32.exe.
        """
        try:
            logger.info("Finding and closing all verified Acrobat windows...")

            # Wait a moment for any new windows to fully open
            await asyncio.sleep(2)

            # Find all Acrobat windows by title pattern
            acrobat_patterns = ["Adobe Acrobat", "Acrobat"]
            all_acrobat_windows = []

            for pattern in acrobat_patterns:
                try:
                    windows = gw.getWindowsWithTitle(pattern)
                    all_acrobat_windows.extend(windows)
                except Exception as e:
                    logger.warning(f"Error finding windows with pattern '{pattern}': {e}")

            # Deduplicate by hwnd and verify each is actually Acrobat
            verified_windows = []
            seen_hwnds = set()
            for window in all_acrobat_windows:
                try:
                    hwnd = window._hWnd
                    if hwnd in seen_hwnds:
                        continue
                    seen_hwnds.add(hwnd)

                    if self._is_acrobat_process(hwnd):
                        verified_windows.append(window)
                    else:
                        logger.debug(f"Skipping non-Acrobat window: '{window.title}' (hwnd={hwnd})")
                except Exception as e:
                    logger.debug(f"Error checking window: {e}")
                    continue

            logger.info(f"Found {len(verified_windows)} verified Acrobat windows to close")

            # Close each verified window
            for i, window in enumerate(verified_windows):
                try:
                    logger.info(f"Closing Acrobat window {i+1}: '{window.title}'")

                    # Activate the window
                    self._force_activate_window(window._hWnd)
                    await asyncio.sleep(0.5)

                    # Send Alt+F4 to close
                    pyautogui.hotkey('alt', 'F4')
                    await asyncio.sleep(1)

                    # Handle save prompt if it appears
                    try:
                        fg_hwnd = win32gui.GetForegroundWindow()
                        if self._is_acrobat_process(fg_hwnd):
                            fg_title = win32gui.GetWindowText(fg_hwnd)
                            if "Adobe Acrobat" in fg_title:
                                pyautogui.press('n')
                                await asyncio.sleep(0.5)
                    except Exception as e:
                        logger.debug(f"Save prompt handling: {e}")

                    await asyncio.sleep(1)  # Wait for window to close

                except Exception as e:
                    logger.warning(f"Error closing Acrobat window '{window.title}': {e}")
                    # Try force close as fallback — only if still Acrobat
                    try:
                        if self._is_acrobat_process(window._hWnd):
                            window.close()
                            await asyncio.sleep(0.5)
                    except Exception as force_err:
                        logger.warning(f"Force close also failed: {force_err}")

            logger.info("Acrobat window closing completed")

        except Exception as e:
            logger.error(f"Error in _close_all_acrobat_windows: {e}")



    async def cleanup(self):
        """Cleanup resources used by the automation service"""
        logger.info("Cleaning up PDF print automation service...")
        # Add any additional cleanup logic here if needed
        logger.info("PDF print automation cleanup complete")
