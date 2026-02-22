"""
PDF Print Automation Service
"""
import os
import asyncio
import logging
import shutil
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

        logger.debug(f"[{job.id}] PHASE: Input validation passed. Path: {job.input_path}")

        # Create persistent output directory for this job
        job_output_dir = os.path.join(config.OUTPUT_DIR, job.id)
        os.makedirs(job_output_dir, exist_ok=True)

        # Generate output path directly in the output directory
        output_filename = f"{os.path.splitext(job.filename)[0]}_printed.pdf"
        output_path = os.path.join(job_output_dir, output_filename)

        logger.debug(f"[{job.id}] PHASE: Output target set. Path: {output_path}")

        # Execute print automation (raises on failure)
        logger.debug(f"[{job.id}] PHASE: Starting print automation")
        await self._print_single_pdf(job.input_path, output_path)

        # Verify output file exists and has content
        if not os.path.exists(output_path) or os.path.getsize(output_path) < 1000:
            raise FileValidationError(
                f"Output file missing or too small after automation: {output_path}"
            )

        logger.info(f"[{job.id}] Job completed. Output saved to: {output_path}")
        return output_path

    async def _print_single_pdf(self, input_path: str, output_path: str) -> None:
        """Execute Adobe Acrobat 9 Pro print automation for a single PDF.

        Raises:
            AcrobatWindowError: Acrobat window not found
            AutomationTimeoutError: Dialog or file save timed out
            UIAutomationError: General UI interaction failure
        """
        try:
            # --- PHASE 1: Open PDF in Acrobat ---
            logger.info(f"PHASE 1: Opening PDF: {os.path.basename(input_path)}")
            os.startfile(input_path)
            await asyncio.sleep(config.DOCUMENT_LOAD_WAIT)

            # --- PHASE 2: Find and activate Acrobat window ---
            logger.debug("PHASE 2: Searching for Acrobat window...")
            patterns = ["Adobe Acrobat 9 Pro", "Adobe Acrobat 9", "Adobe Acrobat", os.path.basename(input_path), "Adobe"]
            acrobat_window = None

            for pattern in patterns:
                windows = gw.getWindowsWithTitle(pattern)
                if windows:
                    acrobat_window = windows[0]
                    acrobat_window.activate()
                    await asyncio.sleep(1)
                    try:
                        acrobat_window.maximize()
                    except Exception:
                        pass
                    logger.info(f"PHASE 2: Acrobat window activated via pattern '{pattern}'")
                    break

            if not acrobat_window:
                raise AcrobatWindowError(
                    "Could not find Adobe Acrobat window. Is Acrobat installed and configured as the default PDF handler?"
                )

            # --- PHASE 3: Open Print dialog ---
            logger.info("PHASE 3: Opening Print dialog (Ctrl+P)...")
            pyautogui.hotkey('ctrl', 'p')
            await asyncio.sleep(config.PRINT_DIALOG_WAIT)

            # --- PHASE 4: Start printing ---
            logger.info("PHASE 4: Pressing Enter to start Adobe PDF printing...")
            pyautogui.press('enter')
            await asyncio.sleep(3)

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
                    logger.debug(f"PHASE 5: Save dialog appeared after {i+1}s")
                    break
                await asyncio.sleep(1)

            if not save_dialog_found:
                raise AutomationTimeoutError(
                    "Save As PDF dialog did not appear within 30 seconds. "
                    "Acrobat may have encountered an error during printing."
                )

            # --- PHASE 6: Activate Save As dialog ---
            logger.debug("PHASE 6: Activating Save As dialog...")
            save_dialog = None
            possible_titles = ["Save As", "Save PDF", "Save As PDF", "Save PDF File"]
            for title in possible_titles:
                dialogs = gw.getWindowsWithTitle(title)
                if dialogs:
                    save_dialog = dialogs[0]
                    break

            if not save_dialog:
                raise UIAutomationError(
                    "Save As dialog was detected but could not be activated."
                )

            save_dialog.activate()
            await asyncio.sleep(2)
            logger.info("PHASE 6: Save As dialog activated")

            # --- PHASE 7: Enter output path and save ---
            logger.info(f"PHASE 7: Entering output path: {output_path}")

            pyautogui.hotkey('ctrl', 'a')
            await asyncio.sleep(0.3)

            self._clipboard_set(output_path)
            pyautogui.hotkey('ctrl', 'v')
            logger.debug(f"PHASE 7: Pasted output path via clipboard")
            await asyncio.sleep(0.5)

            pyautogui.press('enter')
            logger.info("PHASE 7: Enter pressed to save PDF")

            # --- PHASE 8: Wait for file to be saved ---
            logger.info("PHASE 8: Waiting for file save...")

            def check_file_saved():
                if os.path.exists(output_path):
                    try:
                        size = os.path.getsize(output_path)
                        return size > 1000  # At least 1KB
                    except Exception:
                        return False
                return False

            file_saved = False
            for i in range(config.FILE_SAVE_WAIT):
                if check_file_saved():
                    file_saved = True
                    logger.debug(f"PHASE 8: File saved after {i+1}s")
                    break
                await asyncio.sleep(1)

            if not file_saved:
                raise AutomationTimeoutError(
                    f"File save timed out after {config.FILE_SAVE_WAIT} seconds. "
                    f"Output path: {output_path}"
                )

            logger.info(f"PHASE 8: File saved successfully: {os.path.basename(output_path)}")

            # --- PHASE 9: Close all Acrobat windows ---
            logger.info("PHASE 9: Closing all Acrobat windows...")
            await self._close_all_acrobat_windows()

            logger.info("Print automation completed successfully")

        except (AcrobatWindowError, AutomationTimeoutError, UIAutomationError):
            # Custom exceptions — re-raise as-is after attempting cleanup
            try:
                await self._safe_close_acrobat()
            except Exception:
                pass
            raise

        except Exception as e:
            # Unexpected errors — wrap in UIAutomationError
            logger.error(f"Print automation failed with unexpected error: {str(e)}")
            try:
                await self._safe_close_acrobat()
            except Exception:
                pass
            raise UIAutomationError(f"Unexpected automation failure: {str(e)}")

    async def _safe_close_acrobat(self):
        """Attempt to close Acrobat safely without affecting other windows"""
        try:
            fg_hwnd = win32gui.GetForegroundWindow()
            if self._is_acrobat_process(fg_hwnd):
                pyautogui.hotkey('alt', 'F4')
                await asyncio.sleep(0.5)
                pyautogui.press('n')
                logger.debug("Closed Acrobat window during error cleanup")
            else:
                logger.debug("Foreground window is not Acrobat, skipping Alt+F4")
        except Exception as close_err:
            logger.debug(f"Error during safe Acrobat close: {close_err}")

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
                    window.activate()
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
