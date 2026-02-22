"""
Custom exceptions for PDF Print Automation.
Each exception type maps to a distinct error_type in the API response,
enabling downstream systems to handle failures programmatically.
"""


class AutomationBaseError(Exception):
    """Base class for all automation errors"""
    error_type: str = "unknown"

    def __init__(self, message: str):
        self.message = message
        super().__init__(message)


class AcrobatWindowError(AutomationBaseError):
    """Raised when Adobe Acrobat window cannot be found or activated"""
    error_type = "acrobat_not_found"


class AutomationTimeoutError(AutomationBaseError):
    """Raised when a UI automation phase exceeds its timeout"""
    error_type = "timeout"


class UIAutomationError(AutomationBaseError):
    """Raised when a UI automation interaction fails (dialog, keystroke, etc.)"""
    error_type = "ui_automation_failed"


class FileLockError(AutomationBaseError):
    """Raised when the output file remains locked after max retries"""
    error_type = "file_locked"


class FileValidationError(AutomationBaseError):
    """Raised when input/output file validation fails"""
    error_type = "file_validation_failed"
