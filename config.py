"""
Configuration management for PDF Print Utility
"""
import os

class Config:
    """Centralized configuration for the PDF Print Utility"""

    # API Configuration
    API_KEY: str = "BabajiShivram@1706"
    HOST: str = "127.0.0.1"
    PORT: int = 8001

    # Print Configuration
    MAX_PRINT_WORKERS: int = 1  # Sequential processing
    PRINT_TIMEOUT: int = 120    # seconds

    # Timing constants (seconds) - Adjusted for Adobe Acrobat 9 Pro
    DOCUMENT_LOAD_WAIT: int = 7
    PRINT_DIALOG_WAIT: int = 4
    SAVE_DIALOG_WAIT: int = 4
    FILE_SAVE_WAIT: int = 20

    # Queue Configuration
    QUEUE_CHECK_INTERVAL: int = 1  # seconds
    MAX_QUEUE_SIZE: int = 100      # Maximum jobs allowed in queue
    INPUT_DIR: str = os.path.join(os.getcwd(), "inputs")   # Input PDFs storage
    OUTPUT_DIR: str = os.path.join(os.getcwd(), "outputs")
    
    # Job Cleanup Configuration
    JOB_TTL_HOURS: int = 3         # Time-to-live for jobs in hours (3 hours)
    CLEANUP_INTERVAL_SECONDS: int = 1800  # Run cleanup every 30 minutes


    # PyAutoGUI settings
    PYAUTOGUI_FAILSAFE: bool = True
    PYAUTOGUI_PAUSE: float = 0.8

    @classmethod
    def create_dirs(cls) -> None:
        """Ensure necessary directories exist"""
        os.makedirs(cls.INPUT_DIR, exist_ok=True)
        os.makedirs(cls.OUTPUT_DIR, exist_ok=True)

# Initialize configuration
config = Config()
config.create_dirs()
