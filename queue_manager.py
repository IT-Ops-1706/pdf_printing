"""
Queue management system for PDF print jobs
"""
import asyncio
import logging
import uuid
import os
from datetime import datetime
from typing import Dict, Optional, Any
from concurrent.futures import ThreadPoolExecutor

from config import config
from models import JobData
from print_automation import PDFPrintAutomation
from errors import AutomationBaseError
from logging_setup import get_job_file_handler
import job_state_logger

logger = logging.getLogger(__name__)

# Base log directory — sub-directories are created per job by get_job_file_handler
BASE_LOG_DIR = "logs"


class PDFPrintQueue:
    """Enhanced queue system for managing PDF print jobs sequentially"""

    def __init__(self):
        self.queue: asyncio.Queue = asyncio.Queue(maxsize=config.MAX_QUEUE_SIZE)
        self.processing: bool = False
        self.current_job: Optional[JobData] = None
        self.jobs: Dict[str, JobData] = {}  # Track all jobs
        self.executor: ThreadPoolExecutor = ThreadPoolExecutor(max_workers=config.MAX_PRINT_WORKERS)
        self.automation = PDFPrintAutomation()
        self.job_loggers: Dict[str, logging.Logger] = {}  # Individual job loggers

    def _get_job_logger(self, job_id: str) -> logging.Logger:
        """Get or create a file logger for a specific job.

        Writes to: logs/<year>/<month>/<date>/<job_id>.log
        Uses the same PipeFormatter as the terminal handler.
        """
        if job_id not in self.job_loggers:
            job_logger = logging.getLogger(f"job.{job_id}")
            job_logger.setLevel(logging.DEBUG)
            # Prevent propagation to root — we add our own file handler only
            job_logger.propagate = False
            job_logger.handlers.clear()

            file_handler = get_job_file_handler(job_id, base_log_dir=BASE_LOG_DIR)
            job_logger.addHandler(file_handler)

            self.job_loggers[job_id] = job_logger

        return self.job_loggers[job_id]

    async def add_job(self, job_data: Dict[str, Any]) -> str:
        """Add a print job to the queue with enhanced tracking"""
        # Check if queue is full
        if self.queue.full():
            raise ValueError(f"Queue is full (max {config.MAX_QUEUE_SIZE} jobs). Please try again later.")

        job_id = f"job_{len(self.jobs) + 1}_{uuid.uuid4().hex[:8]}"

        # Create input directory for this job
        job_input_dir = os.path.join(config.INPUT_DIR, job_id)
        os.makedirs(job_input_dir, exist_ok=True)

        # Save input PDF to disk immediately (avoid storing base64 in memory)
        input_filename = f"input_{job_data['filename']}"
        input_path = os.path.join(job_input_dir, input_filename)
        
        try:
            # Decode base64 and write to disk
            import base64
            pdf_bytes = base64.b64decode(job_data['file_content'])
            with open(input_path, 'wb') as f:
                f.write(pdf_bytes)
            logger.info(f"Saved input PDF to disk: {input_path}")
        except Exception as e:
            logger.error(f"Failed to save input PDF for job {job_id}: {e}")
            raise

        # Create job with path reference (not base64 content)
        job = JobData(
            id=job_id,
            filename=job_data['filename'],
            input_path=input_path,  # Store disk path instead of base64
            status='queued',
            created_at=datetime.now()
        )

        # Store in jobs registry
        self.jobs[job_id] = job

        # Add to processing queue
        await self.queue.put(job)

        logger.info(f"Added job {job_id} to queue: {job.filename}")
        job_state_logger.record_queued(job_id, job.filename)
        return job_id

    async def process_queue(self):
        """Process jobs from the queue sequentially"""
        while True:
            if not self.processing and not self.queue.empty():
                self.processing = True
                job = await self.queue.get()
                self.current_job = job

                # Update job status
                job.status = 'processing'
                logger.info(f"Processing job {job.id}: {job.filename}")
                job_state_logger.record_processing(job.id)

                try:
                    # Process the job using the automation service
                    output_path = await self._process_print_job(job)
                    job.status = 'completed'
                    job.completed_at = datetime.now()
                    job.output_path = output_path
                    job.output_filename = os.path.basename(output_path)
                    
                    # We don't store Base64 in result anymore to save memory
                    job.result = "File stored on disk" 
                    
                    logger.info(f"Completed job {job.id}. Output: {output_path}")
                    job_state_logger.record_completed(job.id)

                except AutomationBaseError as e:
                    job.status = 'failed'
                    job.completed_at = datetime.now()
                    job.error = str(e)
                    job.error_type = e.error_type
                    logger.error(f"Failed job {job.id} [{e.error_type}]: {e}")
                    job_state_logger.record_failed(job.id, str(e), e.error_type)

                except Exception as e:
                    job.status = 'failed'
                    job.completed_at = datetime.now()
                    job.error = str(e)
                    job.error_type = 'unknown'
                    logger.error(f"Failed job {job.id} [unknown]: {e}")
                    job_state_logger.record_failed(job.id, str(e), "unknown")

                finally:
                    self.processing = False
                    self.current_job = None

            await asyncio.sleep(config.QUEUE_CHECK_INTERVAL)

    async def _process_print_job(self, job: JobData) -> str:
        """Process a single print job using the automation service"""
        job_logger = self._get_job_logger(job.id)

        try:
            start_time = datetime.now()
            job_logger.info(f"START: Processing job {job.id} - File: {job.filename}")
            job_logger.info(f"Timestamp: {start_time}")
            job_logger.info(f"Status: STARTED")

            logger.info(f"Starting print for job {job.id}: {job.filename}")

            # Use the automation service to process the job
            job_logger.info("PHASE: Calling automation service")
            result = await self.automation.process_pdf_job(job)

            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()

            job_logger.info(f"Status: COMPLETED")
            job_logger.info(f"End Time: {end_time}")
            job_logger.info(f"Duration: {duration:.2f} seconds")
            job_logger.info(f"Result: Base64 output generated ({len(result)} chars)")

            logger.info(f"Print completed for job {job.id}: {job.filename}")
            return result

        except Exception as e:
            job_logger.error(f"Status: FAILED")
            job_logger.error(f"Error: {str(e)}")
            job_logger.error(f"Timestamp: {datetime.now()}")
            logger.error(f"Unexpected error during print job {job.id}: {str(e)}")
            raise

    def get_job_status(self, job_id: str) -> Optional[JobData]:
        """Get status of a specific job"""
        return self.jobs.get(job_id)

    def get_all_jobs(self) -> Dict[str, JobData]:
        """Get all jobs with their current status"""
        return self.jobs.copy()

    def get_queue_size(self) -> int:
        """Get current queue size"""
        return self.queue.qsize()

    def is_processing(self) -> bool:
        """Check if queue is currently processing a job"""
        return self.processing

    def get_current_job(self) -> Optional[JobData]:
        """Get currently processing job"""
        return self.current_job

    async def cleanup_old_jobs(self):
        """Remove jobs older than TTL to prevent memory leaks"""
        try:
            now = datetime.now()
            ttl_seconds = config.JOB_TTL_HOURS * 3600
            jobs_to_remove = []

            for job_id, job in self.jobs.items():
                # Use completed_at if available, otherwise created_at
                job_time = job.completed_at if job.completed_at else job.created_at
                age_seconds = (now - job_time).total_seconds()

                if age_seconds > ttl_seconds:
                    jobs_to_remove.append(job_id)

            # Remove old jobs
            for job_id in jobs_to_remove:
                # Cleanup physical input folder
                job_input_dir = os.path.join(config.INPUT_DIR, job_id)
                if os.path.exists(job_input_dir):
                    try:
                        import shutil
                        shutil.rmtree(job_input_dir)
                        logger.info(f"Cleaned up input storage for job: {job_id}")
                    except Exception as e:
                        logger.error(f"Failed to delete input dir for {job_id}: {e}")

                # Cleanup physical output folder
                job_output_dir = os.path.join(config.OUTPUT_DIR, job_id)
                if os.path.exists(job_output_dir):
                    try:
                        import shutil
                        shutil.rmtree(job_output_dir)
                        logger.info(f"Cleaned up output storage for job: {job_id}")
                    except Exception as e:
                        logger.error(f"Failed to delete output dir for {job_id}: {e}")

                del self.jobs[job_id]
                logger.info(f"Cleaned up job metadata: {job_id}")

            if jobs_to_remove:
                logger.info(f"Cleanup complete: Removed {len(jobs_to_remove)} old jobs and their files")

        except Exception as e:
            logger.error(f"Error during job cleanup: {str(e)}")

    async def cleanup_task(self):
        """Background task to periodically clean up old jobs"""
        # Run one initial cleanup for orphaned folders on startup
        await self.startup_cleanup()
        
        while True:
            await asyncio.sleep(config.CLEANUP_INTERVAL_SECONDS)
            await self.cleanup_old_jobs()

    async def startup_cleanup(self):
        """Clean up any orphaned input/output folders from previous sessions"""
        try:
            import shutil
            
            # Cleanup orphaned INPUT folders
            logger.info("Running startup cleanup for input directories...")
            if os.path.exists(config.INPUT_DIR):
                for folder_name in os.listdir(config.INPUT_DIR):
                    folder_path = os.path.join(config.INPUT_DIR, folder_name)
                    if os.path.isdir(folder_path):
                        mtime = os.path.getmtime(folder_path)
                        age_hours = (datetime.now().timestamp() - mtime) / 3600
                        
                        if age_hours > config.JOB_TTL_HOURS:
                            shutil.rmtree(folder_path)
                            logger.info(f"Startup cleanup: Removed orphaned input folder {folder_name}")

            # Cleanup orphaned OUTPUT folders
            logger.info("Running startup cleanup for output directories...")
            if os.path.exists(config.OUTPUT_DIR):
                for folder_name in os.listdir(config.OUTPUT_DIR):
                    folder_path = os.path.join(config.OUTPUT_DIR, folder_name)
                    if os.path.isdir(folder_path):
                        mtime = os.path.getmtime(folder_path)
                        age_hours = (datetime.now().timestamp() - mtime) / 3600
                        
                        if age_hours > config.JOB_TTL_HOURS:
                            shutil.rmtree(folder_path)
                            logger.info(f"Startup cleanup: Removed orphaned output folder {folder_name}")

        except Exception as e:
            logger.error(f"Error during startup cleanup: {e}")

    async def shutdown(self):
        """Gracefully shutdown the queue system"""
        logger.info("Shutting down print queue...")
        self.executor.shutdown(wait=True)
        await self.automation.cleanup()
        logger.info("Print queue shutdown complete")
