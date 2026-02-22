"""
PDF Print API - Main Application
"""
import asyncio
import base64
import logging
import os

from logging_setup import configure_root_logger
import uvicorn
from contextlib import asynccontextmanager
from fastapi import FastAPI, HTTPException, UploadFile, File, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials

from config import config
from models import FileItem, PrintJob, QueueResponse, HealthResponse
from queue_manager import PDFPrintQueue

# Configure structured terminal logging (pipe-delimited format)
configure_root_logger(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize global queue instance
print_queue = PDFPrintQueue()

@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan context manager for startup and shutdown events"""
    # Startup
    logger.info("Starting PDF Print API...")
    asyncio.create_task(print_queue.process_queue())
    asyncio.create_task(print_queue.cleanup_task())
    logger.info("Queue processor and cleanup task started")
    yield
    # Shutdown
    logger.info("Shutting down PDF Print API...")
    await print_queue.shutdown()
    logger.info("Shutdown complete")

# Initialize FastAPI app with lifespan
app = FastAPI(
    title="PDF Print API",
    version="2.0.0",
    description="Enhanced PDF Print API with robust queue management",
    lifespan=lifespan
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize security scheme
bearer_scheme = HTTPBearer()

def require_api_key(
    credentials: HTTPAuthorizationCredentials = Depends(bearer_scheme)
):
    """Validate API key for authentication"""
    if credentials.scheme.lower() != "bearer" or credentials.credentials != config.API_KEY:
        raise HTTPException(status_code=401, detail="Invalid or missing API key")

# --- API ENDPOINTS ---

@app.post("/print-queue", dependencies=[Depends(require_api_key)])
async def queue_print_job(file: UploadFile = File(...)):
    """Queue a PDF for printing using Adobe Acrobat automation."""
    try:
        # Read file content
        content = await file.read()

        # Encode to base64
        file_content = base64.b64encode(content).decode('utf-8')

        # Queue the job
        job_id = await print_queue.add_job({
            'filename': file.filename,
            'file_content': file_content
        })

        return {
            "job_id": job_id,
            "filename": file.filename,
            "message": "Print job queued successfully",
            "status": "queued"
        }

    except ValueError as e:
        # Queue is full
        logger.warning(f"Queue full error: {str(e)}")
        raise HTTPException(status_code=429, detail=str(e))
    except Exception as e:
        logger.error(f"Failed to queue print job: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Failed to queue print job: {str(e)}")

@app.get("/job-status/{job_id}", dependencies=[Depends(require_api_key)])
async def get_job_status(job_id: str):
    """Get the status of a print job."""
    job = print_queue.get_job_status(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")

    # If completed, load result from disk on-demand
    if job.status == 'completed' and job.output_path:
        try:
            if os.path.exists(job.output_path):
                with open(job.output_path, 'rb') as f:
                    output_bytes = f.read()
                job.result = base64.b64encode(output_bytes).decode('utf-8')
            else:
                # File might have been cleaned up or missing
                logger.warning(f"Output file missing for job {job_id} at {job.output_path}")
                job.status = 'failed'
                job.error = "Output file missing on disk"
        except Exception as e:
            logger.error(f"Failed to read output file for {job_id}: {e}")
            job.status = 'failed'
            job.error = f"Failed to read result: {str(e)}"

    return job



@app.get("/health")
async def health_check():
    """Health check endpoint."""
    queue_size = print_queue.get_queue_size()
    return {
        "status": "healthy",
        "queue_size": queue_size,
        "queue_limit": config.MAX_QUEUE_SIZE,
        "queue_full": queue_size >= config.MAX_QUEUE_SIZE,
        "processing": print_queue.is_processing(),
        "current_job": print_queue.get_current_job().id if print_queue.get_current_job() else None
    }

if __name__ == '__main__':
    uvicorn.run(app, host="127.0.0.1", port=8001)
