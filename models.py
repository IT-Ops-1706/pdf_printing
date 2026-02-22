"""
Data models for PDF Print Utility
"""
from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class FileItem(BaseModel):
    """Model for file upload with base64 content"""
    filename: str
    docType: str
    fileContent: str  # Base64 encoded PDF content


class PrintJob(BaseModel):
    """Model for print job status"""
    id: str
    filename: str
    status: str  # 'queued', 'processing', 'completed', 'failed'
    result: Optional[str] = None
    error: Optional[str] = None


class JobData(BaseModel):
    """Internal model for job data"""
    id: str
    filename: str
    input_path: Optional[str] = None  # Path to input PDF on disk (not base64 in memory)
    status: str
    result: Optional[str] = None
    output_path: Optional[str] = None
    output_filename: Optional[str] = None
    error: Optional[str] = None
    error_type: Optional[str] = None  # e.g. "timeout", "acrobat_not_found", "ui_automation_failed"
    created_at: datetime
    completed_at: Optional[datetime] = None


class QueueResponse(BaseModel):
    """Response model for queue operations"""
    job_id: str
    message: str
    status: str



class HealthResponse(BaseModel):
    """Response model for health check"""
    status: str
    queue_size: int
    processing: bool
