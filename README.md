# PDF Print API & Automation Utility

## Overview
A robust, high-performance API and RPA service designed for automated PDF processing using Adobe Acrobat 9 Pro. This utility bridges the gap between modern RESTful microservices and legacy desktop applications by providing a reliable, queued interface for UI-driven print automation.

## Core Features
- **RESTful API**: FastAPI-based architecture with secure `Bearer` authentication.
- **Asynchronous Queue Management**: Built-in job queuing with configurable concurrency and capacity limits.
- **RPA Engine**: Robust UI automation using `PyAutoGUI` and `Win32API` for reliable interaction with Adobe Acrobat 9 Pro.
- **Fault-Tolerant Processing**: Comprehensive error handling for window detection, dialog timeouts, and file validations.
- **Structured Logging**: Pipe-delimited terminal logs for enhanced observability and system monitoring.
- **Persistence & Cleanup**: Automated directory management for inputs/outputs with configurable job TTL (Time-To-Live).

## Technical Architecture
The system operates as a state-machine managed through a background worker:
1. **Ingestion**: PDFs are received via `/print-queue` and stored in the `inputs/` directory.
2. **Queuing**: Jobs are assigned a unique ID and placed in the processing queue.
3. **Execution**: The RPA service activates Acrobat, executes the print command (Ctrl+P), handles the "Save As" dialog, and verifies the output.
4. **Retrieval**: Completed jobs provide a Base64 encoded result or path via the `/job-status/{job_id}` endpoint.

## Prerequisites
- **Operating System**: Windows (required for Win32 UI automation).
- **Software**: Adobe Acrobat 9 Pro (must be installed and set as the default PDF handler).
- **Environment**: Python 3.9+

## Installation

1. **Clone the Repository**:
   ```powershell
   git clone <repository-url>
   cd "PDF Utility/Printing"
   ```

2. **Setup Virtual Environment**:
   ```powershell
   python -m venv print_env
   .\print_env\Scripts\activate
   ```

3. **Install Dependencies**:
   ```powershell
   pip install -r requirements.txt
   ```

## Configuration
All system parameters are defined in `config.py`. Key configurations include:
- `API_KEY`: Authentication token for protected endpoints.
- `PORT`: API port (default: 8001).
- `MAX_QUEUE_SIZE`: Maximum number of pending jobs.
- `JOB_TTL_HOURS`: Lifespan of job files before automatic cleanup.

## API Documentation

### 1. Health Check
`GET /health`  
Returns system status, queue metrics, and current processing state.

### 2. Queue Print Job
`POST /print-queue` (Protected)  
**Requirement**: Multipart file upload.  
**Response**: Returns `job_id` for status tracking.

### 3. Get Job Status
`GET /job-status/{job_id}` (Protected)  
**Status Values**: `queued`, `processing`, `completed`, `failed`.

## Development & Usage

### Starting the Server
```powershell
python main.py
```

### Authentication
Include the `API_KEY` in headers for all protected requests:
`Authorization: Bearer <YOUR_API_KEY>`

## Directory Structure
- `inputs/`: Temporary storage for uploaded PDF files.
- `outputs/`: Destination for processed/printed PDFs.
- `logs/`: System and job-specific logs.
- `temp/`: Working directory for UI automation tasks.

## Security & Reliability Note
- **Deterministic UI**: The RPA engine uses strict window handle verification (hwnd) to ensure actions are only performed on Adobe Acrobat processes.
- **Failsafe**: `PyAutoGUI` failsafe is enabled by default. Moving the mouse to the corner of the screen will abort automation.
