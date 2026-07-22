"""FastAPI backend for the ATC Report Dashboard.

Restructured 2026-07-22 into 4 independent report tabs (mirrors the JKLC
dashboard's pattern): each report has only its own required inputs and
produces exactly one output file — MTR Analysis, Trip Repush, Mapping
Issue, Vehicle Status.

Flow (why it's async, not a single request):
  1. POST /api/jobs streams only the files the selected report needs
     straight to per-job temp files on Render's own disk (chunked copy —
     Render's process memory never holds a full file at once, regardless
     of file size) and returns immediately with a job_id.
  2. A background thread zips those temp files (again streamed, not
     buffered — see ssh_worker.build_input_zip) and sends the zip over SSH
     to the company server, which runs the selected report function
     (mtr_analysis.py's run_*_report(), RAM-only on THEIR end) and streams
     the single output file back inside a zip to another temp file here.
     This step can take minutes — it must NOT happen inside the upload
     request, or the browser/Render's reverse proxy will time out.
  3. Frontend polls GET /api/jobs/<job_id> until status == "done" (or
     "failed"), then downloads the output file from
     GET /api/jobs/<job_id>/download/<filename> — streamed out of the
     result zip, never fully buffered either.

Only ONE job runs at a time (max_workers=1) — deliberately serializes
requests to the company server rather than hammering it with concurrent
SSH sessions.

Render's disk is used here only as ephemeral, per-job scratch space
(cleaned up on a TTL below) — that's fine; the "don't write to disk"
restriction is specific to the COMPANY server (server_worker.py, which
this file never touches), not Render.
"""

import shutil
import threading
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

from ssh_worker import build_input_zip, run_on_company_server
from jobs_store import create_job, get_job, update_job, purge_older_than

BASE_DIR = Path(__file__).parent.parent
JOBS_DIR = BASE_DIR / "jobs"
JOBS_DIR.mkdir(exist_ok=True)

JOB_TTL_SECONDS = 2 * 60 * 60  # delete job scratch files 2 hours after creation
COPY_CHUNK_SIZE = 1024 * 1024  # 1MB — keeps upload-save memory flat regardless of file size

# Which uploaded files each report actually needs — the single source of
# truth for both request validation and what gets zipped for the company
# server. Keep in sync with server_worker.py's REPORT_OUTPUT_FILENAMES and
# mtr_analysis.py's run_*_report() signatures.
REPORT_REQUIRED_FILES = {
    "mtr_analysis": ["mtr_csv", "consignment_xlsx", "primary_plants_xlsx"],
    "trip_repush": ["mtr_csv", "consignment_xlsx", "primary_plants_xlsx"],
    "mapping_issue": ["xswift_live_dashboard_xlsx", "at_live_dashboard_xlsx", "primary_plants_xlsx"],
    "vehicle_status": ["xswift_live_dashboard_xlsx", "at_live_dashboard_xlsx"],
}

app = FastAPI(title="ATC Report Dashboard API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Single worker on purpose — see module docstring.
executor = ThreadPoolExecutor(max_workers=1)


async def _save_upload_streamed(upload: UploadFile, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("wb") as f:
        while chunk := await upload.read(COPY_CHUNK_SIZE):
            f.write(chunk)


def _run_job(job_id: str, job_dir: Path, report: str, run_date_label: str, saved_files: dict[str, Path]) -> None:
    update_job(job_id, status="running")
    try:
        input_zip_path = job_dir / "input.zip"
        build_input_zip(
            zip_path=input_zip_path, report=report, run_date_label=run_date_label, files=saved_files,
        )

        output_zip_path = job_dir / "output.zip"
        run_on_company_server(input_zip_path, output_zip_path)

        with zipfile.ZipFile(output_zip_path) as zf:
            output_files = zf.namelist()

        # Input files and the input zip are no longer needed once the
        # remote run has finished — free the disk space now rather than
        # waiting for the TTL sweep.
        shutil.rmtree(job_dir / "input", ignore_errors=True)
        input_zip_path.unlink(missing_ok=True)

        update_job(job_id, status="done", output_zip_path=str(output_zip_path), output_files=output_files)
    except Exception as exc:  # noqa: BLE001 — surface any failure to the UI
        update_job(job_id, status="failed", error=str(exc))


@app.post("/api/jobs")
async def create_analysis_job(
    report: str = Form(...),
    run_date_label: str = Form(...),
    mtr_csv: UploadFile | None = File(None),
    consignment_xlsx: UploadFile | None = File(None),
    primary_plants_xlsx: UploadFile | None = File(None),
    xswift_live_dashboard_xlsx: UploadFile | None = File(None),
    at_live_dashboard_xlsx: UploadFile | None = File(None),
):
    if report not in REPORT_REQUIRED_FILES:
        raise HTTPException(status_code=400, detail=f"Unknown report: {report!r}")

    uploads = {
        "mtr_csv": mtr_csv,
        "consignment_xlsx": consignment_xlsx,
        "primary_plants_xlsx": primary_plants_xlsx,
        "xswift_live_dashboard_xlsx": xswift_live_dashboard_xlsx,
        "at_live_dashboard_xlsx": at_live_dashboard_xlsx,
    }
    required = REPORT_REQUIRED_FILES[report]
    missing = [name for name in required if uploads[name] is None]
    if missing:
        raise HTTPException(status_code=400, detail=f"Missing required file(s) for {report}: {missing}")

    job = create_job()
    job_dir = JOBS_DIR / job.id
    input_dir = job_dir / "input"

    saved_files: dict[str, Path] = {}
    for name in required:
        dest = input_dir / f"{name}.xlsx" if name != "mtr_csv" else input_dir / "mtr_csv.csv"
        await _save_upload_streamed(uploads[name], dest)
        saved_files[name] = dest

    executor.submit(_run_job, job.id, job_dir, report, run_date_label, saved_files)
    return {"job_id": job.id, "status": job.status}


@app.get("/api/jobs/{job_id}")
async def job_status(job_id: str):
    job = get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found (it may have expired)")
    return {
        "job_id": job.id,
        "status": job.status,
        "error": job.error,
        "output_files": sorted(job.output_files),
    }


def _stream_zip_member(zip_path: str, member: str):
    with zipfile.ZipFile(zip_path) as zf, zf.open(member) as f:
        while chunk := f.read(COPY_CHUNK_SIZE):
            yield chunk


@app.get("/api/jobs/{job_id}/download/{filename}")
async def download_output(job_id: str, filename: str):
    job = get_job(job_id)
    if job is None or job.status != "done" or job.output_zip_path is None:
        raise HTTPException(status_code=404, detail="Job not found or not finished yet")
    if filename not in job.output_files:
        raise HTTPException(status_code=404, detail="File not found for this job")
    if not Path(job.output_zip_path).exists():
        raise HTTPException(status_code=404, detail="Result file missing on disk (may have expired)")

    return StreamingResponse(
        _stream_zip_member(job.output_zip_path, filename),
        media_type="application/octet-stream",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


def _cleanup_loop():
    while True:
        time.sleep(15 * 60)
        for job_id in purge_older_than(JOB_TTL_SECONDS):
            shutil.rmtree(JOBS_DIR / job_id, ignore_errors=True)


threading.Thread(target=_cleanup_loop, daemon=True).start()

# Serve the frontend (index.html + assets) from the same service —
# simplest deploy: one Render web service, no separate static host.
frontend_dir = BASE_DIR / "frontend"
if frontend_dir.exists():
    app.mount("/", StaticFiles(directory=str(frontend_dir), html=True), name="frontend")
