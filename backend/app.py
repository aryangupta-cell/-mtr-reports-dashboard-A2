"""FastAPI backend for the MTR Analysis dashboard.

Flow (why it's async, not a single request):
  1. POST /api/jobs streams the 5 uploaded files straight to per-job temp
     files on Render's own disk (chunked copy — Render's process memory
     never holds a full file at once, regardless of file size) and returns
     immediately with a job_id.
  2. A background thread zips those temp files (again streamed, not
     buffered — see ssh_worker.build_input_zip) and sends the zip over SSH
     to the company server, which runs the actual pipeline
     (server_worker.py -> run_in_memory(), RAM-only on THEIR end) and
     streams the output zip straight back to another temp file here. This
     step takes several minutes — it must NOT happen inside the upload
     request, or the browser/Render's reverse proxy will time out.
  3. Frontend polls GET /api/jobs/<job_id> until status == "done" (or
     "failed"), then downloads each output file from
     GET /api/jobs/<job_id>/download/<filename> — streamed out of the
     result zip member by member, never fully buffered either.

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

app = FastAPI(title="MTR Analysis Dashboard API")
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


def _run_job(job_id: str, job_dir: Path, run_date_label: str, has_dashboards: bool) -> None:
    update_job(job_id, status="running")
    input_dir = job_dir / "input"
    try:
        input_zip_path = job_dir / "input.zip"
        build_input_zip(
            zip_path=input_zip_path,
            mtr_csv_path=input_dir / "mtr.csv",
            consignment_xlsx_path=input_dir / "consignment.xlsx",
            primary_plants_xlsx_path=input_dir / "primary_plants.xlsx",
            run_date_label=run_date_label,
            xswift_live_dashboard_xlsx_path=(input_dir / "xswift_live_dashboard.xlsx") if has_dashboards else None,
            at_live_dashboard_xlsx_path=(input_dir / "at_live_dashboard.xlsx") if has_dashboards else None,
        )

        output_zip_path = job_dir / "output.zip"
        run_on_company_server(input_zip_path, output_zip_path)

        with zipfile.ZipFile(output_zip_path) as zf:
            output_files = zf.namelist()

        # Input files and the input zip are no longer needed once the
        # remote run has finished — free the disk space now rather than
        # waiting for the TTL sweep.
        shutil.rmtree(input_dir, ignore_errors=True)
        input_zip_path.unlink(missing_ok=True)

        update_job(job_id, status="done", output_zip_path=str(output_zip_path), output_files=output_files)
    except Exception as exc:  # noqa: BLE001 — surface any failure to the UI
        update_job(job_id, status="failed", error=str(exc))


@app.post("/api/jobs")
async def create_analysis_job(
    run_date_label: str = Form(...),
    mtr_csv: UploadFile = File(...),
    consignment_xlsx: UploadFile = File(...),
    primary_plants_xlsx: UploadFile = File(...),
    xswift_live_dashboard_xlsx: UploadFile | None = File(None),
    at_live_dashboard_xlsx: UploadFile | None = File(None),
):
    job = create_job()
    job_dir = JOBS_DIR / job.id
    input_dir = job_dir / "input"

    await _save_upload_streamed(mtr_csv, input_dir / "mtr.csv")
    await _save_upload_streamed(consignment_xlsx, input_dir / "consignment.xlsx")
    await _save_upload_streamed(primary_plants_xlsx, input_dir / "primary_plants.xlsx")

    has_dashboards = bool(xswift_live_dashboard_xlsx and at_live_dashboard_xlsx)
    if has_dashboards:
        await _save_upload_streamed(xswift_live_dashboard_xlsx, input_dir / "xswift_live_dashboard.xlsx")
        await _save_upload_streamed(at_live_dashboard_xlsx, input_dir / "at_live_dashboard.xlsx")

    executor.submit(_run_job, job.id, job_dir, run_date_label, has_dashboards)
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
