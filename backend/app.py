"""FastAPI backend for the MTR Analysis dashboard.

Flow (why it's async, not a single request):
  1. POST /api/jobs reads the 5 uploaded files into memory (never written to
     Render's disk) and returns immediately with a job_id.
  2. A background thread sends those bytes over SSH to the company server,
     which runs the actual pipeline (server_worker.py -> run_in_memory())
     and streams the output files back. This step takes several minutes —
     it must NOT happen inside the upload request, or the browser/Render's
     reverse proxy will time out.
  3. Frontend polls GET /api/jobs/<job_id> until status == "done" (or
     "failed"), then downloads each output file from
     GET /api/jobs/<job_id>/download/<filename> — served straight from the
     in-memory bytes we got back over SSH.

Only ONE job runs at a time (max_workers=1) — deliberately serializes
requests to the company server rather than hammering it with concurrent
SSH sessions.

Nothing in this file ever touches Render's disk. All files are bytes in
memory, both incoming (uploads) and outgoing (results from the company
server).
"""

import threading
import time
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles
from pathlib import Path

from ssh_worker import run_on_company_server
from jobs_store import create_job, get_job, update_job, purge_older_than

BASE_DIR = Path(__file__).parent.parent
JOB_TTL_SECONDS = 2 * 60 * 60  # forget job results 2 hours after creation

app = FastAPI(title="MTR Analysis Dashboard API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Single worker on purpose — see module docstring.
executor = ThreadPoolExecutor(max_workers=1)


def _run_job(
    job_id: str,
    mtr_csv: bytes,
    consignment_xlsx: bytes,
    primary_plants_xlsx: bytes,
    run_date_label: str,
    xswift_live_dashboard_xlsx: bytes | None,
    at_live_dashboard_xlsx: bytes | None,
) -> None:
    update_job(job_id, status="running")
    try:
        outputs = run_on_company_server(
            mtr_csv=mtr_csv,
            consignment_xlsx=consignment_xlsx,
            primary_plants_xlsx=primary_plants_xlsx,
            run_date_label=run_date_label,
            xswift_live_dashboard_xlsx=xswift_live_dashboard_xlsx,
            at_live_dashboard_xlsx=at_live_dashboard_xlsx,
        )
        update_job(job_id, status="done", output_files=outputs)
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

    mtr_csv_bytes = await mtr_csv.read()
    consignment_bytes = await consignment_xlsx.read()
    plants_bytes = await primary_plants_xlsx.read()
    xswift_bytes = await xswift_live_dashboard_xlsx.read() if xswift_live_dashboard_xlsx else None
    at_bytes = await at_live_dashboard_xlsx.read() if at_live_dashboard_xlsx else None

    executor.submit(
        _run_job, job.id,
        mtr_csv_bytes, consignment_bytes, plants_bytes, run_date_label,
        xswift_bytes, at_bytes,
    )
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
        "output_files": sorted(job.output_files.keys()),
    }


@app.get("/api/jobs/{job_id}/download/{filename}")
async def download_output(job_id: str, filename: str):
    job = get_job(job_id)
    if job is None or job.status != "done":
        raise HTTPException(status_code=404, detail="Job not found or not finished yet")
    data = job.output_files.get(filename)
    if data is None:
        raise HTTPException(status_code=404, detail="File not found for this job")
    return Response(
        content=data,
        media_type="application/octet-stream",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


def _cleanup_loop():
    while True:
        time.sleep(15 * 60)
        purge_older_than(JOB_TTL_SECONDS)


threading.Thread(target=_cleanup_loop, daemon=True).start()

# Serve the frontend (index.html + assets) from the same service —
# simplest deploy: one Render web service, no separate static host.
frontend_dir = BASE_DIR / "frontend"
if frontend_dir.exists():
    app.mount("/", StaticFiles(directory=str(frontend_dir), html=True), name="frontend")
