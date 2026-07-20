"""FastAPI backend for the MTR Analysis dashboard.

Flow (why it's async, not a single request):
  1. POST /api/jobs receives the 5 uploaded files + run_date_label, saves them
     to disk under jobs/<job_id>/input/, and returns immediately with a job_id.
  2. A background worker thread runs the actual mtr_analysis.run() pipeline
     against those saved files. This step can take minutes (200MB+ xlsx read
     with openpyxl) — it must NOT happen inside the upload request, or the
     browser/Render's reverse proxy will time out.
  3. Frontend polls GET /api/jobs/<job_id> until status == "done" (or
     "failed"), then downloads each output file from
     GET /api/jobs/<job_id>/download/<filename>.

Only ONE job runs at a time (max_workers=1) — this pipeline briefly holds
both the 200MB Consignment Report and the 300k-row MTR frame in memory;
running two concurrently on a small Render instance risks OOM.
"""

import shutil
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from mtr_analysis import Config, run
from jobs_store import create_job, get_job, update_job, purge_older_than

BASE_DIR = Path(__file__).parent.parent
JOBS_DIR = BASE_DIR / "jobs"
JOBS_DIR.mkdir(exist_ok=True)

JOB_TTL_SECONDS = 2 * 60 * 60  # delete job files 2 hours after creation

app = FastAPI(title="MTR Analysis Dashboard API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Single worker on purpose — see module docstring.
executor = ThreadPoolExecutor(max_workers=1)


def _save_upload(upload: UploadFile, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("wb") as f:
        shutil.copyfileobj(upload.file, f, length=1024 * 1024)


def _run_job(job_id: str, cfg: Config) -> None:
    update_job(job_id, status="running")
    try:
        run(cfg)
        output_dir = cfg.output_dir
        files = sorted(p.name for p in output_dir.iterdir() if p.is_file())
        update_job(job_id, status="done", output_files=files)
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
    output_dir = job_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    mtr_csv_path = input_dir / "mtr.csv"
    consignment_path = input_dir / "consignment.xlsx"
    plants_path = input_dir / "primary_plants.xlsx"
    _save_upload(mtr_csv, mtr_csv_path)
    _save_upload(consignment_xlsx, consignment_path)
    _save_upload(primary_plants_xlsx, plants_path)

    # NOTE: Mapping issue (the feature these two dashboard files feed) is
    # still ON HOLD in the pipeline — see mtr_analysis.py / README. These
    # are accepted and saved for forward-compatibility but currently unused
    # by run().
    xswift_dashboard_path = None
    at_dashboard_path = None
    if xswift_live_dashboard_xlsx is not None:
        xswift_dashboard_path = input_dir / "xswift_live_dashboard.xlsx"
        _save_upload(xswift_live_dashboard_xlsx, xswift_dashboard_path)
    if at_live_dashboard_xlsx is not None:
        at_dashboard_path = input_dir / "at_live_dashboard.xlsx"
        _save_upload(at_live_dashboard_xlsx, at_dashboard_path)

    cfg = Config(
        mtr_csv=mtr_csv_path,
        consignment_xlsx=consignment_path,
        primary_plants_xlsx=plants_path,
        xswift_live_dashboard_xlsx=xswift_dashboard_path,
        at_live_dashboard_xlsx=at_dashboard_path,
        output_dir=output_dir,
        run_date_label=run_date_label,
    )

    executor.submit(_run_job, job.id, cfg)
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
        "output_files": job.output_files,
    }


@app.get("/api/jobs/{job_id}/download/{filename}")
async def download_output(job_id: str, filename: str):
    job = get_job(job_id)
    if job is None or job.status != "done":
        raise HTTPException(status_code=404, detail="Job not found or not finished yet")
    if filename not in job.output_files:
        raise HTTPException(status_code=404, detail="File not found for this job")
    path = JOBS_DIR / job_id / "output" / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="File missing on disk")
    return FileResponse(path, filename=filename)


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
