"""In-memory job store. Single-process only — fine for one Render instance.
If you ever scale to multiple workers/instances, swap this for Redis."""

import threading
import time
import uuid
from dataclasses import dataclass, field


@dataclass
class Job:
    id: str
    status: str = "queued"          # queued -> running -> done | failed
    error: str | None = None
    output_files: dict[str, bytes] = field(default_factory=dict)   # {filename: file_bytes}, RAM-only
    created_at: float = field(default_factory=time.time)


_lock = threading.Lock()
_jobs: dict[str, Job] = {}


def create_job() -> Job:
    job = Job(id=str(uuid.uuid4()))
    with _lock:
        _jobs[job.id] = job
    return job


def get_job(job_id: str) -> Job | None:
    with _lock:
        return _jobs.get(job_id)


def update_job(job_id: str, **kwargs) -> None:
    with _lock:
        job = _jobs[job_id]
        for k, v in kwargs.items():
            setattr(job, k, v)


def purge_older_than(seconds: float) -> list[str]:
    """Returns job ids to clean up (caller deletes the actual files)."""
    cutoff = time.time() - seconds
    with _lock:
        stale = [jid for jid, j in _jobs.items() if j.created_at < cutoff]
        for jid in stale:
            del _jobs[jid]
    return stale
