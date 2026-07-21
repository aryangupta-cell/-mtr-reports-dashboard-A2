"""Render-side connector: runs the pipeline on the company server over SSH
instead of locally, so Render never loads the 200MB+ files into its own
(RAM-constrained) process at all — it just proxies bytes in and out.

One SSH connection per job, running server_worker.py fresh each time (no
persistent process on their end, per their restriction). Auth is via an
SSH key (not a password).

Uses the system `ssh` binary via subprocess rather than paramiko —
paramiko's Channel.write() reliably raised "Socket is closed" on this
payload size/remote-command combination (confirmed by direct testing,
both with and without chunked writes); the plain OpenSSH client handled
the identical multi-hundred-MB payload without issue.
"""

import io
import json
import logging
import os
import subprocess
import tempfile
import zipfile

log = logging.getLogger("ssh_worker")


def _build_input_zip(
    mtr_csv: bytes,
    consignment_xlsx: bytes,
    primary_plants_xlsx: bytes,
    run_date_label: str,
    xswift_live_dashboard_xlsx: bytes | None,
    at_live_dashboard_xlsx: bytes | None,
) -> bytes:
    has_dashboards = bool(xswift_live_dashboard_xlsx and at_live_dashboard_xlsx)
    manifest = {"run_date_label": run_date_label, "has_dashboards": has_dashboards}

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("manifest.json", json.dumps(manifest))
        zf.writestr("mtr_csv", mtr_csv)
        zf.writestr("consignment_xlsx", consignment_xlsx)
        zf.writestr("primary_plants_xlsx", primary_plants_xlsx)
        if has_dashboards:
            zf.writestr("xswift_live_dashboard_xlsx", xswift_live_dashboard_xlsx)
            zf.writestr("at_live_dashboard_xlsx", at_live_dashboard_xlsx)
    return buf.getvalue()


def run_on_company_server(
    mtr_csv: bytes,
    consignment_xlsx: bytes,
    primary_plants_xlsx: bytes,
    run_date_label: str,
    xswift_live_dashboard_xlsx: bytes | None = None,
    at_live_dashboard_xlsx: bytes | None = None,
) -> dict[str, bytes]:
    """SSHes into the company server, runs server_worker.py once, streams
    the input zip in over stdin and the output zip back over stdout.
    Returns {output_filename: file_bytes}, identical shape to
    mtr_analysis.run_in_memory() — this function is a drop-in remote
    version of it.
    """
    input_zip = _build_input_zip(
        mtr_csv, consignment_xlsx, primary_plants_xlsx, run_date_label,
        xswift_live_dashboard_xlsx, at_live_dashboard_xlsx,
    )

    ssh_host = os.environ["SSH_HOST"]
    ssh_user = os.environ["SSH_USER"]
    ssh_key_path = os.environ["SSH_KEY_PATH"]
    remote_dir = os.environ["REMOTE_DIR"]
    remote_python = os.environ.get("REMOTE_PYTHON", "python3")
    ssh_port = os.environ.get("SSH_PORT", "22")
    # Render's disk is ephemeral (wiped on every restart/redeploy), so
    # there's no persistent known_hosts to trust-on-first-use against.
    # Pin the server's host key explicitly via SSH_HOST_KEY (the base64
    # blob from `ssh-keyscan -t ed25519 <host>`, third field) so we still
    # verify we're talking to the real server. Falls back to
    # accept-new (trust on first use, logged loudly) if unset.
    pinned_host_key = os.environ.get("SSH_HOST_KEY")

    known_hosts_path = None
    ssh_opts = ["-o", "BatchMode=yes", "-p", ssh_port]
    try:
        if pinned_host_key:
            fd, known_hosts_path = tempfile.mkstemp(suffix="_known_hosts")
            with os.fdopen(fd, "w") as f:
                f.write(f"{ssh_host} ssh-ed25519 {pinned_host_key}\n")
            ssh_opts += ["-o", f"UserKnownHostsFile={known_hosts_path}", "-o", "StrictHostKeyChecking=yes"]
        else:
            log.warning(
                "SSH_HOST_KEY not set — accepting the server's host key on trust "
                "instead of verifying it. Set SSH_HOST_KEY to pin it properly."
            )
            ssh_opts += ["-o", "UserKnownHostsFile=/dev/null", "-o", "StrictHostKeyChecking=accept-new"]

        command = [
            "ssh", "-i", ssh_key_path, *ssh_opts, f"{ssh_user}@{ssh_host}",
            f"cd {remote_dir} && {remote_python} server_worker.py",
        ]
        proc = subprocess.run(command, input=input_zip, capture_output=True, timeout=1800)
    finally:
        if known_hosts_path:
            os.unlink(known_hosts_path)

    if proc.returncode != 0:
        err_text = proc.stderr.decode("utf-8", errors="replace")
        raise RuntimeError(
            f"Remote pipeline failed (exit {proc.returncode}) on {ssh_host}:\n{err_text}"
        )

    outputs: dict[str, bytes] = {}
    with zipfile.ZipFile(io.BytesIO(proc.stdout)) as zf:
        for name in zf.namelist():
            outputs[name] = zf.read(name)
    return outputs
