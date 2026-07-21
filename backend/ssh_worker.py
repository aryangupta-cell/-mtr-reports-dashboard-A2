"""Render-side connector: runs the pipeline on the company server over SSH
instead of locally.

Works against file PATHS, not in-memory bytes — Render's own process must
stay well under its RAM limit while relaying files that can be 250MB+, so
everything here streams through disk/OS pipes instead of holding a full
file as a Python bytes object. (Render's disk has no "no storage"
restriction — that restriction is specific to the company server, which
this file never writes anything to; server_worker.py there still runs
RAM-only, untouched.)

One SSH connection per job, running server_worker.py fresh each time on
the company server (no persistent process on their end, per their
restriction). Auth is via an SSH key (not a password).

Uses the system `ssh` binary via subprocess rather than paramiko —
paramiko's Channel.write() reliably raised "Socket is closed" on large
payloads against this remote command (confirmed by direct testing); the
plain OpenSSH client handles the same payload without issue.
"""

import json
import logging
import os
import subprocess
import tempfile
import zipfile
from pathlib import Path

log = logging.getLogger("ssh_worker")


def build_input_zip(
    zip_path: Path,
    mtr_csv_path: Path,
    consignment_xlsx_path: Path,
    primary_plants_xlsx_path: Path,
    run_date_label: str,
    xswift_live_dashboard_xlsx_path: Path | None = None,
    at_live_dashboard_xlsx_path: Path | None = None,
) -> None:
    """Zips the already-on-disk input files into zip_path. zipfile.write()
    streams each source file in chunks internally — it never loads a full
    file into memory, so this stays flat regardless of file size.
    """
    has_dashboards = bool(xswift_live_dashboard_xlsx_path and at_live_dashboard_xlsx_path)
    manifest = {"run_date_label": run_date_label, "has_dashboards": has_dashboards}

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("manifest.json", json.dumps(manifest))
        zf.write(mtr_csv_path, arcname="mtr_csv")
        zf.write(consignment_xlsx_path, arcname="consignment_xlsx")
        zf.write(primary_plants_xlsx_path, arcname="primary_plants_xlsx")
        if has_dashboards:
            zf.write(xswift_live_dashboard_xlsx_path, arcname="xswift_live_dashboard_xlsx")
            zf.write(at_live_dashboard_xlsx_path, arcname="at_live_dashboard_xlsx")


def run_on_company_server(input_zip_path: Path, output_zip_path: Path) -> None:
    """SSHes into the company server, runs server_worker.py once, streaming
    input_zip_path's bytes in over stdin and writing the response straight
    to output_zip_path — via OS-level file-descriptor redirection
    (subprocess stdin=<file>, stdout=<file>), so Python itself never holds
    the file content in memory, only the OS pipes do.
    """
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

        with open(input_zip_path, "rb") as stdin_f, \
             open(output_zip_path, "wb") as stdout_f, \
             tempfile.TemporaryFile() as stderr_f:
            proc = subprocess.run(
                command, stdin=stdin_f, stdout=stdout_f, stderr=stderr_f, timeout=1800,
            )
            stderr_f.seek(0)
            err_text = stderr_f.read().decode("utf-8", errors="replace")
    finally:
        if known_hosts_path:
            os.unlink(known_hosts_path)

    if proc.returncode != 0:
        raise RuntimeError(
            f"Remote pipeline failed (exit {proc.returncode}) on {ssh_host}:\n{err_text}"
        )
