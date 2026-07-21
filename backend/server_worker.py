"""Runs on the company server, one process per request — no persistent
service, no disk writes. Invoked over SSH by the Render-hosted app.py.

Protocol (all binary, over stdin/stdout — logging goes to stderr so it
never corrupts the binary stdout stream):
  stdin:  a single zip file containing:
            manifest.json          {"run_date_label": "...", "has_dashboards": true/false}
            mtr_csv
            consignment_xlsx
            primary_plants_xlsx
            xswift_live_dashboard_xlsx   (only if has_dashboards)
            at_live_dashboard_xlsx       (only if has_dashboards)
  stdout: a single zip file containing the 3 output reports, xlsx only
            (MTR_Analysis_-_<label>.xlsx, Trip_Repush_-_<label>.xlsx,
             Mapping_issue_-_<label>.xlsx) — filenames are whatever
            mtr_analysis.run_in_memory() returns, not hardcoded here.

Everything happens in RAM via run_in_memory() — this script never calls
open() on a real path, never writes a temp file, and exits as soon as
the run is done.
"""

import io
import json
import sys
import zipfile

from mtr_analysis import run_in_memory


def main() -> int:
    input_zip_bytes = sys.stdin.buffer.read()

    with zipfile.ZipFile(io.BytesIO(input_zip_bytes)) as zf:
        manifest = json.loads(zf.read("manifest.json"))
        mtr_csv = zf.read("mtr_csv")
        consignment_xlsx = zf.read("consignment_xlsx")
        primary_plants_xlsx = zf.read("primary_plants_xlsx")
        xswift_live_dashboard_xlsx = None
        at_live_dashboard_xlsx = None
        if manifest.get("has_dashboards"):
            xswift_live_dashboard_xlsx = zf.read("xswift_live_dashboard_xlsx")
            at_live_dashboard_xlsx = zf.read("at_live_dashboard_xlsx")

    outputs = run_in_memory(
        mtr_csv=mtr_csv,
        consignment_xlsx=consignment_xlsx,
        primary_plants_xlsx=primary_plants_xlsx,
        run_date_label=manifest["run_date_label"],
        xswift_live_dashboard_xlsx=xswift_live_dashboard_xlsx,
        at_live_dashboard_xlsx=at_live_dashboard_xlsx,
    )

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, data in outputs.items():
            zf.writestr(filename, data)

    sys.stdout.buffer.write(out_buf.getvalue())
    sys.stdout.buffer.flush()
    return 0


if __name__ == "__main__":
    sys.exit(main())
