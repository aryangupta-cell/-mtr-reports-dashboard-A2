"""Runs on the company server, one process per request — no persistent
service, no disk writes. Invoked over SSH by the Render-hosted app.py.

Protocol (all binary, over stdin/stdout — logging goes to stderr so it
never corrupts the binary stdout stream):
  stdin:  a single zip file containing:
            manifest.json   {"report": "mtr_analysis"|"trip_repush"|
                                        "mapping_issue"|"vehicle_status",
                              "run_date_label": "..."}
            + only the input files that specific report needs, using
              these arcnames: mtr_csv, consignment_xlsx,
              primary_plants_xlsx, xswift_live_dashboard_xlsx,
              at_live_dashboard_xlsx
  stdout: a single zip file containing exactly ONE output file — the
            dashboard was restructured (2026-07-22) into 4 independent
            report tabs, each producing one file, matching the JKLC
            dashboard's pattern.

Everything happens in RAM — this script never calls open() on a real
path, never writes a temp file, and exits as soon as the run is done.
"""

import io
import json
import sys
import zipfile

from mtr_analysis import (
    run_mtr_analysis_report,
    run_trip_repush_report,
    run_mapping_issue_report,
    run_vehicle_status_report,
)

REPORT_OUTPUT_FILENAMES = {
    "mtr_analysis": "MTR_Analysis_-_{label}.xlsx",
    "trip_repush": "Trip_Repush_-_{label}.xlsx",
    "mapping_issue": "Mapping_issue_-_{label}.xlsx",
    "vehicle_status": "Vehicle_Status_-_{label}.xlsx",
}


def main() -> int:
    input_zip_bytes = sys.stdin.buffer.read()

    with zipfile.ZipFile(io.BytesIO(input_zip_bytes)) as zf:
        manifest = json.loads(zf.read("manifest.json"))
        names = set(zf.namelist())

        def read(arcname):
            return zf.read(arcname) if arcname in names else None

        report = manifest["report"]
        run_date_label = manifest["run_date_label"]

        if report == "mtr_analysis":
            output_bytes = run_mtr_analysis_report(
                mtr_csv=read("mtr_csv"),
                consignment_xlsx=read("consignment_xlsx"),
                primary_plants_xlsx=read("primary_plants_xlsx"),
                run_date_label=run_date_label,
            )
        elif report == "trip_repush":
            output_bytes = run_trip_repush_report(
                mtr_csv=read("mtr_csv"),
                consignment_xlsx=read("consignment_xlsx"),
                primary_plants_xlsx=read("primary_plants_xlsx"),
                run_date_label=run_date_label,
            )
        elif report == "mapping_issue":
            output_bytes = run_mapping_issue_report(
                xswift_live_dashboard_xlsx=read("xswift_live_dashboard_xlsx"),
                at_live_dashboard_xlsx=read("at_live_dashboard_xlsx"),
                primary_plants_xlsx=read("primary_plants_xlsx"),
                run_date_label=run_date_label,
            )
        elif report == "vehicle_status":
            output_bytes = run_vehicle_status_report(
                xswift_live_dashboard_xlsx=read("xswift_live_dashboard_xlsx"),
                at_live_dashboard_xlsx=read("at_live_dashboard_xlsx"),
            )
        else:
            raise ValueError(f"Unknown report type: {report!r}")

    output_filename = REPORT_OUTPUT_FILENAMES[report].format(label=run_date_label)

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(output_filename, output_bytes)

    sys.stdout.buffer.write(out_buf.getvalue())
    sys.stdout.buffer.flush()
    return 0


if __name__ == "__main__":
    sys.exit(main())
