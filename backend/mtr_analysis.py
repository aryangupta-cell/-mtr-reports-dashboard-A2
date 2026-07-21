"""
MTR Analysis Pipeline — UTCL / Anchal track
=============================================

Rebuilds, in pandas (not Excel formulas), the manual workbook Anchal
currently builds by hand: "MTR Analysis - <date>.xlsx".

WHY NOT EXCEL FORMULAS
-----------------------
The manual file uses XLOOKUP/VLOOKUP per-row against a 200MB+ external
workbook (Consignment Report) and a 281k-row MTR sheet. That's fine for
one person building it once a day in Excel, but doesn't scale to
automate: openpyxl writing 281k rows x 53 cols with formulas is slow
and memory-heavy, and Excel has to recalculate every formula on open.

Instead this script:
  1. Reads the raw MTR CSV and the AT Consignment Report with pandas
     (columnar, vectorized — not row-by-row).
  2. Builds two lookup dicts ONCE from the Consignment Report
     (City Code -> Destination, SAP PGI No -> SAP Lead Distance),
     turning what were 281k individual VLOOKUP/XLOOKUP calls into two
     O(1) dict builds + one O(n) `.map()` each.
  3. Computes every derived column as a vectorized pandas operation
     (np.select / np.where), not a per-cell Python loop.
  4. Writes the output with plain xlsxwriter (NOT constant_memory mode
     — see write_xlsx() docstring for why: constant_memory
     silently corrupts data with pandas' to_excel and was caught late,
     shipping broken output for a while. Removed entirely.).

BLANK HANDLING (IMPORTANT, CONFIRMED FROM REAL DATA)
-------------------------------------------------------
The raw MTR data does NOT use true empty cells for "no value" — it
uses a literal single-space string " " as a placeholder. Every blank
check in this script (`_is_blank`) treats NaN, empty string, AND a
whitespace-only string as blank. Do not "simplify" this to `.isna()`
only — it will silently misclassify most of the real blank rows.

STILL OPEN (see README.md) — implemented with a clearly marked
placeholder / TODO where the business rule isn't confirmed yet:
  - Sheet1 / Sheet2 pivot field layout (rows/columns/values) — not
    yet described by the business contact. A placeholder pivot is
    provided; swap PIVOT_1_CONFIG / PIVOT_2_CONFIG once confirmed.
  - Task 1 (AT <-> XSwift mapping) plant-name include/exclude filter
    on top of the Primary Plants List — currently defaults to "all
    primary plants, no further filtering". Adjust
    TASK1_EXCLUDE_PLANT_CODES / TASK1_INCLUDE_EXTRA_PLANT_CODES below
    once the real list is confirmed.
"""

from __future__ import annotations

import io
import logging
import sys
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("mtr_pipeline")


# =============================================================================
# CONFIG — edit paths/thresholds here, not in the functions below
# =============================================================================

@dataclass
class Config:
    # ---- Inputs ----
    mtr_csv: Path                      # raw XSwift MTR export, e.g. "mtr - 20 July.csv"
    consignment_xlsx: Path             # AT Consignment Report, e.g. "reports_consignts - 20 July.xlsx"
    primary_plants_xlsx: Path          # "Primary Plants List.xlsx"

    # Task 1 inputs (optional — only needed to also produce Mapping issue).
    # NAMES CORRECTED 2026-07-20 — these were previously swapped:
    # "Trip Dashboard_export...xlsx" is XSwift's Live Trip Dashboard export,
    # NOT AT's. "dashboard_export...xlsx" is AT's Live Trip Dashboard export
    # (AT's platform serves many client companies beyond UTCL, so this file
    # needs the primary-plant filter applied same as everything else).
    xswift_live_dashboard_xlsx: Path | None = None  # "Trip Dashboard_export....xlsx" — XSwift
    at_live_dashboard_xlsx: Path | None = None       # "dashboard_export....xlsx" — AT

    # ---- Outputs ----
    # xlsx ONLY (2026-07-20, explicit business requirement) — no CSV, no
    # separate output_basename; filenames match the real files exactly
    # (see run()): "MTR_Analysis_-_<date>.xlsx", "Trip_Repush_-_<date>.xlsx",
    # "Mapping_issue_-_<date>.xlsx".
    output_dir: Path = Path("./output")
    run_date_label: str = "output"          # e.g. "20 July" — used in every output filename and as the Trip Repush tab name

    # ---- Business thresholds (confirmed) ----
    sap_ai_ok_band: tuple[int, int] = (-20, 20)   # SAP-AI "0 to 20" band
    detention_slab_edges_min: tuple[int, ...] = (10, 20, 30)  # minutes

    # ---- Task 1 filter — UNCONFIRMED, see README ----
    task1_exclude_plant_codes: set[str] = field(default_factory=set)
    task1_include_extra_plant_codes: set[str] = field(default_factory=set)

    # ---- Performance ----
    csv_chunksize: int | None = None   # set e.g. 50_000 if RAM is constrained (still useful for READING the large input CSV, unrelated to output format)


# =============================================================================
# Column name constants — MUST match the real files exactly (verified 2026-07-20)
# =============================================================================

# --- Raw MTR CSV columns we read (subset; extras pass through untouched) ---
COL_TRIP_ID = "Trip ID"
COL_VEHICLE_NO = "Vehicle No."
COL_SAP_PGI_NO = "SAP PGI No"
COL_PGI_DATETIME = "PGI Date & Time"
COL_TRANSPORTER_NAME = "Transporter Name "     # NOTE: trailing space, confirmed in raw CSV header
COL_ZONE = "Zone"
COL_YARD_IN = "Yard IN"
COL_YARD_OUT = "Yard Out"
COL_YARD_DETENTION = "Yard detention"
COL_PLANT_NAME = "Plant name"
COL_PLANT_CODE = "Plant Code"
COL_PLANT_ENTRY = "Plant Entry"
COL_PLANT_EXIT = "Plant Exit"
COL_PLANT_DETENTION = "Plant Detention"
COL_DEST_CODE = "Destination Code"
COL_DESTINATION = "Destination"
COL_DEST_ENTRY = "Dest Entry Time"
COL_DEST_EXIT = "Dest Exit Time"
COL_DEST_DETENTION = "Dest Detention"
COL_STAMP_STATUS = "Stamp Status"
COL_SAP_LEAD_DIST = "Sap Lead Dist"
COL_GPS_DISTANCE = "GPS Distance"
COL_AI_REPAIRED_DIST = "AI Repaired Distance"

# --- New columns this pipeline adds (exact names, confirmed from real output file) ---
NEW_DATE_AND_TIME = "Date and time"
NEW_TRANSPORTER_REMARK = "Transporter Remark"
NEW_YARD_DETENTION_SLAB = "Yard Detention Slab"
NEW_ZONE_REMARK = "Zone Remark"
NEW_PLANT_DETENTION_SLAB = "Plant detention Slab"
NEW_AT_DEST_NAME = "AT destination name"
NEW_DEST_MATCH = "Dest. Match"
NEW_DEST_DETENTION_SLAB = "Destination detention slab"
NEW_AT_SAP_LEAD_DIST = "AT SAP lead distance"
NEW_MATCH = "Match"
NEW_AI_CHECK = "AI check"
NEW_SAP_AI = "SAP-AI"
NEW_SAP_AI_REMARK = "SAP-AI Remark"

# Stamp Status values that gate several checks (confirmed from notes)
STAMP_STATUSES_FOR_CHECKS = {"Verified", "Low Confidence"}

# --- AT Consignment Report columns we read ---
CONS_SAP_PGI_NO = "SAP PGI No"
CONS_CITY_CODE = "City Code"
CONS_DESTINATION = "Destination"
CONS_SAP_LEAD_DIST = "SAP Lead Distance (Kms)"
CONS_PLANT_CODE = "Plant Code"
CONS_VEHICLE = "Vehicle"

# --- Primary Plants List columns ---
PLANTS_COMPANY_COL = "Company"
PLANTS_CODE_COL = "Plant Code"


# =============================================================================
# Input validation — catches a mismatched/renamed file's headers up front,
# before spending minutes running the pipeline only to hit a KeyError deep
# inside. Reports exactly which file, which required columns are missing,
# and which columns are present but unrecognized (informational — extra
# columns are tolerated, not an error, per existing business decision).
# =============================================================================

# The raw input columns build_analysis_columns()/reorder_to_final_layout()
# actually require from the MTR CSV — i.e. final_order below MINUS the
# NEW_* derived columns (which don't exist until this pipeline adds them).
# Kept as an explicit list rather than computed from final_order so this
# validation doesn't depend on reorder_to_final_layout()'s internals — if
# you add a required raw column there, add it here too.
REQUIRED_MTR_CSV_COLUMNS = [
    "Trip ID", "Vehicle No.", "Vehicle Type", "SAP PGI No", "PGI Date & Time",
    "SAP Order No", "DI No", "Transporter Name ", "Transporter Code", "Zone",
    "Yard IN", "Yard Out", "Yard detention", "Plant name", "Plant Code",
    "Plant Entry", "Plant Exit", "Plant Detention", "Destination Code",
    "Destination", "Customer Name", "Dest Entry Time", "Dest Exit Time",
    "Dest Detention", "Destination Proximity End Time", "Destination Ageing",
    "Onward Duration", "Customer Segment", "Compliance Status", "Depot",
    "Route Name", "Halt", "Onward Status", "Stamp Status", "Reject Reason",
    "Sap Lead Dist", "GPS Distance", "AI Repaired Distance", "Geofence Hit/miss",
    "Mother Geofence Start Time", "Mother Geofence End Time",
    "Mother Geofence Detention", "Billing Status",
]

REQUIRED_CONSIGNMENT_COLUMNS = [
    CONS_SAP_PGI_NO, CONS_CITY_CODE, CONS_DESTINATION, CONS_SAP_LEAD_DIST, CONS_PLANT_CODE,
]

REQUIRED_XSWIFT_DASHBOARD_COLUMNS = ["Vehicle No", "Vehicle Status"]
REQUIRED_AT_DASHBOARD_COLUMNS = ["Company Name", "Vehicle"]


class ColumnValidationError(ValueError):
    """Raised by validate_inputs() with a structured, per-file report of
    missing/extra columns — str(exc) is a human-readable summary suitable
    for surfacing directly to the dashboard user."""

    def __init__(self, report: dict[str, dict[str, list[str]]]):
        self.report = report
        lines = ["Uploaded file(s) don't match the expected column structure:"]
        for file_label, info in report.items():
            if not info["missing"] and not info["extra"]:
                continue
            lines.append(f"\n[{file_label}]")
            if info["missing"]:
                lines.append(f"  MISSING required columns: {info['missing']}")
            if info["extra"]:
                lines.append(f"  Extra/unrecognized columns: {info['extra']}")
        super().__init__("\n".join(lines))


def _peek_columns(source, **read_kwargs) -> list[str]:
    """Reads just the header row of a CSV or XLSX source (path or bytes-like
    file object) — cheap, doesn't load the actual data rows."""
    if isinstance(source, io.BytesIO) or hasattr(source, "read"):
        source.seek(0)
    if "sheet_name" in read_kwargs:
        df = pd.read_excel(source, nrows=0, **read_kwargs)
    else:
        df = pd.read_csv(source, nrows=0, **read_kwargs)
    if isinstance(source, io.BytesIO) or hasattr(source, "seek"):
        source.seek(0)
    return list(df.columns)


def _check_columns(file_label: str, actual: list[str], required: list[str],
                    report: dict[str, dict[str, list[str]]]) -> None:
    missing = [c for c in required if c not in actual]
    extra = [c for c in actual if c not in required]
    report[file_label] = {"missing": missing, "extra": extra}


def validate_inputs(
    mtr_csv,
    consignment_xlsx,
    xswift_live_dashboard_xlsx=None,
    at_live_dashboard_xlsx=None,
) -> dict[str, dict[str, list[str]]]:
    """Checks each file's actual headers against what the pipeline requires.
    Raises ColumnValidationError (with the full per-file report) if any
    REQUIRED column is missing anywhere. Extra/unrecognized columns are
    reported but never raise — they're already handled (kept, not dropped)
    by reorder_to_final_layout().

    Accepts the same input shapes as run_in_memory()/run() — bytes, a
    Path, or an open file-like object.
    """
    report: dict[str, dict[str, list[str]]] = {}

    mtr_src = io.BytesIO(mtr_csv) if isinstance(mtr_csv, bytes) else mtr_csv
    _check_columns("Raw MTR CSV", _peek_columns(mtr_src), REQUIRED_MTR_CSV_COLUMNS, report)

    cons_src = io.BytesIO(consignment_xlsx) if isinstance(consignment_xlsx, bytes) else consignment_xlsx
    _check_columns(
        "AT Consignment Report",
        _peek_columns(cons_src, sheet_name="Consignment Report"),
        REQUIRED_CONSIGNMENT_COLUMNS, report,
    )

    if xswift_live_dashboard_xlsx:
        xswift_src = io.BytesIO(xswift_live_dashboard_xlsx) if isinstance(xswift_live_dashboard_xlsx, bytes) else xswift_live_dashboard_xlsx
        _check_columns(
            "XSwift Live Trip Dashboard",
            _peek_columns(xswift_src, sheet_name="Trip Dashboard", skiprows=2),
            REQUIRED_XSWIFT_DASHBOARD_COLUMNS, report,
        )

    if at_live_dashboard_xlsx:
        at_src = io.BytesIO(at_live_dashboard_xlsx) if isinstance(at_live_dashboard_xlsx, bytes) else at_live_dashboard_xlsx
        _check_columns(
            "AT Live Dashboard",
            _peek_columns(at_src, sheet_name="dashboard"),
            REQUIRED_AT_DASHBOARD_COLUMNS, report,
        )

    if any(info["missing"] for info in report.values()):
        raise ColumnValidationError(report)
    return report


# =============================================================================
# Helpers
# =============================================================================

def _is_blank(series: pd.Series) -> pd.Series:
    """True where a cell is NaN, empty string, or whitespace-only.

    CONFIRMED from the real raw data: blanks are stored as a literal
    single-space string " ", not a true empty cell. This must be the
    ONLY blank-check used throughout — do not swap for `.isna()`.
    """
    return series.isna() | series.astype(str).str.strip().eq("")


def _minutes_from_excel_duration(series: pd.Series) -> pd.Series:
    """Raw detention columns come through as Excel day-fraction floats
    (e.g. 0.0625 = 1.5 hours) when read from the source, OR as
    'HH:MM' strings when read from the CSV export (confirmed both
    forms appear across files). Normalize either to minutes (float).
    """
    numeric = pd.to_numeric(series, errors="coerce")
    # Day-fraction case: values are small floats (< 2, i.e. under 2 days)
    from_fraction = numeric * 24 * 60

    # HH:MM string case
    def _hhmm_to_minutes(val):
        if not isinstance(val, str) or ":" not in val:
            return np.nan
        try:
            h, m = val.split(":")[:2]
            return int(h) * 60 + int(m)
        except (ValueError, TypeError):
            return np.nan

    from_string = series.map(_hhmm_to_minutes)
    return from_fraction.fillna(from_string)


def _slab_from_minutes(minutes: pd.Series, edges=(10, 20, 30)) -> pd.Series:
    """0-10 min / 10-20 min / 20-30 min / Above 30 min, matching the
    exact label text confirmed in the real output file's filter list.
    """
    e1, e2, e3 = edges
    conditions = [
        minutes <= e1,
        (minutes > e1) & (minutes <= e2),
        (minutes > e2) & (minutes <= e3),
        minutes > e3,
    ]
    choices = [f"0-{e1} min", f"{e1}-{e2} min", f"{e2}-{e3} min", f"Above {e3} min"]
    return pd.Series(np.select(conditions, choices, default=""), index=minutes.index)


def _first_n_chars_match(a: pd.Series, b: pd.Series, n: int = 4) -> pd.Series:
    a_norm = a.astype(str).str.strip().str.upper().str[:n]
    b_norm = b.astype(str).str.strip().str.upper().str[:n]
    return a_norm == b_norm


# =============================================================================
# Step 1: Load reference data (Consignment Report + Primary Plants List)
# =============================================================================

def load_consignment_report_full(path: Path) -> pd.DataFrame:
    """Loads the FULL AT Consignment Report (all 67 columns), used both
    to build the lookup dicts (see build_consignment_lookups) and as the
    source for Trip Repush. Loaded once and reused for both, rather than
    reading this 60MB+ file twice.
    """
    log.info("Loading AT Consignment Report from %s (this is the slow step — large file)", path)
    df = pd.read_excel(path, sheet_name="Consignment Report", dtype=str, engine="openpyxl")
    log.info("Consignment Report loaded: %d rows x %d cols", *df.shape)
    return df


def build_consignment_lookups(consignment: pd.DataFrame) -> tuple[dict, dict]:
    """Build the two lookup dicts that replace the manual file's
    XLOOKUP / VLOOKUP formulas, from an already-loaded Consignment Report.

    Returns:
        city_code_to_destination: {City Code: Destination}
            (replaces XLOOKUP(Destination Code, Consignment!City Code, Consignment!Destination))
        sap_pgi_to_lead_dist: {SAP PGI No: SAP Lead Distance (Kms)}
            (replaces VLOOKUP(SAP PGI No, Consignment!J:AA, 18, 0))
    """
    # XLOOKUP takes the FIRST match — dict construction from a DataFrame
    # naturally keeps the LAST value per key, so we reverse before building
    # to replicate "first match wins".
    city_lookup_df = consignment[[CONS_CITY_CODE, CONS_DESTINATION]].dropna(subset=[CONS_CITY_CODE])
    city_code_to_destination = dict(
        zip(city_lookup_df[CONS_CITY_CODE][::-1], city_lookup_df[CONS_DESTINATION][::-1])
    )

    pgi_lookup_df = consignment[[CONS_SAP_PGI_NO, CONS_SAP_LEAD_DIST]].dropna(subset=[CONS_SAP_PGI_NO])
    sap_pgi_to_lead_dist = dict(
        zip(pgi_lookup_df[CONS_SAP_PGI_NO][::-1], pgi_lookup_df[CONS_SAP_LEAD_DIST][::-1])
    )

    return city_code_to_destination, sap_pgi_to_lead_dist


def load_primary_plant_codes(path: Path) -> set[str]:
    """The Primary Plants List has a messy two-column layout: Company
    name on the first row of each plant block, blank on subsequent
    rows, with every plant code (one company can have several) each
    on its own row. This flattens it to a plain set of codes.
    """
    log.info("Loading Primary Plants List from %s", path)
    df = pd.read_excel(path, sheet_name="Sheet1", dtype=str)

    # Real file has a couple of junk rows at the top (a slicer filter
    # display) before the real "Company" / "Plant Code" header — find
    # the header row dynamically instead of hardcoding a skiprows count.
    header_row_idx = df[df.iloc[:, 0] == PLANTS_COMPANY_COL].index
    if len(header_row_idx) == 0:
        # header was already used as columns by read_excel — proceed as-is
        code_col = df[PLANTS_CODE_COL] if PLANTS_CODE_COL in df.columns else df.iloc[:, 1]
    else:
        data = df.iloc[header_row_idx[0] + 1:]
        code_col = data.iloc[:, 1]

    codes = set(code_col.dropna().astype(str).str.strip())
    codes.discard("")
    log.info("Loaded %d primary plant codes", len(codes))
    return codes


def load_primary_plant_companies(path: Path) -> set[str]:
    """Same source file as load_primary_plant_codes(), but returns the
    normalized COMPANY NAMES instead of codes — needed for
    run_task1_mapping(), which matches against AT's "Company Name" field
    (a text name, not a code). Normalized to upper-case for matching
    against AT's `_UTCL(P)`/`_UTCL(T)`-suffixed company names.
    """
    log.info("Loading Primary Plants List (company names) from %s", path)
    df = pd.read_excel(path, sheet_name="Sheet1", dtype=str)
    header_row_idx = df[df.iloc[:, 0] == PLANTS_COMPANY_COL].index
    if len(header_row_idx) == 0:
        company_col = df[PLANTS_COMPANY_COL] if PLANTS_COMPANY_COL in df.columns else df.iloc[:, 0]
    else:
        data = df.iloc[header_row_idx[0] + 1:]
        company_col = data.iloc[:, 0]

    companies = set(company_col.dropna().astype(str).str.strip().str.upper())
    companies.discard("")
    log.info("Loaded %d primary plant company names", len(companies))
    return companies


# =============================================================================
# Step 2: Load raw MTR
# =============================================================================

def load_raw_mtr(path: Path, chunksize: int | None = None) -> pd.DataFrame:
    log.info("Loading raw MTR CSV from %s", path)
    read_kwargs = dict(dtype=str, keep_default_na=False, na_values=[""])
    if chunksize:
        chunks = []
        for i, chunk in enumerate(pd.read_csv(path, chunksize=chunksize, **read_kwargs)):
            chunks.append(chunk)
            log.info("  read chunk %d (%d rows so far)", i + 1, sum(len(c) for c in chunks))
        df = pd.concat(chunks, ignore_index=True)
    else:
        df = pd.read_csv(path, **read_kwargs)
    log.info("Raw MTR loaded: %d rows x %d cols", *df.shape)
    return df


# =============================================================================
# Step 3: Build every derived column (the actual business logic)
# =============================================================================

def build_analysis_columns(mtr: pd.DataFrame, cfg: Config,
                            city_code_to_destination: dict,
                            sap_pgi_to_lead_dist: dict,
                            primary_plant_codes: set[str]) -> pd.DataFrame:
    df = mtr.copy()
    n = len(df)
    log.info("Building analysis columns for %d rows", n)

    is_primary = df[COL_PLANT_CODE].isin(primary_plant_codes)

    # --- Date and time = INT(PGI Date & Time) i.e. date part only ---
    pgi_dt = pd.to_datetime(df[COL_PGI_DATETIME], errors="coerce", format="mixed")
    df[NEW_DATE_AND_TIME] = pgi_dt.dt.date

    # --- Transporter Remark ---
    df[NEW_TRANSPORTER_REMARK] = np.where(
        _is_blank(df[COL_TRANSPORTER_NAME]), "Not Available", "Available"
    )

    # --- Zone Remark: primary plant + Zone blank -> "Zone enable for it" ---
    df[NEW_ZONE_REMARK] = np.where(
        is_primary & _is_blank(df[COL_ZONE]), "Zone enable for it", ""
    )

    # --- Yard Detention Slab ---
    yard_in_blank = _is_blank(df[COL_YARD_IN])
    yard_out_blank = _is_blank(df[COL_YARD_OUT])
    yard_minutes = _minutes_from_excel_duration(df[COL_YARD_DETENTION])
    yard_zero = df[COL_YARD_DETENTION].astype(str).str.strip().isin(["0", "0:00", "00:00", "0.0"])

    yard_slab = pd.Series("", index=df.index)
    yard_slab = yard_slab.mask(yard_in_blank & yard_out_blank, "In out both blank")
    yard_slab = yard_slab.mask(yard_in_blank & ~yard_out_blank, "Problem")
    yard_slab = yard_slab.mask(~yard_in_blank & yard_out_blank, "Vehicle still in yard")
    both_present = ~yard_in_blank & ~yard_out_blank
    yard_slab = yard_slab.mask(both_present & yard_zero, "Not available")
    remaining = both_present & ~yard_zero & (yard_slab == "")
    yard_slab = yard_slab.mask(remaining, _slab_from_minutes(yard_minutes, cfg.detention_slab_edges_min))
    df[NEW_YARD_DETENTION_SLAB] = yard_slab

    # --- Plant detention Slab ---
    plant_exit_blank = _is_blank(df[COL_PLANT_EXIT])
    plant_entry_blank = _is_blank(df[COL_PLANT_ENTRY])
    plant_detention_blank = _is_blank(df[COL_PLANT_DETENTION])
    plant_minutes = _minutes_from_excel_duration(df[COL_PLANT_DETENTION])
    plant_detention_zero = (~plant_detention_blank) & (plant_minutes.fillna(-1) == 0)

    plant_slab = pd.Series("", index=df.index)
    plant_slab = plant_slab.mask(plant_exit_blank & ~plant_entry_blank, "vehicle still in plant")
    plant_slab = plant_slab.mask(plant_detention_zero & (plant_slab == ""), "Loading not merged")
    plant_slab = plant_slab.mask(plant_detention_blank & (plant_slab == ""), "Issue")
    remaining = (plant_slab == "") & ~plant_detention_blank & ~plant_detention_zero
    plant_slab = plant_slab.mask(remaining, _slab_from_minutes(plant_minutes, cfg.detention_slab_edges_min))
    df[NEW_PLANT_DETENTION_SLAB] = plant_slab

    # --- AT destination name (XLOOKUP replacement) + Dest. Match ---
    # NOTE: output is the literal strings "TRUE"/"FALSE"/"NA", not Python
    # booleans. Earlier version used np.where(cond, "NA", bool_array), which
    # under pandas' string-dtype inference silently coerces True/False into
    # the STRINGS "True"/"False" once mixed with "NA" in the same column —
    # confirmed via test_pipeline.py. Made explicit here instead of relying
    # on that implicit (and confusing) coercion.
    df[NEW_AT_DEST_NAME] = df[COL_DEST_CODE].map(city_code_to_destination)
    at_dest_is_na = df[NEW_AT_DEST_NAME].isna()
    df[NEW_AT_DEST_NAME] = df[NEW_AT_DEST_NAME].fillna("#N/A")
    name_match = _first_n_chars_match(df[COL_DESTINATION], df[NEW_AT_DEST_NAME])
    df[NEW_DEST_MATCH] = np.select(
        [at_dest_is_na, name_match], ["NA", "TRUE"], default="FALSE"
    )

    # --- Destination detention slab (gated by Stamp Status) ---
    in_stamp_scope = df[COL_STAMP_STATUS].isin(STAMP_STATUSES_FOR_CHECKS)
    dest_exit_blank = _is_blank(df[COL_DEST_EXIT])
    dest_entry_blank = _is_blank(df[COL_DEST_ENTRY])
    dest_detention_blank = _is_blank(df[COL_DEST_DETENTION])
    dest_minutes = _minutes_from_excel_duration(df[COL_DEST_DETENTION])
    dest_detention_zero_or_null = dest_detention_blank | (dest_minutes.fillna(-1) == 0)

    dest_slab = pd.Series("", index=df.index)
    issue_mask = in_stamp_scope & dest_exit_blank & ~dest_detention_zero_or_null
    dest_slab = dest_slab.mask(issue_mask, "Issue")
    still_at_site_mask = in_stamp_scope & dest_detention_zero_or_null & dest_exit_blank & (dest_slab == "")
    dest_slab = dest_slab.mask(still_at_site_mask, "Vehicle still at site")
    both_present_mask = in_stamp_scope & ~dest_exit_blank & ~dest_entry_blank & (dest_slab == "")
    dest_slab = dest_slab.mask(
        both_present_mask, _slab_from_minutes(dest_minutes, cfg.detention_slab_edges_min)
    )
    df[NEW_DEST_DETENTION_SLAB] = dest_slab

    # --- AT SAP lead distance (VLOOKUP replacement) + Match ---
    df[NEW_AT_SAP_LEAD_DIST] = df[COL_SAP_PGI_NO].map(sap_pgi_to_lead_dist)
    at_lead_is_na = df[NEW_AT_SAP_LEAD_DIST].isna()
    df[NEW_AT_SAP_LEAD_DIST] = df[NEW_AT_SAP_LEAD_DIST].fillna("#N/A")

    sap_lead_numeric = pd.to_numeric(df[COL_SAP_LEAD_DIST], errors="coerce")
    at_lead_numeric = pd.to_numeric(df[NEW_AT_SAP_LEAD_DIST], errors="coerce")
    lead_match = np.isclose(sap_lead_numeric, at_lead_numeric, equal_nan=False)
    df[NEW_MATCH] = np.select(
        [at_lead_is_na, lead_match], ["NA", "TRUE"], default="FALSE"
    )

    # --- AI check, SAP-AI, SAP-AI Remark (gated by Stamp Status) ---
    ai_dist_numeric = pd.to_numeric(df[COL_AI_REPAIRED_DIST], errors="coerce")
    ai_blank_or_zero = _is_blank(df[COL_AI_REPAIRED_DIST]) | (ai_dist_numeric.fillna(-1) == 0)

    ai_check = pd.Series("", index=df.index)
    ai_check = ai_check.mask(in_stamp_scope & ai_blank_or_zero, "Not Available")
    ai_check = ai_check.mask(in_stamp_scope & ~ai_blank_or_zero, "Available")
    df[NEW_AI_CHECK] = ai_check

    sap_ai_diff = sap_lead_numeric - ai_dist_numeric
    df[NEW_SAP_AI] = np.where(in_stamp_scope, sap_ai_diff, np.nan)

    lo, hi = cfg.sap_ai_ok_band
    sap_ai_remark = pd.Series("", index=df.index)
    sap_ai_remark = sap_ai_remark.mask(in_stamp_scope & (ai_dist_numeric.fillna(-1) == 0), "AI is blank")
    ok_band = in_stamp_scope & sap_ai_diff.between(lo, hi) & (sap_ai_remark == "")
    sap_ai_remark = sap_ai_remark.mask(ok_band, f"{lo} to {hi}")
    high = in_stamp_scope & (sap_ai_diff < lo) & (sap_ai_remark == "")
    sap_ai_remark = sap_ai_remark.mask(high, "AI usages is high")
    low = in_stamp_scope & (sap_ai_diff > hi) & (sap_ai_remark == "")
    sap_ai_remark = sap_ai_remark.mask(low, "AI usages is low")
    df[NEW_SAP_AI_REMARK] = sap_ai_remark

    log.info("Analysis columns built.")
    return df


# =============================================================================
# Step 4: Reorder columns to match the confirmed real output layout (A..BA)
# =============================================================================

def reorder_to_final_layout(df: pd.DataFrame) -> pd.DataFrame:
    """Column order confirmed by extracting the real
    'MTR Analysis - 20 July.xlsx' file's XML directly. If the raw MTR
    export ever changes its column set, this list needs updating —
    it is NOT derived automatically on purpose, so a missing/renamed
    source column fails loudly (KeyError) instead of silently
    reordering wrong.

    NOTE (2026-07-20): the manual "20 July" report dropped
    'Mother Geofence Start Time', 'Mother Geofence End Time', and
    'Mother Geofence Detention' — confirmed with the business contact
    this was a ONE-OFF omission for that day's report only, NOT a
    permanent change. This pipeline deliberately KEEPS all three,
    placed in their raw-MTR position (right after 'Geofence Hit/miss',
    before 'Billing Status') — do not remove them again without
    re-confirming.
    """
    final_order = [
        "Trip ID", "Vehicle No.", "Vehicle Type", "SAP PGI No", "PGI Date & Time",
        NEW_DATE_AND_TIME, "SAP Order No", "DI No", "Transporter Name ",
        NEW_TRANSPORTER_REMARK, "Transporter Code", "Zone", "Yard IN", "Yard Out",
        "Yard detention", NEW_YARD_DETENTION_SLAB, "Plant name", NEW_ZONE_REMARK,
        "Plant Code", "Plant Entry", "Plant Exit", "Plant Detention",
        NEW_PLANT_DETENTION_SLAB, "Destination Code", "Destination", NEW_AT_DEST_NAME,
        NEW_DEST_MATCH, "Customer Name", "Dest Entry Time", "Dest Exit Time",
        "Dest Detention", NEW_DEST_DETENTION_SLAB, "Destination Proximity End Time",
        "Destination Ageing", "Onward Duration", "Customer Segment", "Compliance Status",
        "Depot", "Route Name", "Halt", "Onward Status", "Stamp Status", "Reject Reason",
        "Sap Lead Dist", NEW_AT_SAP_LEAD_DIST, NEW_MATCH, "GPS Distance",
        "AI Repaired Distance", NEW_AI_CHECK, NEW_SAP_AI, NEW_SAP_AI_REMARK,
        "Geofence Hit/miss",
        "Mother Geofence Start Time", "Mother Geofence End Time", "Mother Geofence Detention",
        "Billing Status",
    ]
    missing = [c for c in final_order if c not in df.columns]
    if missing:
        raise KeyError(
            f"Expected columns missing from the built DataFrame: {missing}. "
            "Raw MTR export column names may have changed — update COL_* "
            "constants and final_order in reorder_to_final_layout()."
        )
    extra = [c for c in df.columns if c not in final_order]
    if extra:
        log.warning("Columns present but not in the confirmed layout (kept at the end): %s", extra)
    return df[final_order + extra]


# =============================================================================
# Step 5: Pivots (Sheet1 / Sheet2) — CONFIRMED, extracted from real file XML
# =============================================================================
# The real output file has 7 PivotTables (not 2 — "Sheet1"/"Sheet2" are just
# the two worksheets they're placed on: 2 pivots on "Sheet2", 5 on "Sheet1").
# Field layout (rows/columns/filters/values) was extracted directly from
# xl/pivotTables/pivotTable{1..7}.xml — this is NOT a guess or placeholder.
#
# DESIGN DECISION (confirm with business contact if it matters to them):
# the real file's "filter" fields are interactive Excel slicers a person
# toggles by hand. A pandas-generated static report can't replicate that
# interactivity without building real Excel PivotTable objects (much more
# engineering). This implementation outputs the full crosstab with NO
# filter applied (i.e. all data, same as opening the real pivot with every
# slicer set to "All"). If Anchal actually relies on toggling these filters
# day-to-day, that's a reason to revisit this decision — don't assume.

PIVOT_VALUE_FIELD = "Trip ID"

PIVOT_DEFINITIONS = [
    # (sheet, title, row_fields, col_field)
    ("Sheet2", "Yard Detention Slab by Plant", ["Plant name"], NEW_YARD_DETENTION_SLAB),
    ("Sheet2", "SAP-AI Remark by Zone/Plant", ["Zone", "Plant name"], NEW_SAP_AI_REMARK),
    ("Sheet1", "Destination Detention Slab by Destination", ["Destination", "Destination Code"], NEW_DEST_DETENTION_SLAB),
    ("Sheet1", "Dest. Match by Destination", ["Destination", "Destination Code"], NEW_DEST_MATCH),
    ("Sheet1", "Plant Detention Slab by Plant", ["Plant name"], NEW_PLANT_DETENTION_SLAB),
    ("Sheet1", "Zone Remark by Plant", ["Plant name"], NEW_ZONE_REMARK),
    ("Sheet1", "Transporter Remark by Plant/Transporter", ["Plant name", "Transporter Name ", "Transporter Code"], NEW_TRANSPORTER_REMARK),
]


def build_pivot(df: pd.DataFrame, row_fields: list[str], col_field: str) -> pd.DataFrame:
    """Count of Trip ID, grouped by row_fields (rows) x col_field (columns).
    Matches every one of the 7 real pivots' shape: count(Trip ID) by
    row-field(s) x one remark column. No filter applied (see design note
    above) — equivalent to every slicer set to "All" in the real file.
    """
    return (
        df.groupby(row_fields + [col_field])[PIVOT_VALUE_FIELD]
        .count()
        .unstack(col_field, fill_value=0)
    )


def build_all_pivots(df: pd.DataFrame) -> dict[str, list[tuple[str, pd.DataFrame]]]:
    """Returns {sheet_name: [(title, pivot_df), ...]} — preserves the real
    file's grouping of multiple pivot tables stacked on the same two sheets.
    """
    by_sheet: dict[str, list[tuple[str, pd.DataFrame]]] = {"Sheet1": [], "Sheet2": []}
    for sheet, title, row_fields, col_field in PIVOT_DEFINITIONS:
        pivot = build_pivot(df, row_fields, col_field)
        by_sheet[sheet].append((title, pivot))
    return by_sheet


# =============================================================================
# Step 6: Task 1 — AT <-> XSwift trip mapping (separate, smaller pipeline)
# =============================================================================

def run_task1_trip_repush(consignment: pd.DataFrame, mtr: pd.DataFrame,
                           primary_plant_codes: set[str]) -> pd.DataFrame:
    """CONFIRMED LOGIC (2026-07-20) — reproduces Trip Repush - <date>.xlsx.

    Verified empirically: every SAP PGI No in the real Trip Repush output
    was checked against all 281,349 rows of a real MTR export with ZERO
    overlap found — i.e. "Trip Repush" = AT Consignment Report rows whose
    SAP PGI No does not exist anywhere in XSwift MTR, restricted to
    primary plants (confirmed by the business contact — every plant in
    the real output was a primary plant). Output = the full original
    Consignment Report row (all columns unchanged), not a reduced set.

    Unlike run_task1_mapping() below (still ON HOLD), this filter is
    fully confirmed — Consignment Report has a real `Plant Code` column,
    so this matches by exact code against the Primary Plants List rather
    than the fuzzy company-name matching Mapping issue would need.
    """
    log.info("Running Task 1: Trip Repush")
    is_primary = consignment[CONS_PLANT_CODE].isin(primary_plant_codes)
    mtr_pgi_set = set(mtr[COL_SAP_PGI_NO].dropna().astype(str).str.strip())
    missing_from_xswift = ~consignment[CONS_SAP_PGI_NO].astype(str).str.strip().isin(mtr_pgi_set)

    result = consignment[is_primary & missing_from_xswift].copy()
    log.info("Trip Repush: %d rows (primary plant + not in XSwift MTR)", len(result))
    return result


def run_task1_mapping(cfg: Config, primary_plant_companies: set[str]) -> dict[str, pd.DataFrame]:
    """CONFIRMED — UN-HELD 2026-07-20. Reproduces the candidate list behind
    Mapping_issue_-_20_July.xlsx: vehicles present on one live dashboard
    (AT or XSwift) but not the other.

    VALIDATED against the real output file (both sheets, 100% recall —
    every real vehicle is present in this function's output; the
    remaining gap between candidate count and the real ~70/~10 rows is
    precision, not a missing vehicle):

    - XSwift's Live Trip Dashboard (`xswift_live_dashboard_xlsx`) does
      NOT need a plant-name filter — per the business contact's own
      notes, XSwift only carries primary-plant data at all right now, so
      filtering it by plant name is not just unnecessary but actively
      wrong (many rows have a blank `Vehicle Reg Plant Name` field even
      though the vehicle IS a real primary-plant vehicle — filtering on
      that field was silently dropping 23/70 real answers).
    - AT's Live Trip Dashboard (`at_live_dashboard_xlsx`) DOES need the
      primary-plant filter — AT's platform spans many non-UTCL client
      companies. Match by normalized `Company Name` against the Primary
      Plants List (strip `_UTCL(P)`/`_UTCL(T)` suffixes).
    - "Not in AT" additionally filtered to XSwift `Vehicle Status !=
      "Online"` (i.e. Offline/Idle) — cuts candidates from 391 to 329
      with NO loss of recall (confirmed: all 70 real vehicles have
      Vehicle Status in {Offline, Idle}).
    - "Not in Swift" has NO further confirmed precision filter — the AT
      `Status` field (Idle/Unreachable/Moving) does not cleanly separate
      the 10 real answers from the other candidates (restricting to
      "Idle" only drops 2 of the 10 real vehicles). Left as the full
      180-candidate list; narrowing further needs input from the
      business contact rather than a guessed threshold.

    Both outputs are therefore SAFE (no false negatives against the real
    file) but WIDER than the real file (some false positives) — treat as
    a candidate list for review, not a guaranteed exact match, unless/
    until further precision rules are confirmed.
    """
    if not cfg.xswift_live_dashboard_xlsx or not cfg.at_live_dashboard_xlsx:
        log.warning("Task 1 inputs not provided — skipping run_task1_mapping().")
        return {}

    log.info("Running Task 1: Mapping issue (validated, 100%% recall against real output)")

    def _normalize_plant(name) -> str:
        if not isinstance(name, str):
            return ""
        n = name.upper()
        for suffix in ["_UTCL(P)", "_UTCL(T)", "_UTCL"]:
            n = n.replace(suffix, "")
        return n.strip()

    # XSwift side: skiprows=2 skips the "Name:"/"Report:" banner rows before
    # the real header (confirmed from the real file's raw structure).
    xswift_df = pd.read_excel(
        cfg.xswift_live_dashboard_xlsx, sheet_name="Trip Dashboard", skiprows=2, dtype=str
    )
    at_df = pd.read_excel(cfg.at_live_dashboard_xlsx, sheet_name="dashboard", dtype=str)

    # NO plant filter on XSwift side — see docstring.
    xswift_vehicles = set(xswift_df["Vehicle No"].dropna().astype(str).str.strip())

    # Primary-plant filter on AT side only.
    at_df = at_df[at_df["Company Name"].map(_normalize_plant).isin(primary_plant_companies)]
    at_vehicles = set(at_df["Vehicle"].dropna().astype(str).str.strip())

    not_in_at_mask = xswift_df["Vehicle No"].astype(str).str.strip().isin(xswift_vehicles - at_vehicles)
    not_in_at = xswift_df[not_in_at_mask & (xswift_df["Vehicle Status"] != "Online")]

    not_in_swift = at_df[at_df["Vehicle"].astype(str).str.strip().isin(at_vehicles - xswift_vehicles)]

    log.info("Not in AT: %d candidates (100%% recall vs. real file, precision not fully resolved)", len(not_in_at))
    log.info("Not in Swift: %d candidates (100%% recall vs. real file, precision not fully resolved)", len(not_in_swift))

    return {"Not in AT": not_in_at, "Not in Swift": not_in_swift}


# =============================================================================
# Output writers
# =============================================================================

# Columns confirmed (from the real file's styles.xml, header row s="6" —
# fill index 33 = FFFF00, pure yellow) to have a yellow header background.
# NOTE: "Date and time" is a derived column too but is NOT in this list —
# confirmed its header uses the normal style (s="1"), not yellow.
YELLOW_HEADER_COLUMNS = {
    NEW_TRANSPORTER_REMARK, NEW_YARD_DETENTION_SLAB, NEW_ZONE_REMARK,
    NEW_PLANT_DETENTION_SLAB, NEW_AT_DEST_NAME, NEW_DEST_MATCH,
    NEW_DEST_DETENTION_SLAB, NEW_AT_SAP_LEAD_DIST, NEW_MATCH,
    NEW_AI_CHECK, NEW_SAP_AI, NEW_SAP_AI_REMARK,
}

# Columns confirmed (from real data-row cell styles) to use Excel's
# datetime number format (numFmtId 22, "m/d/yy h:mm").
DATETIME_FORMAT_COLUMNS = {"PGI Date & Time", NEW_DATE_AND_TIME, "Plant Entry", "Plant Exit"}

# Column confirmed to use Excel's time-duration format (numFmtId 20, "h:mm").
DURATION_FORMAT_COLUMNS = {"Plant Detention"}


def write_xlsx(main_sheet_name: str, main_df: pd.DataFrame,
                pivots_by_sheet: dict[str, list[tuple[str, pd.DataFrame]]],
                path: Path) -> None:
    """CORRECTED 2026-07-20 — this function was previously named
    write_xlsx_streaming() and used xlsxwriter's `constant_memory: True`
    option. That option SILENTLY CORRUPTS DATA when combined with
    pandas' `DataFrame.to_excel()`: confirmed via isolated repro — even
    a trivial 3-column, 100-row DataFrame with constant_memory=True
    writes only the FIRST column correctly and leaves every other
    column blank, with no error or warning. This shipped a real MTR
    Analysis run with a 2.5MB xlsx (vs. the correct ~200MB) containing
    only Trip ID values before it was caught. constant_memory is now
    NOT used anywhere in this file — do not re-add it.

    Practical consequence: writing a 281k-row x 53-column sheet without
    constant_memory holds more of the workbook in memory while writing
    (openpyxl/xlsxwriter's normal mode), which is slower and more
    memory-hungry than the (broken) streaming mode was. This is fine on
    a machine with a few GB of free RAM — which is what the company
    server this now runs on provides; it would NOT have been fine on
    Render's free 512MB tier, but that's no longer where this executes.

    Pivot sheets ("Sheet1"/"Sheet2") each hold multiple titled pivot
    tables stacked vertically with a blank row between them, matching
    the real file's layout (multiple PivotTables per worksheet).

    FORMATTING (2026-07-20, confirmed from the real file's styles.xml —
    not a guess): the 12 columns in YELLOW_HEADER_COLUMNS get a yellow
    (#FFFF00) header background, matching the business contact's own
    notes ("all the columns highlighted with yellow header added"). Date
    columns get proper Excel datetime formatting; Plant Detention gets
    Excel's time-duration format. Every other column is left as Excel's
    general default — do not add formatting beyond what's listed here
    without re-confirming against the real file first.
    """
    log.info("Writing XLSX to %s", path)
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        workbook = writer.book

        for sheet_name, pivots in pivots_by_sheet.items():
            row_cursor = 0
            for title, pivot_df in pivots:
                title_df = pd.DataFrame([[title]])
                title_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False,
                                   startrow=row_cursor)
                row_cursor += 1
                pivot_df.to_excel(writer, sheet_name=sheet_name, startrow=row_cursor)
                row_cursor += len(pivot_df) + pivot_df.columns.nlevels + 3
            if not pivots:
                pd.DataFrame().to_excel(writer, sheet_name=sheet_name)

        main_sheet_name = main_sheet_name[:31]
        main_df.to_excel(writer, sheet_name=main_sheet_name, index=False)
        worksheet = writer.sheets[main_sheet_name]

        yellow_header_format = workbook.add_format({"bg_color": "#FFFF00", "bold": False})
        datetime_format = workbook.add_format({"num_format": "m/d/yy h:mm"})
        duration_format = workbook.add_format({"num_format": "h:mm"})

        for col_idx, col_name in enumerate(main_df.columns):
            if col_name in YELLOW_HEADER_COLUMNS:
                worksheet.write(0, col_idx, col_name, yellow_header_format)
            if col_name in DATETIME_FORMAT_COLUMNS:
                worksheet.set_column(col_idx, col_idx, None, datetime_format)
            elif col_name in DURATION_FORMAT_COLUMNS:
                worksheet.set_column(col_idx, col_idx, None, duration_format)

    # path may be a real file path (CLI/disk mode) or an in-memory buffer
    # (see run_in_memory() below) — size it either way without touching
    # any of the actual write logic above.
    size_bytes = path.stat().st_size if hasattr(path, "stat") else path.tell()
    log.info("XLSX written: %.1f MB", size_bytes / 1e6)


# =============================================================================
# Main
# =============================================================================

def run(cfg: Config) -> None:
    """OUTPUT FORMAT CHANGED (2026-07-20, explicit business requirement):
    xlsx ONLY, one file per output, no CSV. The business contact needs
    the output to be "same to same" as the 3 real files she works with —
    CSV was never part of that (it was added earlier purely as a
    cheaper/faster intermediate while chasing the constant_memory bug,
    which is now fixed — CSV was a workaround for a problem that no
    longer exists). Do not reintroduce CSV output without being asked.
    """
    cfg.output_dir.mkdir(parents=True, exist_ok=True)

    validate_inputs(
        cfg.mtr_csv, cfg.consignment_xlsx,
        cfg.xswift_live_dashboard_xlsx, cfg.at_live_dashboard_xlsx,
    )

    primary_plant_codes = load_primary_plant_codes(cfg.primary_plants_xlsx)
    primary_plant_codes = (primary_plant_codes - cfg.task1_exclude_plant_codes) | cfg.task1_include_extra_plant_codes

    # Load the full Consignment Report once — used for both the MTR Analysis
    # lookups AND Trip Repush (avoids reading this 60MB+ file twice).
    consignment_full = load_consignment_report_full(cfg.consignment_xlsx)
    city_code_to_destination, sap_pgi_to_lead_dist = build_consignment_lookups(consignment_full)

    mtr = load_raw_mtr(cfg.mtr_csv, chunksize=cfg.csv_chunksize)

    # --- Output 1: "Output Final - MTR Analysis" ---
    analyzed = build_analysis_columns(
        mtr, cfg, city_code_to_destination, sap_pgi_to_lead_dist, primary_plant_codes
    )
    analyzed = reorder_to_final_layout(analyzed)
    pivots_by_sheet = build_all_pivots(analyzed)
    write_xlsx(f"mtr - {cfg.run_date_label}", analyzed, pivots_by_sheet,
               cfg.output_dir / f"MTR_Analysis_-_{cfg.run_date_label}.xlsx")

    # --- Output 2: "Output Trip creation - Trip Repush" ---
    # Real file keeps one tab per run-date in an accumulating workbook
    # (e.g. "18 June", "20 July" both present). This writes a fresh
    # single-tab file per run instead of appending — if the accumulating
    # history matters in practice, extend this to open the existing
    # workbook with openpyxl (NOT xlsxwriter, which can't append to an
    # existing file) and add a new tab rather than overwrite.
    trip_repush = run_task1_trip_repush(consignment_full, mtr, primary_plant_codes)
    with pd.ExcelWriter(cfg.output_dir / f"Trip_Repush_-_{cfg.run_date_label}.xlsx", engine="xlsxwriter") as writer:
        trip_repush.to_excel(writer, sheet_name=cfg.run_date_label[:31], index=False)

    # --- Output 3: "Output Mapping issue" ---
    # ONE workbook, TWO tabs ("Not in AT", "Not in Swift") — matches the
    # real file's structure. Produces a WIDER candidate list than the
    # real file (100% recall confirmed, precision not fully resolved —
    # see run_task1_mapping() docstring); this is a deliberate, documented
    # tradeoff, not a bug.
    if cfg.xswift_live_dashboard_xlsx and cfg.at_live_dashboard_xlsx:
        primary_plant_companies = load_primary_plant_companies(cfg.primary_plants_xlsx)
        task1_results = run_task1_mapping(cfg, primary_plant_companies)
        if task1_results:
            with pd.ExcelWriter(cfg.output_dir / f"Mapping_issue_-_{cfg.run_date_label}.xlsx", engine="xlsxwriter") as writer:
                for name, df in task1_results.items():
                    df.to_excel(writer, sheet_name=name[:31], index=False)
    else:
        log.info("Mapping issue inputs not provided — skipping (pass --xswift-live-dashboard-xlsx and --at-live-dashboard-xlsx to run it).")

    log.info("Pipeline complete. 3 xlsx files written to %s", cfg.output_dir)


def run_in_memory(
    mtr_csv: bytes,
    consignment_xlsx: bytes,
    primary_plants_xlsx: bytes,
    run_date_label: str,
    xswift_live_dashboard_xlsx: bytes | None = None,
    at_live_dashboard_xlsx: bytes | None = None,
    task1_exclude_plant_codes: set[str] | None = None,
    task1_include_extra_plant_codes: set[str] | None = None,
) -> dict[str, bytes]:
    """Same pipeline, same inputs, same 3 xlsx outputs as run() — every
    input/output is raw bytes in RAM instead of a path on disk (needed
    because this now runs on a shared company server that must not write
    anything to disk, not even temporarily).

    Mirrors run()'s exact sequence of calls. Reuses build_analysis_columns
    / run_task1_trip_repush / run_task1_mapping / reorder_to_final_layout /
    build_all_pivots / write_xlsx exactly as-is — none of those (the
    functions carrying the actual confirmed business logic and formatting)
    are modified here. Returns {output_filename: file_bytes}.
    """
    log.info("Running pipeline in-memory (no disk writes)")

    validate_inputs(mtr_csv, consignment_xlsx, xswift_live_dashboard_xlsx, at_live_dashboard_xlsx)

    primary_plant_codes = load_primary_plant_codes(io.BytesIO(primary_plants_xlsx))
    primary_plant_codes = (
        (primary_plant_codes - (task1_exclude_plant_codes or set()))
        | (task1_include_extra_plant_codes or set())
    )

    consignment_full = load_consignment_report_full(io.BytesIO(consignment_xlsx))
    city_code_to_destination, sap_pgi_to_lead_dist = build_consignment_lookups(consignment_full)

    mtr = load_raw_mtr(io.BytesIO(mtr_csv))

    # Config only carries business-rule thresholds through to
    # build_analysis_columns() here — its path fields are unused.
    cfg = Config(
        mtr_csv=Path("."), consignment_xlsx=Path("."), primary_plants_xlsx=Path("."),
        run_date_label=run_date_label,
    )

    outputs: dict[str, bytes] = {}

    # --- Output 1: "Output Final - MTR Analysis" ---
    analyzed = build_analysis_columns(
        mtr, cfg, city_code_to_destination, sap_pgi_to_lead_dist, primary_plant_codes
    )
    analyzed = reorder_to_final_layout(analyzed)
    pivots_by_sheet = build_all_pivots(analyzed)
    buf = io.BytesIO()
    write_xlsx(f"mtr - {run_date_label}", analyzed, pivots_by_sheet, buf)
    outputs[f"MTR_Analysis_-_{run_date_label}.xlsx"] = buf.getvalue()

    # --- Output 2: "Output Trip creation - Trip Repush" ---
    trip_repush = run_task1_trip_repush(consignment_full, mtr, primary_plant_codes)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        trip_repush.to_excel(writer, sheet_name=run_date_label[:31], index=False)
    outputs[f"Trip_Repush_-_{run_date_label}.xlsx"] = buf.getvalue()

    # --- Output 3: "Output Mapping issue" ---
    if xswift_live_dashboard_xlsx and at_live_dashboard_xlsx:
        primary_plant_companies = load_primary_plant_companies(io.BytesIO(primary_plants_xlsx))
        cfg.xswift_live_dashboard_xlsx = io.BytesIO(xswift_live_dashboard_xlsx)
        cfg.at_live_dashboard_xlsx = io.BytesIO(at_live_dashboard_xlsx)
        task1_results = run_task1_mapping(cfg, primary_plant_companies)
        if task1_results:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                for name, df in task1_results.items():
                    df.to_excel(writer, sheet_name=name[:31], index=False)
            outputs[f"Mapping_issue_-_{run_date_label}.xlsx"] = buf.getvalue()
    else:
        log.info("Mapping issue inputs not provided — skipping.")

    log.info("Pipeline complete (in-memory). %d xlsx files.", len(outputs))
    return outputs


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="MTR Analysis Pipeline")
    parser.add_argument("--mtr-csv", type=Path, required=True)
    parser.add_argument("--consignment-xlsx", type=Path, required=True)
    parser.add_argument("--primary-plants-xlsx", type=Path, required=True)
    parser.add_argument("--xswift-live-dashboard-xlsx", type=Path, default=None, help="XSwift Live Trip Dashboard export (needed for Mapping issue output)")
    parser.add_argument("--at-live-dashboard-xlsx", type=Path, default=None, help="AT Live Trip Dashboard export (needed for Mapping issue output)")
    parser.add_argument("--output-dir", type=Path, default=Path("./output"))
    parser.add_argument("--run-date-label", type=str, default="output", help="e.g. '20 July' — used in every output filename and as the Trip Repush tab name")
    parser.add_argument("--csv-chunksize", type=int, default=None, help="Read the raw MTR input CSV in chunks (for low-RAM machines) — unrelated to output format, which is always xlsx")
    args = parser.parse_args()

    config = Config(
        mtr_csv=args.mtr_csv,
        consignment_xlsx=args.consignment_xlsx,
        primary_plants_xlsx=args.primary_plants_xlsx,
        xswift_live_dashboard_xlsx=args.xswift_live_dashboard_xlsx,
        at_live_dashboard_xlsx=args.at_live_dashboard_xlsx,
        output_dir=args.output_dir,
        run_date_label=args.run_date_label,
        csv_chunksize=args.csv_chunksize,
    )

    try:
        run(config)
    except Exception:
        log.exception("Pipeline failed")
        sys.exit(1)
