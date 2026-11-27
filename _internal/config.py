"""
config.py
---------
All the things you'll likely need to tweak live here, in one place.
"""

import os

# ---------------------------------------------------------------------------
# GCP settings
# ---------------------------------------------------------------------------
PROJECT_ID = "ingka-fms-ccsync-prod"

# Path to the service account key file used for `gcloud auth activate-service-account`.
# You can also override this from the Streamlit sidebar at runtime.
DEFAULT_SERVICE_ACCOUNT_KEY_PATH = (
    r"C:/Users/SASIN88/Documents/FMS/pubsub/serviceAccountKey/ingka-fms-prod.json"
)

# Topic used to trigger the baseload
BASELOAD_TOPIC = "hcl_cutover_trigger_topic"

# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------
# Full catalog kept here so re-enabling the other 4 reports later is just
# adding their key back to ENABLED_REPORTS - no rewrite needed.
REPORTS = {
    "error_report": {
        "subscription": "sce_error_report_subscription",
        "display_name": "Error Report",
    },
    # The error subscription now returns TWO entries for one correlation ID.
    # This is the second one. Both are pulled from the same subscription; see
    # LAST_COUNT_DATE_HINTS below for how the two are told apart.
    "last_count_date_error_report": {
        "subscription": "sce_error_report_subscription",
        "display_name": "Last Count Date Error Report",
        "match_hints": ["lastcount", "countdate"],
    },
    "comparison_report": {
        "subscription": "fusion_comparison_report_subscription",
        "display_name": "Comparison Report",
    },
    "fusion_inventory": {
        "subscription": "fusion_inventory_update_subscription",
        "display_name": "Fusion Inventory Report",
    },
    "mhs_snapshot": {
        "subscription": "mhs_gis_snapshot_subscription",
        "display_name": "MHS Snapshot",
    },
    "sce_snapshot": {
        "subscription": "sce_gis_snapshot_subscription",
        "display_name": "SCE Snapshot",
    },
}

# Only these run when you click "Generate Reports". Add keys back here later
# (e.g. "sce_snapshot") to bring a report back online.
ENABLED_REPORTS = [
    "error_report", "last_count_date_error_report",
    "comparison_report", "fusion_inventory", "mhs_snapshot", "sce_snapshot",
]

# --- Telling the two error-subscription entries apart -----------------------
# For a report with "match_hints", the entry whose text contains ANY hint is
# that report. Hints are compared case-insensitively with punctuation ignored,
# so "lastcount" matches lastCountDate, last_count_date and last-count-date.
# Whatever entry is left over is the plain Error Report. If NO entry contains a
# hint, they're assigned by order of appearance (first = Error Report, second
# = Last Count Date Error Report) and the run log says so. Edit the hints in
# the "last_count_date_error_report" entry above once you know what really
# differs between the two messages.

# ---------------------------------------------------------------------------
# Error report - field names (confirmed from real data: Error_Report_059.json)
# ---------------------------------------------------------------------------
# Each error record looks like:
#   {"lineId": 42, "handlingUnit": "...", "article": "70627545",
#    "location": "680510", "errors": [{"errorDescription": "Missing Location"}]}
# "handlingUnit" is sometimes absent. "errors" is a list (usually one entry,
# but the code handles more) - each error entry becomes its own row.
ARTICLE_FIELD = "article"
LOCATION_FIELD = "location"
ERRORS_FIELD = "errors"

# ---------------------------------------------------------------------------
# Comparison report - structure (confirmed from real data: 268_inventory_
# comparison_20260911.json / .xlsx)
# ---------------------------------------------------------------------------
# {"header": {...}, "data": {"quantityMismatches": [...], "missingInMHS": [...],
#  "missingInSCX": [...], "summary": {...}}}
# quantityMismatches record: {articleNumber, mhsQuantity, scxQuantity, price, reasonCodes}
# missingInMHS / missingInSCX record: {articleNumber, quantity, price, reasonCodes}
# reasonCodes is a list - joined with ", " in the xlsx, blank when empty.
COMPARISON_QM_COLUMNS = ["articleNumber", "mhsQuantity", "scxQuantity", "Difference", "price", "reasonCodes"]
COMPARISON_MISSING_COLUMNS = ["articleNumber", "quantity", "price", "reasonCodes"]

# ---------------------------------------------------------------------------
# Fusion Inventory report - structure (confirmed from real data: 268_Inventory_
# Fusion_Articles_20260911.json / 268_Inventory_Article_Report_20260911.xlsx)
# ---------------------------------------------------------------------------
# {"header": {...}, "data": [ ...15653 real records... ]}
# Each record is a flat dict (handlingUnit, status, article, location, quantity,
# handlingUnitType, palletLength/Width/Height/Weight, supplier, productionWeek,
# lastMoveDate, lastCountDate, expireDate, nonSellableGroup,
# nonSellableDescription, lineId) - columns are taken directly from whatever
# keys the records have, in first-seen order, so no fixed column list is
# needed here (unlike Error/Comparison reports).

# ---------------------------------------------------------------------------
# MHS Snapshot report - structure (confirmed from real data: 268_MHS_Snapshot_
# report_20260911_183613.json / .xlsx, 17818 real events)
# ---------------------------------------------------------------------------
# {"header": {...}, "data": {"inventoryEvents": [...]}}
# Each event: {"businessUnit": {"type", "code"}, "item": {"number", "type", "uomCode"},
#              "stockOnHand": {"totalQuantity", "creationTimestamp",
#                               "storageAreas": [{"id","totalQuantity","totalQuantityBlocked"}],
#                               "reservations": [{"clientName","totalQuantity"}]}}
# storageAreas -> joined as "id:totalQuantity" entries, "; " separated, blank when empty.
# reservations -> joined as "clientName:totalQuantity" entries, "; " separated, blank when empty.
# totalQuantityBlocked is not used in the report (not in the reference file).
MHS_SNAPSHOT_COLUMNS = [
    "businessUnitType", "businessUnitCode", "itemNumber", "itemType", "uomCode",
    "totalQuantity", "creationTimestamp", "storageAreas", "reservations",
]

# ---------------------------------------------------------------------------
# SCE Snapshot report - structure (confirmed from real data: 268_SCE_Snapshot_
# 20260911.json / .xlsx, 9911 real events)
# ---------------------------------------------------------------------------
# {"header": {...}, "data": {"inventoryEvents": [...]}}
# Same event shape as MHS Snapshot, MINUS "reservations" (not present at all
# in this feed). Column names and the storageAreas format are both different
# from MHS Snapshot though - this reference file was clearly produced by a
# different script (dotted pandas.json_normalize()-style headers, and
# storageAreas written as the literal Python str() of the list of dicts,
# e.g. "[{'id': '01', 'totalQuantity': 132.0, 'totalQuantityBlocked': 0.0}]",
# rather than the clean "id:qty; id:qty" format MHS Snapshot uses). Both are
# replicated exactly as each reference file has them - not made consistent
# with each other.
SCE_SNAPSHOT_COLUMNS = [
    "businessUnit.type", "businessUnit.code", "item.number", "item.type", "item.uomCode",
    "stockOnHand.totalQuantity", "stockOnHand.creationTimestamp", "stockOnHand.storageAreas",
]

# ---------------------------------------------------------------------------
# Colors (hex, no '#')
# ---------------------------------------------------------------------------
# Note: Copilot's own header coloring turned out to be inconsistent between
# runs (blue on one file, green on another), so rather than chase a moving
# target, the app uses one fixed color per sheet. Comparison/Fusion colors
# below ARE confirmed exact matches from your real reference xlsx files though.
COLOR_SUMMARY_HEADER_FILL = "4472C4"     # Error report Summary header
COLOR_QM_HEADER_FILL = "5B9BD5"          # Comparison report: Quantity Mismatches header
COLOR_MISSING_HEADER_FILL = "92D050"     # Comparison report: Missing In SCX / Missing In MHS header
COLOR_COMPARISON_SUMMARY_HEADER_FILL = "C65911"  # Comparison report: Summary header
COLOR_FUSION_HEADER_FILL = "1F4E78"      # Fusion Inventory report header
COLOR_SCE_HEADER_FILL = "1F4E78"         # SCE Snapshot report header
COLOR_WHITE = "FFFFFF"

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
# Root folder under which "<storecode>_<date>" folders are created.
DEFAULT_OUTPUT_ROOT = os.path.expanduser(r"~/Documents/FMS/baseload_reports")

# ---------------------------------------------------------------------------
# Fetch polling
# ---------------------------------------------------------------------------
# Pulls + greps for the correlation ID every FETCH_POLL_INTERVAL_SECONDS,
# for up to FETCH_MAX_WAIT_SECONDS total, before giving up. Matches the
# polling behavior of the working reference script (30s / 1 hour).
FETCH_POLL_INTERVAL_SECONDS = 30
FETCH_MAX_WAIT_SECONDS = 3600
