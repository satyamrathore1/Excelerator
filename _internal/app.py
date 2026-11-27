"""
app.py
------
Excelerator - Streamlit UI for the baseload + report automation tool.

Generates: Error Report, Comparison Report, Fusion Inventory Report,
MHS Snapshot, SCE Snapshot - see config.ENABLED_REPORTS. Every enabled report
gets a Preview button on the dashboard.

Run with:
    streamlit run app.py
or, once packaged, just double-click Excelerator.exe.
"""

import json as _json
import os
import platform
import time
from datetime import datetime

import streamlit as st

import config
import gcloud_ops
import report_builder
import settings as app_settings

st.set_page_config(
    page_title="Excelerator",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

STORE_CODE_KEY = "store_code_input"
CORRELATION_ID_KEY = "correlation_id_input"
_PENDING_CORRELATION_ID_KEY = "_pending_correlation_id"

# Dispatch table: report key -> (xlsx builder function, output file prefix)
REPORT_BUILDERS = {
    "error_report": (report_builder.build_error_report_xlsx, "Error_Report"),
    "last_count_date_error_report": (report_builder.build_last_count_date_error_report_xlsx,
                                     "last-count-date-error-report"),
    "comparison_report": (report_builder.build_comparison_report_xlsx, "Comparison_Report"),
    "fusion_inventory": (report_builder.build_fusion_report_xlsx, "Fusion_Inventory_Report"),
    "mhs_snapshot": (report_builder.build_mhs_snapshot_xlsx, "MHS_Snapshot_Report"),
    "sce_snapshot": (report_builder.build_sce_snapshot_xlsx, "SCE_Snapshot_Report"),
}

# ---------------------------------------------------------------------------
# styling
# ---------------------------------------------------------------------------
st.markdown("""
<style>
  /* tighten the default top padding so the header sits higher */
  .block-container { padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1200px; }

  .app-header {
    display: flex; align-items: baseline; gap: 12px;
    padding-bottom: 6px; margin-bottom: 4px;
    border-bottom: 3px solid #1F4E78;
  }
  .app-header h1 {
    margin: 0; font-size: 2.1rem; font-weight: 700;
    color: #1F4E78; letter-spacing: -0.5px;
  }
  .app-header .byline { font-size: 0.85rem; color: #8a8a8a; }
  .app-tagline { color: #6b6b6b; font-size: 0.9rem; margin-bottom: 1.4rem; }

  .step-label {
    font-size: 0.78rem; font-weight: 700; letter-spacing: 1.2px;
    text-transform: uppercase; color: #1F4E78; margin-bottom: 2px;
  }

  /* report row cards */
  .report-card {
    border: 1px solid #e4e7eb; border-left: 4px solid #cbd2d9;
    border-radius: 6px; padding: 12px 16px; margin-bottom: 10px;
    background: #fbfcfd;
  }
  .report-card.done    { border-left-color: #2e7d32; }
  .report-card.waiting { border-left-color: #f0a202; }
  .report-card.failed  { border-left-color: #c62828; }
  .report-name { font-weight: 600; font-size: 1.02rem; color: #23282d; }

  .pill {
    display: inline-block; padding: 2px 10px; border-radius: 11px;
    font-size: 0.72rem; font-weight: 700; letter-spacing: 0.4px;
    text-transform: uppercase; margin-left: 8px; vertical-align: middle;
  }
  .pill.done    { background: #e6f4ea; color: #1e6b2a; }
  .pill.waiting { background: #fdf3d8; color: #8a6100; }
  .pill.failed  { background: #fbe4e4; color: #a11a1a; }

  .metric-inline { display: flex; gap: 26px; }
  .metric-inline .mi-label {
    font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.6px; color: #78828c;
  }
  .metric-inline .mi-value { font-size: 1.45rem; font-weight: 700; color: #1F4E78; line-height: 1.15; }

  .err-note { color: #a11a1a; font-size: 0.82rem; margin-top: 4px; }

  .app-footer {
    margin-top: 2.5rem; padding-top: 10px; border-top: 1px solid #e4e7eb;
    color: #9aa0a6; font-size: 0.78rem; text-align: center;
  }

  div[data-testid="stSidebar"] .stTextInput label { font-weight: 600; font-size: 0.82rem; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# session state
# ---------------------------------------------------------------------------
defaults = {
    STORE_CODE_KEY: "",
    CORRELATION_ID_KEY: "",
    "message_id": None,
    "output_dir": None,
    "report_results": {},  # report_key -> {"json":..., "xlsx":..., ...metrics}
    "log": [],
    # Rerun-driven polling state. Streamlit only paints the page once a script
    # run finishes, so a blocking wait loop would keep the Step 3 summary
    # hidden until every report was done. Instead each script run does ONE
    # polling pass and then reruns itself, letting the whole page (summary
    # included) repaint between passes.
    "run_active": False,
    "run_pending": [],      # report keys still waiting to appear
    "run_deadline": None,   # epoch seconds; None when no run in progress
    "run_store_code": "",   # frozen at click time so edits mid-run don't retarget files
    "run_correlation_id": "",
    "run_project_id": "",
    "run_date_str": "",
    "run_time_str": "",
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# A widget's session_state key can't be written to AFTER that widget has been
# instantiated this run (Streamlit raises StreamlitAPIException). So instead
# of setting st.session_state[CORRELATION_ID_KEY] directly inside the Trigger
# button's handler (which runs after the text_input below has already been
# created), we stage the new value here - BEFORE the widget exists - and
# apply it. The Trigger handler further down just sets the pending key and
# calls st.rerun(), which brings execution back to this exact point.
if st.session_state.get(_PENDING_CORRELATION_ID_KEY) is not None:
    st.session_state[CORRELATION_ID_KEY] = st.session_state.pop(_PENDING_CORRELATION_ID_KEY)


def log(msg: str):
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.log.append(f"[{timestamp}] {msg}")


def open_file(path: str):
    """Opens a file with the OS default viewer (preview button)."""
    if not path or not os.path.exists(path):
        st.warning("File not found yet.")
        return
    if platform.system() == "Windows":
        os.startfile(path)  # noqa
    elif platform.system() == "Darwin":
        os.system(f'open "{path}"')
    else:
        os.system(f'xdg-open "{path}"')


def save_settings(values: dict) -> None:
    """Save settings without ever crashing the page.

    A refused write (locked-down AppData, read-only file, antivirus holding
    the file) used to raise a PermissionError straight into the UI. Now the
    values still apply for this session, and a warning explains that they
    won't be remembered next launch."""
    persisted, info = app_settings.save(values)
    if persisted:
        st.session_state.pop("settings_warning", None)
    else:
        st.session_state["settings_warning"] = info


# ---------------------------------------------------------------------------
# First-run setup
# ---------------------------------------------------------------------------
# Shown once, the very first time anyone opens Excelerator on a machine (no
# settings file yet). Also reachable later via "Redo setup" in the sidebar.
# Everything entered here is saved to a per-user settings file, so later
# launches skip straight to the main app.
show_setup = (
    (app_settings.is_first_run() or st.session_state.get("force_setup"))
    and not st.session_state.get("setup_just_completed")
)

if show_setup:
    st.markdown(
        "<div class='app-header'><h1>⚡ Welcome to Excelerator</h1></div>"
        "<div class='app-tagline'>First-time setup - this only takes a minute, "
        "and you can change any of it later from the sidebar.</div>",
        unsafe_allow_html=True,
    )

    current = app_settings.load()
    detected_gcloud = app_settings.detect_gcloud()

    st.markdown("#### 1. Google Cloud SDK")
    if detected_gcloud:
        st.success(f"Found `gcloud` on this machine automatically: `{detected_gcloud}`", icon="✅")
    else:
        st.warning("Couldn't auto-detect `gcloud`. Open a terminal, run `where gcloud`, "
                   "and paste the path it gives you below.", icon="⚠️")
    gcloud_path_input = st.text_input(
        "Path to gcloud (leave blank to rely on PATH)",
        value=current["gcloud_path"] or detected_gcloud,
        placeholder=r"C:\Users\you\AppData\Local\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd",
    )

    st.markdown("#### 2. Service account key")
    key_path_input = st.text_input(
        "Path to your service account key file (.json)",
        value=current["service_account_key"],
        placeholder=r"C:\Users\you\Documents\ServiceAccount\ingka-fms-prod.json",
    )
    if key_path_input and not os.path.exists(key_path_input):
        st.caption("⚠️ Nothing found at that path yet - you can still save and fix it later.")

    st.markdown("#### 3. Where should reports be saved?")
    output_root_input = st.text_input(
        "Reports folder",
        value=current["output_root"],
        help="Each run creates its own timestamped subfolder inside this one.",
    )

    st.write("")
    if st.button("✅ Save & Continue", type="primary"):
        save_settings({
            "gcloud_path": gcloud_path_input.strip(),
            "service_account_key": key_path_input.strip(),
            "output_root": output_root_input.strip() or app_settings.DEFAULTS["output_root"],
        })
        st.session_state.pop("force_setup", None)
        st.session_state.setup_just_completed = True
        st.rerun()

    st.stop()


# ---------------------------------------------------------------------------
# sidebar - one-off settings
# ---------------------------------------------------------------------------
user_settings = app_settings.load()

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    st.caption("Changes here are saved automatically for next time.")

    key_file = st.text_input("Service account key file", value=user_settings["service_account_key"])
    if key_file and not os.path.exists(key_file):
        st.warning("File not found at this path.", icon="⚠️")

    gcloud_path = st.text_input("Path to gcloud (blank = use PATH)", value=user_settings["gcloud_path"])

    project_id = st.text_input("GCP project", value=config.PROJECT_ID)
    output_root = st.text_input("Reports root folder", value=user_settings["output_root"])
    st.caption("Each run creates its own timestamped folder: "
               "`<root>/<storecode>_<yyyymmdd_hhmmss>/`")

    if (key_file, gcloud_path, output_root) != (
        user_settings["service_account_key"], user_settings["gcloud_path"], user_settings["output_root"]
    ):
        save_settings({
            "service_account_key": key_file,
            "gcloud_path": gcloud_path,
            "output_root": output_root,
        })

    if st.session_state.get("settings_warning"):
        st.warning(
            "Your settings are active for this session, but Windows wouldn't let "
            "Excelerator save them, so you'll be asked again next time. "
            "Try closing other copies of Excelerator, or ask IT whether your "
            "AppData folder is locked down.",
            icon="⚠️",
        )
        with st.expander("Technical detail"):
            st.code(st.session_state["settings_warning"], language=None)

    st.divider()
    st.markdown("**Reports in this run**")
    for key in config.ENABLED_REPORTS:
        st.markdown(f"<span style='font-size:0.82rem;'>• {config.REPORTS[key]['display_name']}</span>",
                    unsafe_allow_html=True)

    st.divider()
    if st.session_state.output_dir:
        st.markdown("**Last output folder**")
        st.code(st.session_state.output_dir, language=None)
        if st.button("📂 Open folder", use_container_width=True):
            open_file(st.session_state.output_dir)

    st.divider()
    if st.button("🔁 Redo first-time setup", use_container_width=True):
        app_settings.reset()
        st.session_state["force_setup"] = True  # works even if a stuck file couldn't be deleted
        st.session_state.pop("setup_just_completed", None)
        st.rerun()

# ---------------------------------------------------------------------------
# header
# ---------------------------------------------------------------------------
st.markdown(
    "<div class='app-header'>"
    "<h1>⚡ Excelerator</h1>"
    "<span class='byline'>by Satyam</span>"
    "</div>"
    "<div class='app-tagline'>Trigger a baseload, then pull every cutover report "
    "straight into Excel and PDF - no manual gcloud, no copy-pasting URLs.</div>",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Step 1: trigger baseload (or supply an existing correlation ID directly)
# ---------------------------------------------------------------------------
st.markdown("<div class='step-label'>Step 1</div>", unsafe_allow_html=True)
st.subheader("Trigger Baseload")

col1, col2 = st.columns([1, 2])
with col1:
    store_code = st.text_input("Store code", key=STORE_CODE_KEY, placeholder="e.g. 268")
with col2:
    correlation_id = st.text_input(
        "Correlation ID",
        key=CORRELATION_ID_KEY,
        placeholder="Auto-filled on trigger, or paste an existing one",
        help="Click Trigger Baseload to generate a new one, or paste an existing "
             "correlation ID and skip straight to Generate Reports.",
    )

trigger_clicked = st.button("🚀 Trigger Baseload", type="primary",
                             disabled=not store_code or st.session_state.run_active)

if trigger_clicked:
    with st.spinner("Authenticating, configuring project, and triggering baseload..."):
        try:
            gcloud_ops.authenticate(key_file)
            gcloud_ops.configure_project(project_id)
            new_correlation_id = gcloud_ops.generate_uuid()
            message_id = gcloud_ops.trigger_baseload(store_code, new_correlation_id, project_id)

            st.session_state[_PENDING_CORRELATION_ID_KEY] = new_correlation_id
            st.session_state.message_id = message_id
            st.session_state.report_results = {}
            log(f"Baseload triggered for store {store_code} "
                f"(correlationId {new_correlation_id}, messageId {message_id})")
            st.rerun()  # re-render so the Correlation ID field shows the new value
        except gcloud_ops.GCloudError as e:
            st.error(f"Baseload trigger failed:\n\n{e}")

if correlation_id or st.session_state.message_id:
    id_col1, id_col2 = st.columns(2)
    with id_col1:
        if correlation_id:
            st.caption("Correlation ID — click the icon to copy")
            st.code(correlation_id, language=None)
    with id_col2:
        if st.session_state.message_id:
            st.caption("Message ID (from the last trigger)")
            st.code(st.session_state.message_id, language=None)

st.divider()

# ---------------------------------------------------------------------------
# Step 2: generate reports
# ---------------------------------------------------------------------------
st.markdown("<div class='step-label'>Step 2</div>", unsafe_allow_html=True)
st.subheader("Generate Reports")
st.caption("Runs once the migration has finished. Uses the Store code and Correlation ID "
           "above, so it works with a freshly triggered ID or one you pasted in yourself.")

btn_col1, btn_col2 = st.columns([1, 4])
with btn_col1:
    generate_clicked = st.button(
        "📊 Generate Reports",
        type="primary",
        disabled=not (store_code and correlation_id) or st.session_state.run_active,
        use_container_width=True,
    )
with btn_col2:
    if st.session_state.run_active:
        if st.button("✋ Stop waiting", key="cancel_run"):
            for report_key in st.session_state.run_pending:
                meta = config.REPORTS[report_key]
                st.session_state.report_results[report_key] = {
                    "error": "Cancelled while still waiting."
                }
                log(f"{meta['display_name']}: CANCELLED")
            st.session_state.run_active = False
            st.session_state.run_pending = []
            st.rerun()


def process_report(report_key: str, url: str, output_dir: str, store_code: str,
                    date_str: str, time_str: str) -> dict:
    """Downloads the JSON at `url` and builds the xlsx + pdf for one report.
    Files are named storeCode_reportName_Date_Time, e.g.
    268_Error_Report_20260917_143022.xlsx - all three files for a report,
    and every report in the same run, share the same date/time stamp."""
    builder_fn, file_prefix = REPORT_BUILDERS[report_key]
    base_name = f"{store_code}_{file_prefix}_{date_str}_{time_str}"

    json_path = os.path.join(output_dir, f"{base_name}.json")
    gcloud_ops.download_json(url, json_path)

    xlsx_path = os.path.join(output_dir, f"{base_name}.xlsx")
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
    with open(json_path, "r", encoding="utf-8") as f:
        json_data = _json.load(f)
    metrics = builder_fn(json_data, xlsx_path, pdf_path)

    return {"json": json_path, "xlsx": xlsx_path, "pdf": pdf_path, "url": url, **metrics}


# --- Starting a run: set up state, then let the polling block below take over ---
if generate_clicked:
    now = datetime.now()
    date_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H%M%S")
    folder_name = f"{store_code}_{date_str}_{time_str}"
    output_dir = os.path.join(output_root, folder_name)
    os.makedirs(output_dir, exist_ok=True)

    st.session_state.output_dir = output_dir
    st.session_state.report_results = {}
    st.session_state.run_pending = list(config.ENABLED_REPORTS)
    st.session_state.run_deadline = time.time() + config.FETCH_MAX_WAIT_SECONDS
    st.session_state.run_active = True
    # Freeze the inputs for the whole run, so editing the fields mid-run
    # doesn't retarget files or subscriptions halfway through.
    st.session_state.run_store_code = store_code
    st.session_state.run_correlation_id = correlation_id
    st.session_state.run_project_id = project_id
    st.session_state.run_date_str = date_str
    st.session_state.run_time_str = time_str
    log(f"Run started for store {store_code} -> {output_dir}")
    st.rerun()

# --- One polling pass per script run ---
# Each pass tries EVERY still-pending report once. Anything ready is
# downloaded and built immediately; anything not ready stays pending for the
# next pass. The script then finishes normally (so Step 3 below repaints with
# whatever is done so far) and schedules another rerun at the very bottom.
if st.session_state.run_active:
    run_store_code = st.session_state.run_store_code
    run_correlation_id = st.session_state.run_correlation_id
    run_project_id = st.session_state.run_project_id
    run_date_str = st.session_state.run_date_str
    run_time_str = st.session_state.run_time_str
    output_dir = st.session_state.output_dir

    status_box = st.empty()
    still_pending = []

    # The final attempt (within one poll interval of the deadline) is allowed
    # to settle for incomplete information rather than time out empty-handed.
    final_attempt = time.time() >= st.session_state.run_deadline - config.FETCH_POLL_INTERVAL_SECONDS

    # Group the still-waiting reports by subscription, so a subscription that
    # carries several reports (the error one now carries two) is pulled ONCE
    # per pass and its messages shared out, rather than pulled once per report.
    pending_by_subscription: dict = {}
    for report_key in st.session_state.run_pending:
        pending_by_subscription.setdefault(
            config.REPORTS[report_key]["subscription"], []
        ).append(report_key)

    for subscription, pending_keys in pending_by_subscription.items():
        names = " + ".join(config.REPORTS[k]["display_name"] for k in pending_keys)
        status_box.info(f"Checking {names}...")

        try:
            entries = gcloud_ops.try_fetch_report_entries(subscription, run_correlation_id, run_project_id)
        except Exception as e:
            for report_key in pending_keys:
                st.session_state.report_results[report_key] = {"error": str(e)}
                log(f"{config.REPORTS[report_key]['display_name']}: FAILED - {e}")
            continue

        # Every enabled report on this subscription takes part in the split
        # (not just the pending ones), so the assignment stays consistent
        # between passes even after one of them has already been built.
        variants = [
            {"key": k, "hints": config.REPORTS[k].get("match_hints", [])}
            for k in config.ENABLED_REPORTS
            if config.REPORTS[k]["subscription"] == subscription
        ]
        assigned = gcloud_ops.assign_entries(entries, variants, final=final_attempt)
        used_urls = {r.get("url") for r in st.session_state.report_results.values() if r.get("url")}

        for report_key in pending_keys:
            meta = config.REPORTS[report_key]
            entry = assigned.get(report_key)
            if entry is None or entry["url"] in used_urls:
                still_pending.append(report_key)
                continue

            try:
                status_box.info(f"Building {meta['display_name']}...")
                st.session_state.report_results[report_key] = process_report(
                    report_key, entry["url"], output_dir, run_store_code, run_date_str, run_time_str
                )
                used_urls.add(entry["url"])
                how = {"name": "matched by name", "elimination": "the other entry",
                       "order": "assigned by order of arrival"}.get(entry["how"], entry["how"])
                log(f"{meta['display_name']}: ready ({how})" if len(variants) > 1
                    else f"{meta['display_name']}: ready")
            except Exception as e:
                # A failure here is this report's problem alone - record it and
                # let the remaining reports carry on.
                st.session_state.report_results[report_key] = {"error": str(e)}
                log(f"{meta['display_name']}: FAILED - {e}")

    st.session_state.run_pending = still_pending
    status_box.empty()

    # Anything still pending past the deadline is a timeout; that ends the run.
    if still_pending and time.time() >= st.session_state.run_deadline:
        for report_key in still_pending:
            meta = config.REPORTS[report_key]
            siblings_done = [
                k for k in config.ENABLED_REPORTS
                if k != report_key
                and config.REPORTS[k]["subscription"] == meta["subscription"]
                and "error" not in st.session_state.report_results.get(k, {"error": 1})
            ]
            if siblings_done:
                reason = (f"Only part of the expected output arrived for correlation ID "
                          f"{run_correlation_id} in subscription '{meta['subscription']}' - "
                          f"this report's entry never showed up within "
                          f"{config.FETCH_MAX_WAIT_SECONDS} seconds.")
            else:
                reason = (f"Correlation ID {run_correlation_id} not found in subscription "
                          f"'{meta['subscription']}' within {config.FETCH_MAX_WAIT_SECONDS} seconds.")
            st.session_state.report_results[report_key] = {"error": reason}
            log(f"{meta['display_name']}: TIMED OUT")
        st.session_state.run_pending = []
        st.session_state.run_active = False
    elif not still_pending:
        st.session_state.run_active = False
        log("Run complete.")

# --- Run status banner + progress bar (reflects the pass that just happened) ---
results_so_far = st.session_state.report_results
if results_so_far or st.session_state.run_active:
    total = len(config.ENABLED_REPORTS)
    succeeded = [k for k, r in results_so_far.items() if "error" not in r]
    failed = [k for k, r in results_so_far.items() if "error" in r]
    finished = len(succeeded) + len(failed)

    st.progress(finished / total if total else 0.0,
                text=f"{finished} of {total} reports finished")

    if st.session_state.run_active:
        waiting_names = ", ".join(
            config.REPORTS[k]["display_name"] for k in st.session_state.run_pending
        )
        remaining = int(max(0, st.session_state.run_deadline - time.time()) // 60)
        st.info(f"Still waiting on: **{waiting_names}** — rechecking every "
                f"{config.FETCH_POLL_INTERVAL_SECONDS}s (about {remaining} min left before timeout). "
                f"Finished reports are ready below.", icon="⏳")
    elif succeeded:
        st.success(f"{len(succeeded)} of {total} reports saved to `{st.session_state.output_dir}`",
                   icon="✅")

st.divider()

# ---------------------------------------------------------------------------
# Step 3: dashboard
# ---------------------------------------------------------------------------
st.markdown("<div class='step-label'>Step 3</div>", unsafe_allow_html=True)
st.subheader("Report Summary")

if not st.session_state.report_results and not st.session_state.run_active:
    st.info("No reports generated yet. Trigger a baseload (or paste a correlation ID), "
            "then hit Generate Reports.", icon="💡")
else:
    for report_key in config.ENABLED_REPORTS:
        result = st.session_state.report_results.get(report_key)
        meta = config.REPORTS[report_key]
        name = meta["display_name"]

        # --- still waiting ---
        if not result:
            if report_key in st.session_state.run_pending:
                st.markdown(
                    f"<div class='report-card waiting'>"
                    f"<span class='report-name'>{name}</span>"
                    f"<span class='pill waiting'>waiting</span></div>",
                    unsafe_allow_html=True,
                )
            continue

        # --- failed / timed out ---
        if "error" in result:
            st.markdown(
                f"<div class='report-card failed'>"
                f"<span class='report-name'>{name}</span>"
                f"<span class='pill failed'>unavailable</span>"
                f"<div class='err-note'>{result['error']}</div></div>",
                unsafe_allow_html=True,
            )
            continue

        # --- done ---
        st.markdown(
            f"<div class='report-card done'>"
            f"<span class='report-name'>{name}</span>"
            f"<span class='pill done'>ready</span></div>",
            unsafe_allow_html=True,
        )

        left, right = st.columns([1, 3])
        with left:
            if st.button("👁 Preview PDF", key=f"preview_{report_key}", use_container_width=True):
                open_file(result["pdf"])
        with right:
            if report_key in ("error_report", "last_count_date_error_report"):
                st.markdown(
                    "<div class='metric-inline'>"
                    f"<div><div class='mi-label'>Total Errors</div>"
                    f"<div class='mi-value'>{result['total_errors']:,}</div></div>"
                    f"<div><div class='mi-label'>Unique Articles</div>"
                    f"<div class='mi-value'>{result['unique_articles']:,}</div></div>"
                    f"<div><div class='mi-label'>Unique Locations</div>"
                    f"<div class='mi-value'>{result['unique_locations']:,}</div></div>"
                    "</div>",
                    unsafe_allow_html=True,
                )
            elif report_key == "comparison_report":
                qty_pct = result.get("percentage_quantity_match")
                val_pct = result.get("percentage_value_match")
                qty_txt = f"{qty_pct:.2f}%" if qty_pct is not None else "N/A"
                val_txt = f"{val_pct:.2f}%" if val_pct is not None else "N/A"
                st.markdown(
                    "<div class='metric-inline'>"
                    f"<div><div class='mi-label'>Quantity Match</div>"
                    f"<div class='mi-value'>{qty_txt}</div></div>"
                    f"<div><div class='mi-label'>Value Match</div>"
                    f"<div class='mi-value'>{val_txt}</div></div>"
                    "</div>",
                    unsafe_allow_html=True,
                )
            elif "record_count" in result:
                st.markdown(
                    "<div class='metric-inline'>"
                    f"<div><div class='mi-label'>Records</div>"
                    f"<div class='mi-value'>{result['record_count']:,}</div></div>"
                    "</div>",
                    unsafe_allow_html=True,
                )

        with st.expander("Files on disk"):
            for label in ("json", "xlsx", "pdf"):
                if result.get(label):
                    st.code(result[label], language=None)

st.divider()
with st.expander("🧾 Run log"):
    if st.session_state.log:
        st.text("\n".join(st.session_state.log))
    else:
        st.caption("Nothing logged yet.")

st.markdown("<div class='app-footer'>Excelerator · by Satyam</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Schedule the next polling pass
# ---------------------------------------------------------------------------
# This sits at the very BOTTOM on purpose: everything above has already been
# rendered by now, so the finished reports are on screen (and their Preview
# buttons usable) while we wait here for the next pass. Sleeping in short
# slices rather than one long block keeps the Stop button and Preview clicks
# reasonably responsive - a click during the wait reruns the script, which
# just performs the next pass early.
if st.session_state.run_active:
    slept = 0.0
    while slept < config.FETCH_POLL_INTERVAL_SECONDS:
        time.sleep(0.5)
        slept += 0.5
    st.rerun()
