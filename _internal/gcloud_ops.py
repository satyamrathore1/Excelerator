"""
gcloud_ops.py
-------------
Wraps every manual step you currently do by hand in Google Cloud SDK / Chrome
into plain Python functions. Rewritten to match your confirmed-working
reference script (error_report_extract.txt):

  1. authenticate()          -> gcloud auth activate-service-account
  2. configure_project()     -> gcloud config set project
  3. generate_uuid()         -> same as your generate_uuid.py
  4. trigger_baseload()      -> gcloud pubsub topics publish ...
  5. fetch_report_url()      -> poll `pull ... | findstr <correlationId>` every
                                30s for up to 1 hour, regex out the signed URL
  6. download_json()         -> download + unwrap (gzip and/or zip) into plain JSON bytes

All commands run via `shell=True` with plain command strings (exactly like
your reference script), which is what lets Windows resolve `gcloud` to
`gcloud.cmd` automatically - cmd.exe (which shell=True invokes) does the
PATHEXT resolution that a shell=False + argument-list subprocess call
doesn't. This also means no manual "find gcloud.cmd" workaround is needed
anymore.

One deliberate difference from the reference script: downloading uses a
plain `requests.get()` instead of Selenium + ChromeDriver. A signed GCS URL
is fully self-authorizing via its query parameters, so a direct HTTP GET
gets the same bytes a browser would - without the Chrome/driver dependency,
the "wait 10 seconds and hope the download finished" timing, or having to
guess which file just landed in the downloads folder.
"""

from __future__ import annotations

import gzip
import io
import json
import re
import subprocess
import time
import uuid as uuid_lib
import zipfile
from typing import Optional

import requests

import config
import settings


class GCloudError(RuntimeError):
    """Raised when a gcloud/subprocess call fails."""


def _run(command: str, timeout: int = 120) -> str:
    """Run a shell command (string, not list - matches the reference script),
    return stdout, raise GCloudError with stderr on failure."""
    result = subprocess.run(
        command,
        shell=True,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    if result.returncode != 0:
        raise GCloudError(
            f"Command failed ({result.returncode}): {command}\n"
            f"STDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
        )
    return result.stdout


# ---------------------------------------------------------------------------
# 1-4: auth / project / uuid / trigger
# ---------------------------------------------------------------------------

def authenticate(key_file_path: str) -> str:
    """gcloud auth activate-service-account --key-file="<path>" """
    gcloud = settings.gcloud_command_prefix()
    return _run(f'{gcloud} auth activate-service-account --key-file="{key_file_path}"')


def configure_project(project_id: str = config.PROJECT_ID) -> str:
    """gcloud config set project <project_id>"""
    gcloud = settings.gcloud_command_prefix()
    return _run(f"{gcloud} config set project {project_id}")


def generate_uuid() -> str:
    """Same as generate_uuid.py - kept as a function so the UI can call it directly."""
    return str(uuid_lib.uuid4())


def trigger_baseload(store_code: str, correlation_id: str, project_id: str = config.PROJECT_ID) -> Optional[str]:
    """
    gcloud pubsub topics publish hcl_cutover_trigger_topic
      --message="{\\"businessUnitCode\\":\\"<store_code>\\", \\"correlationId\\":\\"<uuid>\\"}"
      --project=<project_id> --format=json

    Returns the publish's messageId (pubsub's own message ID for the publish
    call, distinct from your correlationId), or None if it couldn't be parsed
    out of the response - the publish itself will still have succeeded (an
    exception would have been raised by _run() otherwise), just without a
    messageId to show on the dashboard.
    """
    message_json = f'{{\\"businessUnitCode\\":\\"{store_code}\\", \\"correlationId\\":\\"{correlation_id}\\"}}'
    gcloud = settings.gcloud_command_prefix()
    command = (
        f"{gcloud} pubsub topics publish {config.BASELOAD_TOPIC} "
        f'--message="{message_json}" '
        f"--project={project_id} "
        f"--format=json"
    )
    stdout = _run(command)
    try:
        parsed = json.loads(stdout)
        if isinstance(parsed, list) and parsed:
            message_ids = parsed[0].get("messageIds") or []
            if message_ids:
                return message_ids[0]
    except (json.JSONDecodeError, AttributeError, IndexError):
        pass
    return None


# ---------------------------------------------------------------------------
# 5: fetch report url (poll + findstr, exactly like the reference script)
# ---------------------------------------------------------------------------

_URL_PATTERN = re.compile(r'https?://[^\s\]"]+')


def _norm(text: str) -> str:
    """Lowercase and drop everything but letters/digits, so 'lastCountDate',
    'last_count_date' and 'last-count-date' all compare equal."""
    return re.sub(r"[^a-z0-9]", "", text.lower())


def try_fetch_report_entries(subscription: str, correlation_id: str,
                              project_id: str = config.PROJECT_ID) -> list:
    """
    ONE non-blocking attempt. Runs:
        gcloud pubsub subscriptions pull <subscription> --limit=100 --project=<project_id> | findstr <correlation_id>
    and returns EVERY matching message as {"url": signed_url, "line": raw_text},
    in the order they appeared. Empty list = nothing there yet.

    A subscription can now hold more than one message for the same correlation
    ID (the error subscription returns two), so this returns all of them
    instead of just the first.
    """
    gcloud = settings.gcloud_command_prefix()
    subscription_command = (
        f"{gcloud} pubsub subscriptions pull {subscription} "
        f"--limit=100 --project={project_id}"
    )
    command = f"{subscription_command} | findstr {correlation_id}"

    result = subprocess.run(command, shell=True, capture_output=True, text=True, timeout=120)
    output = result.stdout.strip()
    if not output:
        return []

    cleaned = output.replace("\\/", "/")
    entries, seen = [], set()
    for line in cleaned.splitlines():
        match = _URL_PATTERN.search(line)
        if match and match.group(0) not in seen:
            seen.add(match.group(0))
            entries.append({"url": match.group(0), "line": line})

    if not entries:
        # A single message that got wrapped across lines: fall back to
        # treating the whole output as one entry.
        match = _URL_PATTERN.search(cleaned)
        if match:
            entries.append({"url": match.group(0), "line": cleaned})
    return entries


def assign_entries(entries: list, variants: list, final: bool = False) -> dict:
    """
    Work out which pulled message belongs to which report.

    variants: [{"key": report_key, "hints": [..]}, ...] for every enabled
    report that reads this subscription, in declaration order.
    Returns {report_key: {"url", "line", "how"}} for the reports that could be
    matched right now. "how" is one of:
        "name"         the message text contained one of the report's hints
        "elimination"  it's the one left after the others were matched by name
        "order"        no hints matched, so assigned by order of appearance

    When nothing matches by name and fewer messages have arrived than there
    are reports, this WAITS (returns {}) rather than guess, because the first
    message to arrive isn't necessarily the first in the list. Pass
    final=True on the last attempt to use what's there instead.
    """
    hinted = [v for v in variants if v.get("hints")]
    plain = [v for v in variants if not v.get("hints")]
    result: dict = {}
    leftover: list = []

    for entry in entries:
        norm_line = _norm(entry["line"])
        for variant in hinted:
            if variant["key"] in result:
                continue
            if any(_norm(h) and _norm(h) in norm_line for h in variant["hints"]):
                result[variant["key"]] = {**entry, "how": "name"}
                break
        else:
            leftover.append(entry)

    matched_by_name = bool(result)
    if not matched_by_name and not final and len(leftover) < len(variants):
        return {}

    for variant in plain:
        if leftover:
            how = "elimination" if matched_by_name else "order"
            result[variant["key"]] = {**leftover.pop(0), "how": how}
    for variant in hinted:
        if variant["key"] not in result and leftover:
            result[variant["key"]] = {**leftover.pop(0), "how": "order"}
    return result


def try_fetch_report_url(subscription: str, correlation_id: str,
                          project_id: str = config.PROJECT_ID) -> Optional[str]:
    """First matching message's signed URL, or None. Kept for single-report
    callers; the UI uses try_fetch_report_entries() + assign_entries()."""
    entries = try_fetch_report_entries(subscription, correlation_id, project_id)
    return entries[0]["url"] if entries else None


def fetch_report_url(subscription: str, correlation_id: str, project_id: str = config.PROJECT_ID,
                      poll_interval_seconds: Optional[int] = None,
                      max_wait_seconds: Optional[int] = None) -> str:
    """
    Blocking variant: retries try_fetch_report_url() every
    `poll_interval_seconds` (default 30) for up to `max_wait_seconds`
    (default 3600 = 1 hour), same as the reference script.

    Kept for scripted/one-report use. The Streamlit UI does NOT use this -
    it drives try_fetch_report_url() itself so a slow report can't hold up
    the others.
    """
    poll_interval = poll_interval_seconds if poll_interval_seconds is not None else config.FETCH_POLL_INTERVAL_SECONDS
    max_wait = max_wait_seconds if max_wait_seconds is not None else config.FETCH_MAX_WAIT_SECONDS

    start_time = time.time()
    while time.time() - start_time < max_wait:
        url = try_fetch_report_url(subscription, correlation_id, project_id)
        if url:
            return url
        time.sleep(poll_interval)

    raise GCloudError(
        f"Correlation ID {correlation_id} not found in subscription '{subscription}' "
        f"within {max_wait} seconds."
    )


# ---------------------------------------------------------------------------
# 6: download + unwrap
# ---------------------------------------------------------------------------

def _unwrap_archive(raw: bytes) -> bytes:
    """
    Repeatedly unwraps gzip and/or zip layers until it finds the actual
    JSON content. Handles both cases you've seen: a raw .gz stream, and a
    .zip archive containing a single file.
    """
    data = raw
    for _ in range(4):  # a few layers is more than enough
        if data[:2] == b"\x1f\x8b":  # gzip magic number
            data = gzip.decompress(data)
        elif data[:4] == b"PK\x03\x04":  # zip magic number
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = [n for n in zf.namelist() if not n.endswith("/")]
                if not names:
                    raise GCloudError("Downloaded zip archive is empty.")
                data = zf.read(names[0])
        else:
            break
    return data


def download_json(url: str, save_path: str, auth_token: Optional[str] = None) -> None:
    """Downloads the file at `url`, unwraps gzip/zip if needed, writes plain JSON to `save_path`."""
    headers = {}
    if auth_token:
        headers["Authorization"] = f"Bearer {auth_token}"
    resp = requests.get(url, headers=headers, timeout=120)
    resp.raise_for_status()
    content = _unwrap_archive(resp.content)
    with open(save_path, "wb") as f:
        f.write(content)
