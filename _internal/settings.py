r"""
settings.py
-----------
Small persisted user settings, so each person only has to configure
Excelerator once.

Saving is deliberately defensive. Corporate Windows machines regularly make
the Roaming AppData folder unwritable or locked (folder redirection, DLP /
antivirus holding a file open, a read-only file left behind by an elevated
run, OneDrive sync conflicts). A refused write must never take the app down,
so:

  * writes are atomic (temp file + replace), so a half-written file can't
    corrupt existing settings;
  * several locations are tried in order until one accepts the write;
  * every save is also kept in memory for the life of the process, so the
    current session works with the new values even if NOTHING on disk could
    be written - save() then just reports that it couldn't persist them.

Locations tried, in order (Windows):
    %APPDATA%\Excelerator\settings.json
    %LOCALAPPDATA%\Excelerator\settings.json
    ~\.excelerator\settings.json
    <temp>\Excelerator\settings.json
When reading, the most recently modified settings file among these wins.
"""

from __future__ import annotations

import json
import os
import shutil
import stat
import tempfile

DEFAULTS = {
    "gcloud_path": "",       # empty = resolved via PATH at run time
    "service_account_key": "",
    "output_root": os.path.join(os.path.expanduser("~"), "Documents", "Excelerator Reports"),
}

_FILE_NAME = "settings.json"

# Values saved during this process's lifetime. Always updated by save(), even
# when the disk write fails, and layered over whatever's on disk by load().
_memory: dict = {}

# Tests can point this at a fixed list of directories.
_CANDIDATE_DIRS_OVERRIDE: list[str] | None = None


def _candidate_dirs() -> list[str]:
    if _CANDIDATE_DIRS_OVERRIDE is not None:
        return list(_CANDIDATE_DIRS_OVERRIDE)

    dirs = []
    appdata = os.environ.get("APPDATA")
    localappdata = os.environ.get("LOCALAPPDATA")
    if appdata:
        dirs.append(os.path.join(appdata, "Excelerator"))
    if localappdata:
        dirs.append(os.path.join(localappdata, "Excelerator"))
    dirs.append(os.path.join(os.path.expanduser("~"), ".excelerator"))
    dirs.append(os.path.join(tempfile.gettempdir(), "Excelerator"))

    seen, unique = set(), []
    for d in dirs:
        if d not in seen:
            seen.add(d)
            unique.append(d)
    return unique


def _candidate_paths() -> list[str]:
    return [os.path.join(d, _FILE_NAME) for d in _candidate_dirs()]


def _existing_settings_files() -> list[str]:
    return [p for p in _candidate_paths() if os.path.isfile(p)]


def _newest_settings_file() -> str | None:
    files = _existing_settings_files()
    if not files:
        return None
    try:
        return max(files, key=os.path.getmtime)
    except OSError:
        return files[0]


def _write_json_atomic(path: str, data: dict) -> None:
    """Write via a temp file in the same folder, then replace. Raises OSError
    (including PermissionError) if the folder or file refuses the write."""
    folder = os.path.dirname(path)
    os.makedirs(folder, exist_ok=True)

    fd, tmp_path = tempfile.mkstemp(prefix="settings-", suffix=".tmp", dir=folder)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        try:
            os.replace(tmp_path, path)
        except PermissionError:
            # Most common cause on Windows: the existing file is marked
            # read-only. Clear that bit and try once more.
            if os.path.exists(path):
                os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
            os.replace(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def is_first_run() -> bool:
    """True only when there's neither a settings file anywhere nor anything
    saved earlier in this process."""
    return not _memory and _newest_settings_file() is None


def load() -> dict:
    merged = dict(DEFAULTS)
    path = _newest_settings_file()
    if path:
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            merged.update({k: v for k, v in data.items() if k in DEFAULTS})
        except (json.JSONDecodeError, OSError):
            pass
    merged.update(_memory)
    return merged


def save(new_values: dict) -> tuple[bool, str]:
    """
    Persist settings. NEVER raises.

    Returns (persisted, message):
      (True,  "<path>")   - written to disk at that path.
      (False, "<reason>") - couldn't write anywhere; values are still active
                            for this session (kept in memory) but won't be
                            remembered next launch.
    """
    clean = {k: v for k, v in new_values.items() if k in DEFAULTS}
    _memory.update(clean)

    to_write = load()  # DEFAULTS + newest file + memory (which now includes clean)

    errors = []
    for path in _candidate_paths():
        try:
            _write_json_atomic(path, to_write)
            return True, path
        except OSError as e:
            errors.append(f"{path}: {e.strerror or e}")

    return False, ("Could not write settings to any location. " + " | ".join(errors))


def reset() -> None:
    """Used by the sidebar's "redo setup" action. Best-effort: a file that
    can't be removed is simply left in place."""
    _memory.clear()
    for path in _existing_settings_files():
        try:
            os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
            os.remove(path)
        except OSError:
            pass


def detect_gcloud() -> str:
    """Best-effort auto-detect for pre-filling the setup screen. Empty string
    if nothing is found - the field is still editable either way."""
    return shutil.which("gcloud") or ""


def gcloud_command_prefix() -> str:
    """
    What to put in front of every gcloud subcommand string.

    If the user configured an explicit path (from setup or the sidebar), use
    that exact path, quoted for spaces. Otherwise fall back to the bare word
    "gcloud" and let the shell (cmd.exe, via shell=True) resolve it from PATH.
    """
    path = load().get("gcloud_path", "").strip()
    return f'"{path}"' if path else "gcloud"
