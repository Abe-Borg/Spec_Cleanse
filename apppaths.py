#!/usr/bin/env python3
"""
Filesystem locations SpecCleanse uses at runtime.

Running from source, ``patterns.yaml`` sits beside the modules and there is
nothing to work out. Running from a PyInstaller bundle there are two copies to
tell apart: the read-only one unpacked into a temporary directory with the
code, and the one the user is meant to edit. `README.md` documents editing
``patterns.yaml`` as the way to add detection patterns without touching code,
so a frozen build has to keep that file somewhere that survives the run.

Kept apart from ``gui.py`` so the rules can be tested without tkinter, and
apart from ``docx_xml.py``, which holds WordprocessingML plumbing rather than
install layout.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

APP_NAME = "SpecCleanse"
CONFIG_NAME = "patterns.yaml"


def is_frozen() -> bool:
    """True when running from a PyInstaller bundle rather than the source tree."""
    return bool(getattr(sys, "frozen", False)) and hasattr(sys, "_MEIPASS")


def bundle_dir() -> Path:
    """Directory holding the running code.

    Frozen, that is PyInstaller's extraction directory, which is temporary:
    anything written there is gone when the process exits.
    """
    if is_frozen():
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def executable_dir() -> Path:
    """Directory holding the running executable, or the source tree."""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def user_config_dir() -> Path:
    """Per-user configuration directory, created on demand by the caller.

    ``%APPDATA%\\SpecCleanse`` on Windows, falling back to the XDG location
    elsewhere so the module stays importable and testable off Windows.
    """
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / APP_NAME

    xdg = os.environ.get("XDG_CONFIG_HOME")
    if xdg:
        return Path(xdg) / APP_NAME

    return Path.home() / ".config" / APP_NAME


def resolve_config_path() -> Path:
    """Return the ``patterns.yaml`` the app should load.

    From source, the file beside the modules — the developer layout, unchanged.

    Frozen, a copy beside the executable wins when one is present, so a portable
    unzip-and-run install can be customised by dropping the file next to the
    ``.exe``. Otherwise the per-user copy under `user_config_dir` is used, seeded
    from the bundled default on first run so there is always something to edit.

    Seeding is best-effort. A locked-down or read-only profile falls back to the
    bundled copy: the app still runs against the shipped patterns, it just cannot
    be customised. Never raises — a configuration that cannot be found at all is
    reported by ``load_config`` with a readable message instead.
    """
    if not is_frozen():
        return bundle_dir() / CONFIG_NAME

    beside_exe = executable_dir() / CONFIG_NAME
    if beside_exe.is_file():
        return beside_exe

    bundled = bundle_dir() / CONFIG_NAME
    user_copy = user_config_dir() / CONFIG_NAME

    if user_copy.is_file():
        return user_copy

    try:
        user_copy.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(bundled, user_copy)
        return user_copy
    except OSError:
        return bundled
