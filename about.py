#!/usr/bin/env python3
"""
What the About dialog says: copyright, license, and where to get in touch.

The Windows installer shows the license on its wizard page, but a silent install
skips that page and the portable executable has no installer at all. PolyForm's
Notices clause asks that whoever gets a copy also gets the terms (or their URL)
and every ``Required Notice:`` line, so the application itself states both, and
can open the full license and the third-party notices bundled inside it.

Kept apart from ``gui.py`` so the text can be tested without tkinter.
``REQUIRED_NOTICE`` and ``LICENSE_URL`` repeat what ``LICENSE.md`` says, and a
test holds them to it.
"""

from __future__ import annotations

APP_NAME = "SpecCleanse"

PROJECT_URL = "https://github.com/Abe-Borg/Spec_Cleanse"
SUPPORT_URL = f"{PROJECT_URL}/issues"
RELEASES_URL = f"{PROJECT_URL}/releases"

LICENSE_NAME = "PolyForm Noncommercial License 1.0.0"
LICENSE_URL = "https://polyformproject.org/licenses/noncommercial/1.0.0"
REQUIRED_NOTICE = (
    "Required Notice: Copyright 2025-2026 Abraham Borg "
    "(https://github.com/Abe-Borg/Spec_Cleanse)"
)

LICENSE_FILE = "LICENSE.md"
NOTICES_FILE = "THIRD_PARTY_NOTICES.txt"


def about_text() -> str:
    """The body of the About dialog, as plain text."""
    return "\n".join([
        REQUIRED_NOTICE,
        "",
        f"Licensed under the {LICENSE_NAME}:",
        LICENSE_URL,
        "",
        "Use for any noncommercial purpose is permitted. Commercial use, including "
        "billable or client work at an architecture, engineering or construction "
        "firm, needs a separate license from the copyright holder.",
        "",
        f"Commercial licensing, support and bug reports: {SUPPORT_URL}",
        f"Source code and releases: {RELEASES_URL}",
        "",
        "SpecCleanse bundles Python, Tcl/Tk, lxml, PyYAML and the libraries they "
        "are built with, under their own licenses. Third-party notices lists every "
        "one of them, including where to get the source of GNU libiconv, which is "
        "under the LGPL.",
    ])
