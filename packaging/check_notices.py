#!/usr/bin/env python3
"""
Check THIRD_PARTY_NOTICES.txt against the license files on the build machine.

The release workflow runs this after installing the dependencies and before
PyInstaller runs:

    python packaging/check_notices.py --require-runtime

The notices were assembled from upstream sources at the versions the release
build uses. This is what establishes that they are the texts the build really
ships: each license file below is read from the machine doing the build, and
every one of them must appear in the notices word for word.

- Python's ``LICENSE.txt``, which its Windows build writes by joining the
  Python license with those of the Microsoft C runtime, bzip2, libffi,
  OpenSSL, Tcl, Tk and Tix. It is checked a paragraph at a time, in order,
  because the notices label each of those parts where the installed file
  runs them together; a paragraph is never split between two of them.
- Tk's ``license.terms`` from the Tcl/Tk library Python installs, when present.
- The license files lxml, PyYAML and PyInstaller install into their
  ``.dist-info``, each checked whole.

Whitespace is not compared: line endings differ between a checkout and an
installation, Python's build joins its file with MSBuild, and none of that is
part of a license's text. Every word, in order, is.

It also checks the notices name the versions installed, so upgrading lxml or
PyYAML in ``requirements.txt`` fails here until the notices follow. PyInstaller
is exempt from that one: ``requirements-build.txt`` accepts any 6.x release,
and its text is what matters.

Without ``--require-runtime`` a Python installation with no ``LICENSE.txt`` in
its prefix (a Linux distribution's, for instance) or no PyInstaller installed
is reported and skipped, so the check can still be run for lxml and PyYAML
anywhere. The release workflow passes the flag, so on the build machine
nothing is skipped.
"""

from __future__ import annotations

import argparse
import platform
import re
import sys
from dataclasses import dataclass
from importlib import metadata
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
NOTICES_PATH = PROJECT_ROOT / "THIRD_PARTY_NOTICES.txt"

# Distribution name as pip knows it, and the name the notices give it.
PINNED_PACKAGES = (("lxml", "lxml"), ("PyYAML", "PyYAML"))
UNPINNED_PACKAGES = (("pyinstaller", "PyInstaller"),)

_PARAGRAPH_BREAK = re.compile(r"\n[ \t]*\n")


def words(text: str) -> str:
    """The text with every run of whitespace collapsed to one space."""
    return " ".join(text.split())


def paragraphs(text: str) -> list[str]:
    """The non-empty paragraphs of a text, split at blank lines."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return [p.strip("\n") for p in _PARAGRAPH_BREAK.split(text) if p.strip()]


def missing_whole(notices: str, text: str) -> bool:
    """True when ``text`` does not appear in ``notices`` as one run of words."""
    return words(text) not in words(notices)


def missing_paragraphs(notices: str, text: str) -> list[str]:
    """The paragraphs of ``text`` that do not appear in ``notices``, in order.

    Each paragraph is looked for after the previous one's match, so anything
    may sit between them — the labels do — but they may not be reordered.
    Order is what tells two near-identical texts apart: Tcl's terms and Tk's
    differ in two lines, and without it a paragraph of one would be vouched
    for by the matching paragraph of the other.
    """
    haystack = words(notices)
    missing = []
    position = 0
    for paragraph in paragraphs(text):
        needle = words(paragraph)
        found = haystack.find(needle, position)
        if found < 0:
            missing.append(paragraph)
        else:
            position = found + len(needle)
    return missing


@dataclass
class Report:
    """What a check found. ``problems`` decide the exit status."""

    problems: list[str]
    checked: list[str]
    skipped: list[str]


def license_files(distribution: str) -> list[tuple[str, str]]:
    """``(name, text)`` for each license file a distribution installed.

    Read from its ``.dist-info``: the files under ``licenses/`` where the
    wheel follows PEP 639, and any ``LICENSE*`` or ``COPYING*`` beside
    ``METADATA`` where it does not.
    """
    dist = metadata.distribution(distribution)
    found = []
    for path in dist.files or ():
        parts = path.parts
        if not parts or not parts[0].endswith(".dist-info"):
            continue
        in_licenses = len(parts) > 2 and parts[1] == "licenses"
        top_level = len(parts) == 2 and path.name.upper().startswith(("LICENSE", "COPYING"))
        if in_licenses or top_level:
            found.append(("/".join(parts), path.read_text(encoding="utf-8")))
    return found


def check_package(notices: str, distribution: str, shown_as: str,
                  pinned: bool, report: Report) -> None:
    """Every license file of one installed distribution, and its version."""
    try:
        version = metadata.version(distribution)
    except metadata.PackageNotFoundError:
        report.problems.append(f"{shown_as} is not installed, so its notice cannot be checked")
        return

    if pinned and f"{shown_as} {version}" not in notices:
        report.problems.append(
            f"{shown_as} {version} is installed, but the notices do not name that "
            f"version. Take its license files from the {version} wheel."
        )

    files = license_files(distribution)
    if not files:
        report.problems.append(f"{shown_as} {version} installed no license file to check")
    for name, text in files:
        label = f"{shown_as} {version}: {name}"
        if missing_whole(notices, text):
            report.problems.append(f"{label} is not reproduced in the notices")
        else:
            report.checked.append(label)


def check_runtime(notices: str, require: bool, report: Report) -> None:
    """Python's own LICENSE.txt and Tk's license.terms, from this interpreter."""
    prefix = Path(sys.base_prefix)
    license_txt = prefix / "LICENSE.txt"
    version = platform.python_version()

    if not license_txt.is_file():
        message = f"{license_txt} does not exist, so the Python runtime was not checked"
        (report.problems if require else report.skipped).append(message)
        return

    if f"Python {version}" not in notices:
        report.problems.append(
            f"This is Python {version}, but the notices do not name that version. "
            "Rebuild section 1 from this interpreter's LICENSE.txt."
        )

    lost = missing_paragraphs(notices, license_txt.read_text(encoding="utf-8"))
    if lost:
        first = words(lost[0])
        report.problems.append(
            f"{len(lost)} paragraph(s) of {license_txt} are not reproduced in the "
            f"notices; the first begins: {first[:160]!r}"
        )
    else:
        report.checked.append(f"Python {version}: {license_txt}")

    for terms in sorted(prefix.glob("tcl/*/license.terms")):
        label = f"Python {version}: {terms.relative_to(prefix)}"
        if missing_whole(notices, terms.read_text(encoding="utf-8")):
            report.problems.append(f"{label} is not reproduced in the notices")
        else:
            report.checked.append(label)


def check(notices: str, require_runtime: bool) -> Report:
    """Run every check against the notices text."""
    report = Report(problems=[], checked=[], skipped=[])
    check_runtime(notices, require_runtime, report)
    for distribution, shown_as in PINNED_PACKAGES:
        check_package(notices, distribution, shown_as, pinned=True, report=report)
    for distribution, shown_as in UNPINNED_PACKAGES:
        installed = True
        try:
            metadata.version(distribution)
        except metadata.PackageNotFoundError:
            installed = False
        if installed or require_runtime:
            check_package(notices, distribution, shown_as, pinned=False, report=report)
        else:
            report.skipped.append(f"{shown_as} is not installed, so it was not checked")
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    parser.add_argument(
        "--require-runtime",
        action="store_true",
        help="fail, rather than skip, when Python's LICENSE.txt or PyInstaller "
             "is missing; the release build passes this",
    )
    parser.add_argument("--notices", type=Path, default=NOTICES_PATH)
    args = parser.parse_args(argv)

    notices = args.notices.read_text(encoding="utf-8")
    report = check(notices, args.require_runtime)

    for line in report.checked:
        print(f"ok       {line}")
    for line in report.skipped:
        print(f"skipped  {line}")
    for line in report.problems:
        print(f"PROBLEM  {line}")

    if report.problems:
        print(f"\n{args.notices.name} does not match this machine's license files.")
        return 1
    print(f"\n{args.notices.name} reproduces every license file checked.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
