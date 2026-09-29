"""What the Windows packaging reads from the repository, checked without Windows.

The installer and the executable are only ever built on a Windows runner, by
the release workflow — which runs on a pull request only when the packaging or
what goes into it changes. A file the installer names being renamed or deleted
by any other change would go unnoticed until the next tag failed to build, or,
for the release upload that only a tag runs, until after it had.

So the build definitions are read here as text, and the suite checks them on
every run, anywhere. The notices check that the release workflow runs against
its own machine is tested here too, on the parts that are the same everywhere.
"""

import importlib.util
import platform
import re
import sys
import unittest
from pathlib import Path, PureWindowsPath
from tempfile import TemporaryDirectory
from unittest import mock

import yaml

PROJECT_ROOT = Path(__file__).resolve().parent.parent
PACKAGING = PROJECT_ROOT / "packaging"
INSTALLER = PACKAGING / "installer.iss"
RELEASE_WORKFLOW = PROJECT_ROOT / ".github" / "workflows" / "release.yml"
NOTICES = "THIRD_PARTY_NOTICES.txt"


def _load_check_notices():
    """Import packaging/check_notices.py, which is a script, not a package."""
    path = PACKAGING / "check_notices.py"
    spec = importlib.util.spec_from_file_location("check_notices", path)
    module = importlib.util.module_from_spec(spec)
    # Registered before it runs: dataclasses look their module up here.
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


check_notices = _load_check_notices()


def _installer_text() -> str:
    return INSTALLER.read_text(encoding="utf-8")


def _defines(script: str) -> dict[str, str]:
    """The script's ``#define Name "value"`` preprocessor constants."""
    return dict(re.findall(r'^\s*#define\s+(\w+)\s+"([^"]*)"', script, flags=re.M))


def _section(script: str, name: str) -> list[str]:
    """The entry lines of one ``[Section]``, without comments or blank lines."""
    lines, inside = [], False
    for line in script.splitlines():
        stripped = line.strip()
        if re.fullmatch(r"\[[^\]]+\]", stripped):
            inside = stripped.lower() == f"[{name.lower()}]"
            continue
        if inside and stripped and not stripped.startswith(";"):
            lines.append(stripped)
    return lines


def _resolve(value: str, defines: dict[str, str]) -> Path:
    """An installer source path as the compiler would find it.

    Relative paths are relative to the script's own directory, and are written
    with backslashes. Preprocessor constants are substituted; anything still in
    braces afterwards is an Inno Setup constant, which names nothing in the
    repository, and is refused rather than guessed at.
    """
    for name, replacement in defines.items():
        value = value.replace("{#" + name + "}", replacement)
    if "{" in value:
        raise ValueError(f"cannot resolve {value!r} to a repository path")
    return PACKAGING.joinpath(*PureWindowsPath(value).parts).resolve()


class InstallerSourceTests(unittest.TestCase):
    """Every file the installer takes from the repository must exist.

    The one exception is the executable, which PyInstaller builds into dist\\
    just before the installer is compiled, so it never exists in a checkout.
    """

    def setUp(self):
        self.script = _installer_text()
        self.defines = _defines(self.script)
        self.sources = [
            re.search(r'\bSource:\s*"([^"]*)"', line).group(1)
            for line in _section(self.script, "Files")
            if re.search(r'\bSource:\s*"', line)
        ]

    def _is_built_executable(self, source: str) -> bool:
        return "{#AppExeName}" in source

    def test_the_files_section_was_read(self):
        # A parser that found nothing would pass every test below.
        self.assertGreaterEqual(len(self.sources), 2)
        built = [s for s in self.sources if self._is_built_executable(s)]
        self.assertEqual(len(built), 1, "exactly one source is the built executable")

    def test_every_other_source_exists(self):
        for source in self.sources:
            if self._is_built_executable(source):
                continue
            with self.subTest(source=source):
                path = _resolve(source, self.defines)
                self.assertTrue(
                    path.is_file(),
                    f"installer.iss installs {source}, which resolves to {path} "
                    "and does not exist; renaming it breaks the installer build",
                )

    def test_the_license_page_exists(self):
        # Not a [Files] entry, but compiled in just the same: the wizard shows it.
        match = re.search(r"^\s*LicenseFile\s*=\s*(.+?)\s*$", self.script, flags=re.M)
        self.assertIsNotNone(match)
        self.assertTrue(_resolve(match.group(1), self.defines).is_file())

    def test_the_notices_are_installed(self):
        installed = {_resolve(s, self.defines) for s in self.sources
                     if not self._is_built_executable(s)}
        self.assertIn(PROJECT_ROOT / NOTICES, installed)


class ReleaseWorkflowTests(unittest.TestCase):
    """The portable executable has no installer, so the release carries its notices."""

    def setUp(self):
        workflow = yaml.safe_load(RELEASE_WORKFLOW.read_text(encoding="utf-8"))
        # YAML 1.1 reads the bare key `on` as the Boolean true.
        self.triggers = workflow.get("on", workflow.get(True))
        self.steps = {step.get("name"): step for step in workflow["jobs"]["windows"]["steps"]}

    def test_the_notices_are_uploaded_with_both_executables(self):
        artifact = self.steps["Upload build artifacts"]["with"]["path"]
        release = self.steps["Attach the assets to the release"]["run"]
        staging = self.steps["Stage the portable executable and the notices"]["run"]

        self.assertIn(f"cp {NOTICES} dist/", staging)
        self.assertIn(f"dist/{NOTICES}", artifact)
        self.assertIn(f"dist/{NOTICES}", release)
        self.assertTrue((PROJECT_ROOT / NOTICES).is_file())

    def test_editing_the_notices_builds_the_release(self):
        # Editing the notices has to run the check that compares them with
        # what the runner installs, which only this workflow does.
        self.assertIn(NOTICES, self.triggers["pull_request"]["paths"])

    def test_the_notices_are_checked_before_anything_is_built(self):
        names = list(self.steps)
        check = "Check the third-party notices against this machine"
        self.assertIn("--require-runtime", self.steps[check]["run"])
        self.assertLess(names.index(check), names.index("Build the executable"))


class NoticesComparisonTests(unittest.TestCase):
    """How check_notices decides a license text is reproduced."""

    def test_whitespace_is_not_compared_but_words_are(self):
        notices = "Header\n\nPermission is  hereby\ngranted,\r\nfree of charge.\n"
        self.assertFalse(check_notices.missing_whole(
            notices, "Permission is hereby granted, free of charge."))
        self.assertTrue(check_notices.missing_whole(
            notices, "Permission is hereby granted, free of all charge."))

    def test_paragraphs_may_be_labelled_but_not_reordered(self):
        installed = "First part.\r\n\r\nSecond part.\r\n\r\nThird part.\r\n"
        labelled = ("--- one ---\n\nFirst part.\n\n--- two ---\n\nSecond part.\n\n"
                    "--- three ---\n\nThird part.\n")
        reordered = "Second part.\n\nFirst part.\n\nThird part.\n"

        self.assertEqual(check_notices.missing_paragraphs(labelled, installed), [])
        # "First part." matches after the notices' "Second part.", so it is the
        # second paragraph that can no longer be found in order.
        self.assertEqual(check_notices.missing_paragraphs(reordered, installed),
                         ["Second part."])

    def test_a_paragraph_is_not_vouched_for_by_an_earlier_twin(self):
        # Tcl's terms and Tk's are near-identical. A paragraph that matches only
        # an *earlier* part of the notices must not count.
        installed = "Tcl terms.\n\nTk terms.\n\nTcl terms.\n"
        notices = "Tcl terms.\n\nTk terms.\n\nTk terms.\n"
        self.assertEqual(check_notices.missing_paragraphs(notices, installed),
                         ["Tcl terms."])


class InstalledPackageNoticesTests(unittest.TestCase):
    """The notices reproduce the lxml and PyYAML this suite runs against.

    Both are pinned in requirements.txt, so every environment that installed
    them has identical license files, and this holds on every platform. That
    makes an upgrade fail here until the notices are rebuilt from the new
    wheel. The Python and PyInstaller checks depend on the machine and run in
    the release workflow instead.
    """

    def setUp(self):
        self.notices = (PROJECT_ROOT / NOTICES).read_text(encoding="utf-8")

    def test_lxml_and_pyyaml_are_reproduced_at_the_installed_version(self):
        report = check_notices.Report(problems=[], checked=[], skipped=[])
        for distribution, shown_as in check_notices.PINNED_PACKAGES:
            check_notices.check_package(self.notices, distribution, shown_as,
                                        pinned=True, report=report)

        self.assertEqual(report.problems, [])
        # lxml installs two license files and PyYAML one; none may be skipped.
        self.assertGreaterEqual(len(report.checked), 3)

    def test_a_version_the_notices_do_not_name_is_a_problem(self):
        without_version = self.notices.replace("lxml 6.0.2", "lxml")
        report = check_notices.Report(problems=[], checked=[], skipped=[])
        check_notices.check_package(without_version, "lxml", "lxml",
                                    pinned=True, report=report)
        self.assertTrue(any("do not name that version" in p for p in report.problems))

    def test_a_changed_license_text_is_a_problem(self):
        altered = self.notices.replace("Copyright (c) 2004 Infrae.", "Copyright Infrae.")
        report = check_notices.Report(problems=[], checked=[], skipped=[])
        check_notices.check_package(altered, "lxml", "lxml", pinned=True, report=report)
        self.assertTrue(any("LICENSE.txt is not reproduced" in p for p in report.problems))


class RuntimeNoticesTests(unittest.TestCase):
    """The Python runtime check, against a fake installation prefix."""

    def setUp(self):
        self._tmp = TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.prefix = Path(self._tmp.name)
        patcher = mock.patch.object(sys, "base_prefix", str(self.prefix))
        patcher.start()
        self.addCleanup(patcher.stop)
        self.version = platform.python_version()

    def _install(self, license_txt: str, tk_terms: str):
        (self.prefix / "LICENSE.txt").write_text(license_txt, encoding="utf-8")
        tk = self.prefix / "tcl" / "tk8.6"
        tk.mkdir(parents=True)
        (tk / "license.terms").write_text(tk_terms, encoding="utf-8")

    def _run(self, notices: str, require: bool):
        report = check_notices.Report(problems=[], checked=[], skipped=[])
        check_notices.check_runtime(notices, require, report)
        return report

    def test_a_missing_license_is_skipped_unless_required(self):
        self.assertEqual(len(self._run("", require=False).skipped), 1)
        self.assertEqual(self._run("", require=False).problems, [])
        self.assertEqual(len(self._run("", require=True).problems), 1)

    def test_the_python_license_and_tk_terms_are_both_read(self):
        self._install("PSF terms.\n\nTcl terms.\n\nTk terms.\n", "Tk terms.\n")
        notices = (f"1. Python {self.version}\n\n--- LICENSE ---\n\nPSF terms.\n\n"
                   "--- Tcl ---\n\nTcl terms.\n\n--- Tk ---\n\nTk terms.\n")

        report = self._run(notices, require=True)

        self.assertEqual(report.problems, [])
        self.assertEqual(len(report.checked), 2)

    def test_a_lost_paragraph_or_a_different_python_is_a_problem(self):
        self._install("PSF terms.\n\nTcl terms.\n", "Tk terms.\n")
        notices = "1. Python 0.0.1\n\nPSF terms.\n\nTk terms.\n"

        problems = self._run(notices, require=True).problems

        self.assertTrue(any("do not name that version" in p for p in problems))
        self.assertTrue(any("1 paragraph(s)" in p and "Tcl terms." in p for p in problems))


if __name__ == "__main__":
    unittest.main()
