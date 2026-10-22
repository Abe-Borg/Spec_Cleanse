"""The About dialog's text, and the documents it opens, agree with LICENSE.md.

PolyForm's Notices clause asks that whoever gets a copy also gets the terms or
their URL, and every ``Required Notice:`` line. The portable executable carries
those through the About dialog and through the notices file attached beside it,
so each of them is held to the line LICENSE.md actually contains.
"""

import re
import unittest
from pathlib import Path

import about
import apppaths

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _license() -> str:
    return (PROJECT_ROOT / about.LICENSE_FILE).read_text(encoding="utf-8")


class RequiredNoticeTests(unittest.TestCase):
    def test_the_required_notice_is_the_line_license_md_has(self):
        lines = [line.strip() for line in _license().splitlines()
                 if line.startswith("Required Notice:")]
        self.assertEqual(lines, [about.REQUIRED_NOTICE])

    def test_the_license_url_is_the_one_license_md_names(self):
        self.assertIn(f"<{about.LICENSE_URL}>", _license())

    def test_the_license_name_is_the_one_license_md_names(self):
        self.assertIn(about.LICENSE_NAME, _license().splitlines()[0])

    def test_the_about_text_carries_the_notice_url_and_contact(self):
        text = about.about_text()
        self.assertIn(about.REQUIRED_NOTICE, text.splitlines())
        self.assertIn(about.LICENSE_URL, text)
        self.assertIn(about.SUPPORT_URL, text)

    def test_the_notices_file_carries_the_notice_and_url(self):
        # The notices travel beside the portable executable, so they are
        # enough on their own to satisfy the Notices clause.
        notices = (PROJECT_ROOT / about.NOTICES_FILE).read_text(encoding="utf-8")
        self.assertIn(about.REQUIRED_NOTICE, notices.splitlines())
        self.assertIn(about.LICENSE_URL, notices)


class BundledDocumentTests(unittest.TestCase):
    def test_both_documents_are_found_from_source(self):
        for name in (about.LICENSE_FILE, about.NOTICES_FILE):
            with self.subTest(name=name):
                path = apppaths.bundled_file(name)
                self.assertEqual(path, PROJECT_ROOT / name)
                self.assertTrue(path.is_file())

    def test_the_names_match_what_the_spec_bundles(self):
        spec = (PROJECT_ROOT / "packaging" / "speccleanse.spec").read_text(encoding="utf-8")
        for name in (about.LICENSE_FILE, about.NOTICES_FILE):
            with self.subTest(name=name):
                self.assertRegex(spec, rf'"{re.escape(name)}"\), "\."\)')


if __name__ == "__main__":
    unittest.main()
