"""The GUI's file-level workers, exercised without opening a window."""

import unittest
from unittest.mock import patch

from batch import FileOutcome, ReviewCategory
from tests import docx_builder as db
from tests.support import DocxTestCase

try:
    import gui
    import verify
    GUI_IMPORTABLE = True
except ImportError:  # pragma: no cover - environments without Tk
    GUI_IMPORTABLE = False


DOCUMENT = db.document(
    db.text_para("SECTION 21 13 13 - WET-PIPE SPRINKLER SYSTEMS"),
    db.text_para("PART 1 - GENERAL"),
    db.text_para("Retain subparagraph below for wet-pipe systems."),
    db.text_para("Provide two [Verify quantity with Owner] spare sprinklers."),
    db.para(db.run("Coordinate hangers with structural.", italic=True, color="FF0000")),
    db.text_para("END OF SECTION 21 13 13"),
)


@unittest.skipUnless(GUI_IMPORTABLE, "gui.py needs tkinter, which is unavailable")
class WorkerTests(DocxTestCase):

    def setUp(self):
        super().setUp()
        self.lines: list[str] = []

    def log(self, text: str):
        self.lines.append(text)

    @property
    def output(self) -> str:
        return "\n".join(self.lines)

    def test_preview_reports_every_category(self):
        path = self.build(DOCUMENT)

        self.assertTrue(gui._preview_one(path, self.make_engine(), self.log))
        self.assertIn("REMOVALS — editorial_artifact", self.output)
        self.assertIn("INLINE REDACTIONS", self.output)
        self.assertIn("PRESERVED", self.output)

    def test_preview_labels_a_formatting_only_removal(self):
        # Only reachable through the opt-in: removal on looks alone is off by
        # default, so the italic red aside in DOCUMENT survives a plain run.
        path = self.build(DOCUMENT)
        engine = self.make_engine(specifier_notes={"formatting_only_removal": True})

        self.assertTrue(gui._preview_one(path, engine, self.log))
        self.assertIn("formatting-only", self.output)

    def test_the_default_run_keeps_the_italic_red_aside(self):
        path = self.build(DOCUMENT)

        gui._preview_one(path, self.make_engine(), self.log)

        self.assertNotIn("formatting-only", self.output)
        self.assertNotIn("Coordinate hangers", self.output.split("PRESERVED")[0])

    def test_clean_reports_a_pass(self):
        path = self.build(DOCUMENT)
        out = self.temp_dir / "out.docx"

        outcome = gui._clean_one(path, out, self.make_engine(), self.log)

        self.assertIs(outcome.outcome, FileOutcome.VERIFIED)
        self.assertEqual(outcome.categories, frozenset())
        self.assertTrue(outcome.output_written)
        self.assertIn("PASS", self.output)
        self.assertIn("Paragraphs modified: 1", self.output)
        self.assertTrue(out.exists())

    def test_clean_reports_a_missing_file_without_raising(self):
        out = self.temp_dir / "out.docx"

        outcome = gui._clean_one(
            self.temp_dir / "nope.docx", out, self.make_engine(), self.log
        )

        self.assertIs(outcome.outcome, FileOutcome.FAILED)
        self.assertFalse(outcome.output_written)
        self.assertIn("ERROR", self.output)

    def test_a_failed_verification_is_not_a_success(self):
        # The file is written and the write succeeded; the check did not pass.
        # One Boolean cannot say that, which is how this used to be counted in
        # the "succeeded" total.
        path = self.build(DOCUMENT)
        out = self.temp_dir / "out.docx"
        failing = verify.VerificationResult(input_path=path, output_path=out)
        failing.removed.append(verify.RemovedParagraph("A real requirement.", None))

        with patch.object(gui, "verify_clean", return_value=failing):
            outcome = gui._clean_one(path, out, self.make_engine(), self.log)

        self.assertIs(outcome.outcome, FileOutcome.NEEDS_REVIEW)
        # The verdict has to say which of the four kinds of concern this is:
        # reading the document, fixing the configuration and checking a
        # cross-reference are three different jobs.
        self.assertEqual(outcome.categories, frozenset({ReviewCategory.DETECTED_DAMAGE}))
        self.assertIn("NEEDS REVIEW (1 detected damage)", self.output)
        self.assertIn(str(out), self.output)
        self.assertTrue(out.exists())

    def test_verification_raising_after_a_write_names_the_unverified_file(self):
        path = self.build(DOCUMENT)
        out = self.temp_dir / "out.docx"

        with patch.object(gui, "verify_clean", side_effect=RuntimeError("boom")):
            outcome = gui._clean_one(path, out, self.make_engine(), self.log)

        self.assertIs(outcome.outcome, FileOutcome.FAILED)
        # Failed, and yet a file exists.  The tally needs that as a fact, not
        # only as a line in the log a reader may not scroll back to.
        self.assertTrue(outcome.output_written)
        self.assertIn("UNVERIFIED", self.output)
        self.assertIn(str(out), self.output)
        self.assertTrue(out.exists())

    def test_the_pass_line_does_not_overclaim(self):
        path = self.build(DOCUMENT)
        out = self.temp_dir / "out.docx"

        gui._clean_one(path, out, self.make_engine(), self.log)

        # Verification shares its patterns with the cleaner, so a PASS is a
        # consistency check, not proof the document is intact or that Word
        # will open it.  The wording must not say otherwise.
        self.assertNotIn("no spec content was lost", self.output)
        self.assertNotIn("structure is intact", self.output)

    def test_build_engine_reports_a_broken_config(self):
        bad = self.temp_dir / "patterns.yaml"
        bad.write_text("specifier_notes:\n  text_patterns:\n    - '[unclosed'\n", encoding="utf-8")
        original = gui.CONFIG_PATH
        gui.CONFIG_PATH = bad
        try:
            with self.assertRaises(ValueError):
                gui.build_engine()
        finally:
            gui.CONFIG_PATH = original


if __name__ == "__main__":
    unittest.main()
