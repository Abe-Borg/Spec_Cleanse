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


    def test_a_configuration_notice_makes_a_clean_file_need_review(self):
        # Nothing is wrong with the output: the comparison passes.  What is
        # worth knowing is that these rules can remove text on no content
        # evidence at all, so agreeing with them is not a reason to hand the
        # file on unread.
        path = self.build(DOCUMENT)
        out = self.temp_dir / "out.docx"
        engine = self.make_engine(specifier_notes={"formatting_only_removal": True})

        outcome = gui._clean_one(
            path, out, engine, self.log, configuration_notice=True
        )

        self.assertIs(outcome.outcome, FileOutcome.NEEDS_REVIEW)
        self.assertEqual(outcome.categories, frozenset({ReviewCategory.CONFIGURATION}))
        self.assertIn("NEEDS REVIEW (1 configuration notice)", self.output)
        self.assertTrue(outcome.output_written)

    def test_the_same_file_without_the_notice_verifies(self):
        # The guard against the case above being satisfied by a check that
        # simply never reports Verified.
        path = self.build(DOCUMENT)
        out = self.temp_dir / "out.docx"

        outcome = gui._clean_one(path, out, self.make_engine(), self.log)

        self.assertIs(outcome.outcome, FileOutcome.VERIFIED)

    def test_a_configuration_notice_is_added_to_real_findings(self):
        # Two independent concerns.  Showing only one would be the pooling the
        # category split exists to end.
        path = self.build(DOCUMENT)
        out = self.temp_dir / "out.docx"
        failing = verify.VerificationResult(input_path=path, output_path=out)
        failing.removed.append(verify.RemovedParagraph("A real requirement.", None))

        with patch.object(gui, "verify_clean", return_value=failing):
            outcome = gui._clean_one(
                path, out, self.make_engine(), self.log, configuration_notice=True
            )

        self.assertEqual(outcome.categories, frozenset({
            ReviewCategory.CONFIGURATION, ReviewCategory.DETECTED_DAMAGE,
        }))

    def test_a_failing_verification_always_names_at_least_one_category(self):
        # The anti-drift guard.  _clean_one takes `passed` as the authority and
        # the categories as the explanation; if something new ever contributes
        # to `passed` without a category to match, a real failure would be
        # reported with an empty reason.  Every contributor is checked here.
        path = self.build(DOCUMENT)
        out = self.temp_dir / "out.docx"
        cases = {
            "unexpected removal": lambda r: r.removed.append(
                verify.RemovedParagraph("A real requirement.", None)),
            "preserve violation": lambda r: r.removed.append(
                verify.RemovedParagraph("PART 1 - GENERAL", verify.PRESERVE_VIOLATION)),
            "unexpected modification": lambda r: r.modified.append(
                verify.ModifiedParagraph("before", "be", ["fore"])),
            "structural": lambda r: r.structural.append(
                verify.StructuralViolation("a footer left empty")),
            "reference": lambda r: r.structural.append(
                verify.StructuralViolation("reference broken", kind="reference")),
            "added": lambda r: r.added.append("Provide gold sprinklers."),
            "numbering": lambda r: r.numbering.append(
                verify.NumberingNotice("word/document.xml", "1", "an item")),
        }

        for name, damage in cases.items():
            with self.subTest(name):
                result = verify.VerificationResult(input_path=path, output_path=out)
                damage(result)

                self.assertFalse(result.passed, f"{name} should not pass")
                self.assertTrue(
                    result.review_categories(),
                    f"{name} fails verification but names no category",
                )


    def test_an_earlier_output_is_not_reported_as_this_run_writing(self):
        # A destination left by an earlier good run exists whether or not this
        # run wrote anything.  Reporting it as this run's unverified output is
        # an invitation to delete a document that is perfectly fine.
        out = self.temp_dir / "spec_cleaned.docx"
        out.write_bytes(b"a perfectly good earlier output")
        bad = self.temp_dir / "spec.docx"
        bad.write_bytes(b"this is not a zip")

        outcome = gui._clean_one(bad, out, self.make_engine(), self.log)

        self.assertIs(outcome.outcome, FileOutcome.FAILED)
        self.assertFalse(outcome.output_written)
        self.assertNotIn("UNVERIFIED", self.output)
        self.assertEqual(out.read_bytes(), b"a perfectly good earlier output")

    def test_a_write_this_run_did_make_is_still_reported(self):
        # The guard against the case above being satisfied by never reporting a
        # write at all.
        path = self.build(DOCUMENT)
        out = self.temp_dir / "out.docx"

        with patch.object(gui, "verify_clean", side_effect=RuntimeError("boom")):
            outcome = gui._clean_one(path, out, self.make_engine(), self.log)

        self.assertTrue(outcome.output_written)
        self.assertIn("UNVERIFIED", self.output)


if __name__ == "__main__":
    unittest.main()
