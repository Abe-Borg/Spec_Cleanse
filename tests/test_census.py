"""The developer censuses, on synthetic documents.

They only read, so the risk is a wrong number rather than a damaged file — and
a wrong number is what a decision would be taken on.
"""

import unittest

from docx_xml import load_config
from tools.census_formatting import (
    census_one,
    collect_paths,
    format_report,
)

from tests import docx_builder as db
from tests.support import CONFIG_PATH, DocxTestCase


ITALIC_RED = dict(italic=True, color="FF0000")


class FormattingCensusTests(DocxTestCase):
    """Census A: what turning formatting-only removal off would cost."""

    def census(self, document_xml):
        return census_one(self.build(document_xml), load_config(CONFIG_PATH))

    def test_a_document_with_no_formatting_only_removals_costs_nothing(self):
        report = self.census(db.document(
            db.text_para("Provide sprinklers throughout."),
            db.text_para("[Specifier: delete this note before issue]"),
        ))

        self.assertIsNone(report.error)
        self.assertEqual(report.removals_with_switch_on, 1)
        self.assertEqual(report.would_newly_survive, 0)
        self.assertEqual(report.share_of_removals, 0.0)

    def test_a_formatting_only_removal_is_counted_as_the_cost_of_the_flip(self):
        report = self.census(db.document(
            db.text_para("Provide sprinklers throughout."),
            db.para(db.run("Coordinate hangers with structural.", **ITALIC_RED)),
        ))

        self.assertEqual(report.removals_with_switch_on, 1)
        self.assertEqual(report.removals_with_switch_off, 0)
        self.assertEqual(report.would_newly_survive, 1)
        self.assertEqual(report.share_of_removals, 1.0)
        self.assertIn("Coordinate hangers with structural.", report.examples)

    def test_a_pattern_match_that_is_also_italic_is_not_the_switch_s_doing(self):
        # The paragraph goes either way, so flipping the switch changes nothing
        # for it.  Counting detections carrying the formatting-only flag would
        # have charged this to the switch and overstated the cost.
        report = self.census(db.document(
            db.para(db.run("[Specifier: delete this note]", **ITALIC_RED)),
        ))

        self.assertEqual(report.removals_with_switch_on, 1)
        self.assertEqual(report.removals_with_switch_off, 1)
        self.assertEqual(report.would_newly_survive, 0)

    def test_repeated_paragraphs_keep_their_multiplicity(self):
        # A spec repeats boilerplate.  Deduplicating would understate the cost.
        note = db.para(db.run("Coordinate with structural.", **ITALIC_RED))
        report = self.census(db.document(note, note, note))

        self.assertEqual(report.would_newly_survive, 3)

    def test_inline_and_preserved_counts_are_reported_for_scale(self):
        report = self.census(db.document(
            db.text_para("SECTION 21 13 13 - WET-PIPE SPRINKLER SYSTEMS"),
            db.text_para("Provide two [Verify quantity with Owner] spare sprinklers."),
        ))

        self.assertEqual(report.inline_redactions, 1)
        self.assertEqual(report.preserved, 1)

    def test_an_unreadable_file_is_recorded_not_raised(self):
        broken = self.temp_dir / "broken.docx"
        broken.write_bytes(b"not a zip")

        report = census_one(broken, load_config(CONFIG_PATH))

        self.assertIsNotNone(report.error)
        self.assertEqual(report.removals_with_switch_on, 0)

    def test_a_document_with_nothing_to_remove_reports_no_share(self):
        report = self.census(db.document(db.text_para("Provide sprinklers.")))

        self.assertEqual(report.removals_with_switch_on, 0)
        self.assertIsNone(report.share_of_removals)


class ReportFormattingTests(DocxTestCase):

    def test_the_report_says_when_it_measured_nothing(self):
        report = census_one(
            self.build(db.document(db.text_para("Provide sprinklers."))),
            load_config(CONFIG_PATH),
        )

        text = format_report([report])

        self.assertIn("says nothing about the switch", text)

    def test_the_report_totals_across_documents(self):
        note = db.para(db.run("Coordinate with structural.", **ITALIC_RED))
        config = load_config(CONFIG_PATH)
        reports = [
            census_one(self.build(db.document(note), name="a.docx"), config),
            census_one(self.build(db.document(note, note), name="b.docx"), config),
        ]

        text = format_report(reports)

        self.assertIn("TOTAL", text)
        self.assertIn("100.0%", text)
        self.assertIn("Text that would newly survive:", text)


class PathCollectionTests(DocxTestCase):

    def test_a_folder_expands_to_the_documents_inside_it(self):
        self.build(db.document(db.text_para("One.")), name="one.docx")
        self.build(db.document(db.text_para("Two.")), name="two.docx")

        found = collect_paths([str(self.temp_dir)])

        self.assertEqual([p.name for p in found], ["one.docx", "two.docx"])

    def test_word_lock_files_are_skipped(self):
        self.build(db.document(db.text_para("Real.")), name="real.docx")
        (self.temp_dir / "~$real.docx").write_bytes(b"")

        found = collect_paths([str(self.temp_dir)])

        self.assertEqual([p.name for p in found], ["real.docx"])

    def test_an_explicit_file_is_taken_as_given(self):
        path = self.build(db.document(db.text_para("One.")), name="one.docx")

        self.assertEqual(collect_paths([str(path)]), [path])


if __name__ == "__main__":
    unittest.main()
