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
from tools.census_references import (
    census_one as reference_census_one,
    field_target,
    format_report as reference_format_report,
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

    def test_two_rules_on_one_paragraph_count_as_one_removal(self):
        # The share reported here is what a default decision gets taken on.
        # Counting detections rather than removed paragraphs reported 3
        # removals and 33% for this document instead of 2 and 50%.
        report = self.census(db.document(
            db.text_para("[Specifier: Copyright 2026 ARCOM]"),
            db.para(db.run("Coordinate hangers with structural.", **ITALIC_RED)),
        ))

        self.assertEqual(report.removals_with_switch_on, 2)
        self.assertEqual(report.would_newly_survive, 1)
        self.assertEqual(report.share_of_removals, 0.5)

    def test_a_placeholder_only_paragraph_counts_as_a_removal(self):
        report = self.census(db.document(
            db.text_para("[Verify quantity with Owner]"),
        ))

        self.assertEqual(report.removals_with_switch_on, 1)
        self.assertEqual(report.inline_redactions, 0)

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




class FieldInstructionTests(unittest.TestCase):
    """The reader that says which bookmark a field points at."""

    def test_a_plain_reference(self):
        self.assertEqual(field_target(" REF Target \\h "), "Target")

    def test_a_quoted_name_may_contain_spaces(self):
        self.assertEqual(field_target(' REF "Part 1 General" \\h '), "Part 1 General")

    def test_pageref_and_noteref_name_bookmarks_too(self):
        self.assertEqual(field_target(" PAGEREF _Toc12345 \\h "), "_Toc12345")
        self.assertEqual(field_target(" NOTEREF _Ref99 \\h "), "_Ref99")

    def test_an_unrelated_field_names_no_bookmark(self):
        self.assertIsNone(field_target(" PAGE "))
        self.assertIsNone(field_target(" TOC \\o \"1-3\" "))

    def test_prose_containing_the_word_ref_is_not_a_field(self):
        # Only instructions are read.  This string would never reach the reader
        # from visible text, and must not resolve if it somehow did.
        self.assertIsNone(field_target("Refer to the drawings for REF details"))

    def test_an_empty_instruction_names_nothing(self):
        self.assertIsNone(field_target(""))
        self.assertIsNone(field_target(None))


REF_FIELD = '<w:fldSimple w:instr=" REF Target \\h "><w:r><w:t>3.2</w:t></w:r></w:fldSimple>'


class ReferenceCensusTests(DocxTestCase):
    """Census B: how much of a clean sits inside a referenced bookmark range."""

    def census(self, document_xml):
        return reference_census_one(self.build(document_xml), load_config(CONFIG_PATH))

    def test_a_document_with_no_bookmarks_measures_zero_overlap(self):
        report = self.census(db.document(
            db.text_para("Provide sprinklers throughout."),
            db.text_para("[Specifier: delete this note before issue]"),
        ))

        self.assertIsNone(report.error)
        self.assertEqual(report.bookmarks, 0)
        self.assertEqual(report.removable_paragraphs, 1)
        self.assertEqual(report.removable_inside_referenced_range, 0)
        self.assertEqual(report.share_inside, 0.0)

    def test_a_removal_inside_a_referenced_range_is_counted(self):
        note = (db.bookmark_start("1", "Target")
                + db.run("[Specifier: delete this note before issue]")
                + db.bookmark_end("1"))
        report = self.census(db.document(
            f"<w:p>{note}</w:p>",
            f'<w:p><w:r><w:t>See </w:t></w:r>{REF_FIELD}</w:p>',
        ))

        self.assertEqual(report.bookmarks, 1)
        self.assertEqual(report.referenced_bookmarks, 1)
        self.assertEqual(report.removable_paragraphs, 1)
        self.assertEqual(report.removable_inside_referenced_range, 1)
        self.assertEqual(report.share_inside, 1.0)

    def test_an_unreferenced_bookmark_protects_nothing(self):
        # Word bookmarks far more than it cross-references.  Counting every
        # bookmark would be the measurement that makes retention look
        # impossible when it may not be.
        note = (db.bookmark_start("1", "Unused")
                + db.run("[Specifier: delete this note before issue]")
                + db.bookmark_end("1"))
        report = self.census(db.document(f"<w:p>{note}</w:p>"))

        self.assertEqual(report.bookmarks, 1)
        self.assertEqual(report.referenced_bookmarks, 0)
        self.assertEqual(report.removable_inside_referenced_range, 0)

    def test_an_internal_hyperlink_counts_as_a_consumer(self):
        note = (db.bookmark_start("1", "target")
                + db.run("[Specifier: delete this note before issue]")
                + db.bookmark_end("1"))
        report = self.census(db.document(
            f"<w:p>{note}</w:p>",
            f'<w:p>{db.hyperlink(db.run("jump"), anchor="target")}</w:p>',
        ))

        self.assertEqual(report.referenced_bookmarks, 1)
        self.assertEqual(report.removable_inside_referenced_range, 1)

    def test_a_complex_field_split_across_runs_is_read(self):
        note = (db.bookmark_start("1", "Target")
                + db.run("[Specifier: delete this note before issue]")
                + db.bookmark_end("1"))
        split = (
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> REF </w:instrText></w:r>'
            '<w:r><w:instrText xml:space="preserve">Target \\h </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>3.2</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        )
        report = self.census(db.document(f"<w:p>{note}</w:p>", f"<w:p>{split}</w:p>"))

        self.assertEqual(report.referenced_bookmarks, 1)
        self.assertEqual(report.removable_inside_referenced_range, 1)

    def test_a_range_spanning_several_paragraphs_covers_all_of_them(self):
        report = self.census(db.document(
            f'<w:p>{db.bookmark_start("1", "Target")}{db.run("Keep this requirement.")}</w:p>',
            db.text_para("[Specifier: delete this note before issue]"),
            f'<w:p>{db.run("Retain subparagraph below for wet-pipe systems.")}{db.bookmark_end("1")}</w:p>',
            f'<w:p><w:r><w:t>See </w:t></w:r>{REF_FIELD}</w:p>',
        ))

        self.assertEqual(report.removable_paragraphs, 2)
        self.assertEqual(report.removable_inside_referenced_range, 2)

    def test_a_removal_outside_the_range_is_not_counted(self):
        report = self.census(db.document(
            f'<w:p>{db.bookmark_start("1", "Target")}{db.run("Keep this.")}{db.bookmark_end("1")}</w:p>',
            db.text_para("[Specifier: delete this note before issue]"),
            f'<w:p><w:r><w:t>See </w:t></w:r>{REF_FIELD}</w:p>',
        ))

        self.assertEqual(report.removable_paragraphs, 1)
        self.assertEqual(report.removable_inside_referenced_range, 0)

    def test_a_half_open_bookmark_covers_nothing(self):
        # An unterminated range is not evidence about any paragraph.
        report = self.census(db.document(
            f'<w:p>{db.bookmark_start("1", "Target")}'
            f'{db.run("[Specifier: delete this note before issue]")}</w:p>',
            f'<w:p><w:r><w:t>See </w:t></w:r>{REF_FIELD}</w:p>',
        ))

        self.assertEqual(report.removable_inside_referenced_range, 0)

    def test_an_unreadable_file_is_recorded_not_raised(self):
        broken = self.temp_dir / "broken.docx"
        broken.write_bytes(b"not a zip")

        report = reference_census_one(broken, load_config(CONFIG_PATH))

        self.assertIsNotNone(report.error)

    def test_the_report_says_when_nothing_was_referenced(self):
        report = self.census(db.document(
            db.text_para("[Specifier: delete this note before issue]"),
        ))

        self.assertIn("retention would suppress nothing here",
                      reference_format_report([report]))


if __name__ == "__main__":
    unittest.main()
