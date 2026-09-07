"""Verification: what it classifies, and the damage it can see on its own."""

import unittest

from verify import (
    FORMATTING_BASED,
    PRESERVE_VIOLATION,
    extract_paragraphs,
    inspect_structure,
    lint_structure,
    verify_clean,
)

from tests import docx_builder as db
from tests.support import DocxTestCase


class ExtractionTests(DocxTestCase):
    """Both sides of the comparison must measure the document the same way."""

    def test_text_box_content_is_counted_once(self):
        path = self.build(db.document(
            db.text_para("Body paragraph."),
            db.para(db.run("Before box. "), db.text_box(db.text_para("Inside the box."))),
        ))
        texts = [p.text for p in extract_paragraphs(path, self.config)]

        self.assertEqual(texts, ["Body paragraph.", "Before box.", "Inside the box."])

    def test_tabs_and_breaks_become_whitespace(self):
        path = self.build(db.document(
            db.para(db.run("PART 1", inner="<w:tab/>"), db.run("GENERAL")),
        ))
        texts = [p.text for p in extract_paragraphs(path, self.config)]

        self.assertEqual(texts, ["PART 1\tGENERAL"])


class ClassificationTests(DocxTestCase):

    def test_pattern_removal_is_expected(self):
        path = self.build(db.document(
            db.text_para("Real requirement."),
            db.text_para("[Specifier: delete before issue]"),
        ))
        engine = self.make_engine()
        _, out = self.clean(path, engine)
        result = verify_clean(path, out, engine=engine)

        self.assertTrue(result.passed, [r.text for r in result.unexpected_removals])
        self.assertEqual(len(result.expected_removals), 1)
        self.assertEqual(result.expected_removals[0].category, "specifier_note")

    def test_low_confidence_removal_is_expected(self):
        """A removal that needed formatting to cross the threshold still matches a rule."""
        path = self.build(db.document(
            db.text_para("Real requirement."),
            db.para(db.run("Revise as required for the project.", italic=True, color="C00000")),
        ))
        engine = self.make_engine()
        _, out = self.clean(path, engine)
        result = verify_clean(path, out, engine=engine)

        self.assertEqual(len(result.removed), 1)
        self.assertEqual(result.removed[0].category, "editorial_artifact")
        self.assertTrue(result.passed)

    def test_formatting_only_removal_is_expected_while_the_switch_is_on(self):
        path = self.build(db.document(
            db.text_para("Real requirement."),
            db.para(db.run("Coordinate with structural.", italic=True, color="FF0000")),
        ))
        engine = self.make_engine()
        _, out = self.clean(path, engine)
        result = verify_clean(path, out, engine=engine)

        self.assertEqual(len(result.removed), 1)
        self.assertEqual(result.removed[0].category, FORMATTING_BASED)
        self.assertTrue(result.passed)

    def test_formatting_only_removal_is_flagged_when_the_switch_is_off(self):
        """The same loss, judged by a configuration that never asked for it."""
        source = self.build(db.document(
            db.text_para("Real requirement."),
            db.para(db.run("Coordinate with structural.", italic=True, color="FF0000")),
        ))
        cleaned = self.build(
            db.document(db.text_para("Real requirement.")), name="hand_cleaned.docx"
        )
        engine = self.make_engine(specifier_notes={"formatting_only_removal": False})
        result = verify_clean(source, cleaned, engine=engine)

        self.assertEqual(len(result.unexpected_removals), 1)
        self.assertFalse(result.passed)

    def test_inline_redaction_is_an_expected_modification(self):
        path = self.build(db.document(
            db.text_para("Provide two [Verify quantity with Owner] spare sprinklers."),
        ))
        engine = self.make_engine()
        _, out = self.clean(path, engine)
        result = verify_clean(path, out, engine=engine)

        self.assertEqual(len(result.removed), 0)
        self.assertEqual(len(result.expected_modifications), 1)
        self.assertEqual(result.modified[0].category, "inline_placeholder")
        self.assertTrue(result.passed)

    def test_preserve_violation_is_reported(self):
        source = self.build(db.document(
            db.text_para("PART 1 - GENERAL"),
            db.text_para("Real requirement."),
        ))
        cleaned = self.build(
            db.document(db.text_para("Real requirement.")), name="hand_cleaned.docx"
        )
        result = verify_clean(source, cleaned, engine=self.make_engine())

        self.assertEqual(len(result.preserve_violations), 1)
        self.assertEqual(result.preserve_violations[0].category, PRESERVE_VIOLATION)
        self.assertFalse(result.passed)

    def test_a_mutated_paragraph_is_never_expected(self):
        """Only deletions are legitimate; changed text is always a red flag."""
        source = self.build(db.document(
            db.text_para("Provide two spare sprinklers of each type."),
        ))
        cleaned = self.build(
            db.document(db.text_para("Provide three spare sprinklers of each type.")),
            name="hand_cleaned.docx",
        )
        result = verify_clean(source, cleaned, engine=self.make_engine())

        self.assertFalse(result.passed)
        self.assertEqual(len(result.unexpected_removals), 1)
        self.assertEqual(len(result.added), 1)

    def test_unexplained_trim_of_a_surviving_paragraph_is_flagged(self):
        source = self.build(db.document(
            db.text_para("Provide two spare sprinklers of each type installed."),
        ))
        cleaned = self.build(
            db.document(db.text_para("Provide two spare sprinklers.")),
            name="hand_cleaned.docx",
        )
        result = verify_clean(source, cleaned, engine=self.make_engine())

        self.assertEqual(len(result.unexpected_modifications), 1)
        self.assertFalse(result.passed)


class StructuralTests(DocxTestCase):
    """The one layer that can see damage no pattern describes."""

    def test_lint_flags_an_empty_table_cell(self):
        path = self.build(db.document(
            '<w:tbl><w:tblPr/><w:tr><w:tc><w:tcPr/></w:tc></w:tr></w:tbl>',
        ))
        issues = lint_structure(path)

        self.assertTrue(
            any("w:tc" in issue for issue in issues),
            f"empty cell not reported: {issues}",
        )

    def test_lint_flags_a_cell_not_ending_in_a_paragraph(self):
        path = self.build(db.document(
            db.table(db.text_para("Keep.") + db.table(db.text_para("Nested."))),
        ))
        issues = lint_structure(path)

        self.assertTrue(
            any("does not end with a paragraph" in issue for issue in issues),
            f"cell ending in a table not reported: {issues}",
        )

    def test_lint_flags_an_unbalanced_field(self):
        path = self.build(db.document(
            db.para(db.field_begin(), db.run("result")),
        ))
        issues = lint_structure(path)

        self.assertTrue(
            any("field characters" in issue for issue in issues),
            f"unbalanced field not reported: {issues}",
        )

    def test_a_clean_document_lints_clean(self):
        path = self.build(db.document(
            db.text_para("Body."),
            db.table(db.text_para("Cell.")),
        ))

        self.assertEqual(lint_structure(path), [])

    def test_lost_section_break_is_a_violation(self):
        source = self.build(db.document(
            db.para(db.run("Section one."), sect=True),
            db.text_para("Section two."),
        ))
        cleaned = self.build(
            db.document(db.text_para("Section one."), db.text_para("Section two.")),
            name="hand_cleaned.docx",
        )
        result = verify_clean(source, cleaned, engine=self.make_engine())

        self.assertTrue(
            any("section break" in str(v) for v in result.structural),
            f"lost section break not reported: {result.structural}",
        )
        self.assertFalse(result.passed)

    def test_emptied_footer_is_a_violation(self):
        source = self.build(
            db.document(db.text_para("Body.")),
            {"word/footer1.xml": db.footer(db.text_para("© 2026 ARCOM."))},
        )
        cleaned = self.build(
            db.document(db.text_para("Body.")),
            {"word/footer1.xml": db.footer("")},
            name="hand_cleaned.docx",
        )
        result = verify_clean(source, cleaned, engine=self.make_engine())

        self.assertTrue(
            any("w:ftr" in str(v) for v in result.structural),
            f"emptied footer not reported: {result.structural}",
        )

    def test_damage_already_in_the_input_is_not_blamed_on_the_clean(self):
        broken = '<w:tbl><w:tblPr/><w:tr><w:tc><w:tcPr/></w:tc></w:tr></w:tbl>'
        source = self.build(db.document(broken, db.text_para("[Specifier: note]")))
        engine = self.make_engine()
        _, out = self.clean(source, engine)
        result = verify_clean(source, out, engine=engine)

        self.assertEqual(result.structural, [])
        self.assertTrue(inspect_structure(source).issues)


if __name__ == "__main__":
    unittest.main()
