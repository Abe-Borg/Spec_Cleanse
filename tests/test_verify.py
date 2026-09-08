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

    def test_losing_a_whole_paragraph_to_an_inline_pattern_is_unexpected(self):
        """An inline placeholder never excuses losing the requirement around it."""
        source = self.build(db.document(
            db.text_para("Provide two [Verify quantity] spare sprinklers."),
            db.text_para("Other requirement."),
        ))
        cleaned = self.build(
            db.document(db.text_para("Other requirement.")), name="hand_cleaned.docx"
        )
        result = verify_clean(source, cleaned, engine=self.make_engine())

        self.assertEqual(len(result.unexpected_removals), 1, result.removed)
        self.assertFalse(result.passed)

    def test_a_paragraph_of_only_placeholders_may_be_removed_whole(self):
        source = self.build(db.document(
            db.text_para("[Insert product name]"),
            db.text_para("Other requirement."),
        ))
        cleaned = self.build(
            db.document(db.text_para("Other requirement.")), name="hand_cleaned.docx"
        )
        result = verify_clean(source, cleaned, engine=self.make_engine())

        self.assertEqual(len(result.expected_removals), 1)
        self.assertEqual(result.removed[0].category, "inline_placeholder")
        self.assertTrue(result.passed)

    def test_a_trimmed_editorial_run_is_expected_while_the_switch_is_on(self):
        source = self.build(db.document(
            db.para(
                db.run("Keep this. "),
                db.run("Red italic aside.", italic=True, color="FF0000"),
            ),
        ))
        cleaned = self.build(
            db.document(db.text_para("Keep this.")), name="hand_cleaned.docx"
        )
        result = verify_clean(source, cleaned, engine=self.make_engine())

        self.assertEqual(len(result.expected_modifications), 1, result.modified)
        self.assertTrue(result.passed)

    def test_a_trimmed_editorial_run_is_flagged_when_the_switch_is_off(self):
        """The switch has to govern trims too, not just whole-paragraph losses."""
        source = self.build(db.document(
            db.para(
                db.run("Keep this. "),
                db.run("Red italic aside.", italic=True, color="FF0000"),
            ),
        ))
        cleaned = self.build(
            db.document(db.text_para("Keep this.")), name="hand_cleaned.docx"
        )
        engine = self.make_engine(specifier_notes={"formatting_only_removal": False})
        result = verify_clean(source, cleaned, engine=engine)

        self.assertEqual(len(result.unexpected_modifications), 1, result.modified)
        self.assertFalse(result.passed)

    def test_a_trimmed_hidden_run_is_expected_either_way(self):
        source = self.build(db.document(
            db.para(db.run("Keep this. "), db.run("Hidden aside.", vanish=True)),
        ))
        cleaned = self.build(
            db.document(db.text_para("Keep this.")), name="hand_cleaned.docx"
        )
        engine = self.make_engine(specifier_notes={"formatting_only_removal": False})
        result = verify_clean(source, cleaned, engine=engine)

        self.assertEqual(len(result.expected_modifications), 1, result.modified)
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


class LongRedactionPairingTests(DocxTestCase):
    """A correct redaction must pass however little of the paragraph survives."""

    def test_a_long_placeholder_leaves_a_short_survivor(self):
        # 14 characters survive out of 71.  For a pure deletion the similarity
        # ratio is 2*len(after)/(len(before)+len(after)), which drops below
        # 0.5 once more than two-thirds of the characters go — so this correct
        # clean used to be reported as an unexplained removal plus an invented
        # paragraph.
        path = self.build(db.document(db.text_para(
            "Provide [Verify quantity with the Owner and the AHJ prior to bid] units."
        )))
        engine = self.make_engine()
        _, out = self.clean(path, engine)

        self.assertEqual(self.paragraph_texts(out), ["Provide units."])

        result = verify_clean(path, out, engine=engine)

        self.assertTrue(result.passed, [r.text for r in result.unexpected_removals])
        self.assertEqual(result.added, [])
        self.assertEqual(result.removed, [])
        self.assertEqual(len(result.expected_modifications), 1)
        self.assertEqual(
            result.expected_modifications[0].category, "inline_placeholder"
        )

    def test_a_paragraph_of_nothing_but_a_placeholder_still_follows_removal_rules(self):
        path = self.build(db.document(
            db.text_para("[Verify quantity with the Owner prior to bid]"),
            db.text_para("A real requirement."),
        ))
        engine = self.make_engine()
        _, out = self.clean(path, engine)
        result = verify_clean(path, out, engine=engine)

        self.assertTrue(result.passed)
        self.assertEqual(len(result.expected_removals), 1)
        self.assertEqual(result.expected_removals[0].category, "inline_placeholder")

    def test_an_unrelated_survivor_is_not_paired_by_the_exact_rule(self):
        # The exact-result rule must not become a way to pair a removed
        # paragraph with whatever else happens to be nearby.
        path = self.build(db.document(
            db.text_para("Provide [Verify quantity] units."),
            db.text_para("An entirely different requirement."),
        ))
        engine = self.make_engine()
        _, out = self.clean(path, engine)
        result = verify_clean(path, out, engine=engine)

        self.assertTrue(result.passed)
        self.assertEqual(
            self.paragraph_texts(out),
            ["Provide units.", "An entirely different requirement."],
        )


class InjectedDamageTests(DocxTestCase):
    """Damaged outputs built by hand, never by running the cleaner.

    Agreement between a broken verifier and the cleaner that produced its
    input proves nothing, so these construct both sides independently.
    """

    @unittest.expectedFailure
    def test_v04_an_inline_match_must_not_excuse_an_extra_deleted_word(self):
        """V04 — EXPECTED TO FAIL until W03 lands.

        Deleting ``spare`` alongside ``[Verify quantity]`` is currently
        accepted: ``_classify_modification`` asks whether the lost fragment
        *contains* a pattern match, not whether matches *cover* it, so the
        placeholder vouches for the requirement word beside it.

        W03 replaces that predicate with interval coverage and removes this
        decorator.  It is carried as an expected failure rather than a red
        test so the suite stays green through W00-W02 and a genuine
        regression is still visible; unittest reports an unexpected success
        if the verdict ever changes, which is what makes this a tripwire in
        both directions.
        """
        source = self.build(db.document(db.text_para(
            "Provide two [Verify quantity] spare filters per unit."
        )), name="v04_in.docx")
        damaged = self.build(db.document(db.text_para(
            "Provide two filters per unit."
        )), name="v04_out.docx")

        result = verify_clean(source, damaged, engine=self.make_engine())

        self.assertFalse(
            result.passed,
            "losing 'spare' is not something an inline placeholder authorises",
        )


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
