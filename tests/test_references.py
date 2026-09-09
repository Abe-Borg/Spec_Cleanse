"""A reference the clean broke is reported; one it did not is left alone.

A surviving ``REF`` names a *specific* missing bookmark, so a newly broken
reference is a decidable fact — unlike numbering, whose consequences resist
precise attribution.  That is what makes detection worth doing here.

What this does **not** do is repair anything.  No target is invented, no field
is retargeted, no replacement bookmark is created, and no field result is
rewritten.  Leaving an empty bookmark behind would not count as fixing
reference integrity either: it suppresses one error message while letting the
``REF`` return misleading content.  The output is still written; the run is
reported as needing review.
"""

import unittest

from docx_xml import reference_target
from processor import DocxProcessor
from verify import verify_clean

from tests import docx_builder as db
from tests.support import DocxTestCase


NOTE = "Note to Specifier: delete this paragraph."


def simple_ref(name=None, instruction=None):
    """A ``w:fldSimple`` reference.  Quotes are escaped for the attribute."""
    body = (instruction or f" REF {name} ").replace('"', "&quot;")
    return ('<w:p><w:r><w:t>See </w:t></w:r>'
            f'<w:fldSimple w:instr="{body}"><w:r><w:t>above</w:t></w:r>'
            '</w:fldSimple><w:r><w:t> for details.</w:t></w:r></w:p>')


def complex_ref(name):
    """The same reference the long way, with the instruction split across runs."""
    return ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> REF </w:instrText></w:r>'
            f'<w:r><w:instrText xml:space="preserve">{name} \\h </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>above</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')


def bookmarked(name, *children, bid="1"):
    return db.para(db.bookmark_start(bid, name), *children, db.bookmark_end(bid))


class InstructionGrammar(unittest.TestCase):
    """Only enough grammar to say which bookmark a reference consumes."""

    def test_supported_keywords_yield_their_target(self):
        for instruction, expected in (
            (" REF TargetA \\h ", "TargetA"),
            (" PAGEREF Target \\* MERGEFORMAT ", "Target"),
            (" NOTEREF fn1 ", "fn1"),
            (' REF "Target A" \\h ', "Target A"),
        ):
            with self.subTest(instruction):
                self.assertEqual(reference_target(instruction), expected)

    def test_anything_else_refers_to_no_bookmark(self):
        for instruction in (
            " PAGE ", " TOC \\o \"1-3\" ", " REF ", " REF \\h ", "",
            "Refer to REF above",   # prose never reaches here, but say so anyway
        ):
            with self.subTest(instruction):
                self.assertIsNone(reference_target(instruction))


class BrokenByTheClean(DocxTestCase):
    """Reported: the target went, and something still names it."""

    def clean_and_verify(self, doc_xml, name, strip_revisions=False):
        engine = self.make_engine()
        source = self.build(doc_xml, name=f"{name}.docx")
        cleaned = self.temp_dir / f"{name}_out.docx"
        self.assertEqual(
            DocxProcessor(engine, strip_revisions=strip_revisions)
            .process(source, cleaned).errors, [])
        return verify_clean(
            source, cleaned, engine=engine, strip_revisions=strip_revisions)

    def assertBroken(self, result, name):
        reported = [str(v) for v in result.structural if "reference broken" in str(v)]
        self.assertTrue(reported, f"not reported: {[str(v) for v in result.structural]}")
        self.assertIn(name, reported[0])
        self.assertFalse(result.passed)

    def test_a_simple_reference_to_a_removed_target(self):
        result = self.clean_and_verify(
            db.document(bookmarked("TargetA", db.run(NOTE)), simple_ref("TargetA")),
            "ref_simple")

        self.assertBroken(result, "TargetA")

    def test_a_complex_reference_with_a_split_instruction(self):
        result = self.clean_and_verify(
            db.document(bookmarked("TargetA", db.run(NOTE)), complex_ref("TargetA")),
            "ref_complex")

        self.assertBroken(result, "TargetA")

    def test_a_quoted_name_containing_spaces(self):
        result = self.clean_and_verify(
            db.document(bookmarked("Target A", db.run(NOTE)),
                        simple_ref(instruction=' REF "Target A" \\h ')),
            "ref_quoted")

        self.assertBroken(result, "Target A")

    def test_an_internal_hyperlink_anchor(self):
        """A hyperlink names its target in ``w:anchor``, not through a field."""
        result = self.clean_and_verify(
            db.document(bookmarked("TargetA", db.run(NOTE)),
                        db.para(db.hyperlink(db.run("jump"), anchor="TargetA"))),
            "ref_link")

        self.assertBroken(result, "TargetA")

    def test_names_match_without_regard_to_case(self):
        """Word matches bookmark names case-insensitively, so this does too."""
        result = self.clean_and_verify(
            db.document(bookmarked("targeta", db.run(NOTE)), simple_ref("TargetA")),
            "ref_case")

        self.assertBroken(result, "TargetA")

    def test_a_target_removed_by_an_accepted_revision_is_still_reported(self):
        """The deletion was requested; the broken reference was not.

        This is the one place a tracked deletion buys no exemption.  Accepting
        the revision is what the run was asked to do, and the consequence for
        the reference is exactly what needs review.
        """
        result = self.clean_and_verify(
            db.document(
                db.table_of(db.deleted_row(bookmarked("TargetA", db.run("Row text.")))),
                simple_ref("TargetA")),
            "ref_revision", strip_revisions=True)

        self.assertBroken(result, "TargetA")


class NotTheCleansFault(DocxTestCase):
    """The false-alarm guards: silence unless this run broke something."""

    def clean_and_verify(self, doc_xml, name):
        engine = self.make_engine()
        source = self.build(doc_xml, name=f"{name}.docx")
        cleaned = self.temp_dir / f"{name}_out.docx"
        self.assertEqual(DocxProcessor(engine).process(source, cleaned).errors, [])
        return verify_clean(source, cleaned, engine=engine)

    def assertSilent(self, result):
        self.assertEqual(
            [str(v) for v in result.structural if "reference broken" in str(v)], [])

    def test_a_half_open_range_survives_the_paragraph(self):
        """The markers are relocated beside the deleted paragraph, so the range holds."""
        result = self.clean_and_verify(db.document(
            db.para(db.bookmark_start("1", "TargetA"), db.run(NOTE)),
            db.para(db.run("Requirement."), db.bookmark_end("1")),
            simple_ref("TargetA"),
        ), "ref_halfopen")

        self.assertSilent(result)

    def test_a_multi_paragraph_target_losing_one_paragraph(self):
        result = self.clean_and_verify(db.document(
            db.para(db.bookmark_start("1", "TargetA"), db.run("Requirement one.")),
            db.text_para(NOTE),
            db.para(db.run("Requirement two."), db.bookmark_end("1")),
            simple_ref("TargetA"),
        ), "ref_multi")

        self.assertSilent(result)

    def test_an_unreferenced_bookmark_may_go(self):
        """Nothing points at it, so nothing broke."""
        result = self.clean_and_verify(
            db.document(bookmarked("Unused", db.run(NOTE)), db.text_para("Anchor.")),
            "ref_unused")

        self.assertSilent(result)
        self.assertTrue(result.passed, result.removed)

    def test_a_reference_already_broken_on_the_way_in(self):
        """Only what this run broke is this run's fault."""
        result = self.clean_and_verify(
            db.document(db.text_para(NOTE), simple_ref("NeverExisted")), "ref_prebroken")

        self.assertSilent(result)

    def test_an_intact_reference_is_not_reported(self):
        result = self.clean_and_verify(db.document(
            bookmarked("TargetA", db.run("Requirement.")), simple_ref("TargetA"),
        ), "ref_intact")

        self.assertSilent(result)
        self.assertTrue(result.passed, result.removed)


if __name__ == "__main__":
    unittest.main()
