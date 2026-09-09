"""A removed paragraph that took part in automatic numbering is noticed.

Deliberately its own category — not a structural violation and not an
unexplained removal.  No text-integrity claim is being made: nothing was lost.
What may have changed is the numbers a reader sees, and any reference written
against them.  Saying that as damage would be false; saying nothing would hide a
real consequence.

Nothing is renumbered and no cross-reference is rewritten.  §13.1's reference
check is what asserts a specific reference broke, and it names the bookmark.
"""

import unittest

from processor import DocxProcessor
from verify import verify_clean

from tests import docx_builder as db
from tests.support import DocxTestCase


NOTE = "Note to Specifier: delete this."

#: ``ListItem`` numbers through the style; ``NoNum`` inherits from it and then
#: explicitly switches numbering off, which is an override and not a list.
STYLES = db.styles(
    '<w:style w:type="paragraph" w:styleId="ListItem"><w:name w:val="List Item"/>'
    '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="NoNum"><w:name w:val="No Num"/>'
    '<w:basedOn w:val="ListItem"/>'
    '<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:style>',
)


def numbered(text, num="3"):
    return db.para(
        db.run(text),
        ppr_extra=f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr>')


class NumberingNotices(DocxTestCase):

    def clean_and_verify(self, doc_xml, name):
        engine = self.make_engine()
        source = self.build(doc_xml, {"word/styles.xml": STYLES}, name=f"{name}.docx")
        cleaned = self.temp_dir / f"{name}_out.docx"
        self.assertEqual(DocxProcessor(engine).process(source, cleaned).errors, [])
        return verify_clean(source, cleaned, engine=engine)

    def test_a_removed_numbered_paragraph_is_noticed(self):
        result = self.clean_and_verify(db.document(
            numbered(NOTE), numbered("Requirement one."), db.text_para("Anchor.")), "num_direct")

        self.assertEqual(len(result.numbering), 1, result.numbering)
        self.assertIn("may change", str(result.numbering[0]))
        self.assertIn("3", str(result.numbering[0]))

    def test_the_notice_is_not_damage(self):
        """It makes the run need review without claiming anything was lost."""
        result = self.clean_and_verify(db.document(
            numbered(NOTE), numbered("Requirement one."), db.text_para("Anchor.")), "num_kind")

        self.assertFalse(result.passed)
        self.assertEqual(result.unexpected_removals, [])
        self.assertEqual(result.structural, [])
        self.assertEqual(len(result.expected_removals), 1)

    def test_numbering_inherited_from_a_style_counts(self):
        result = self.clean_and_verify(db.document(
            db.text_para(NOTE, style="ListItem"),
            db.text_para("Requirement.", style="ListItem"),
            db.text_para("Anchor."),
        ), "num_style")

        self.assertEqual(len(result.numbering), 1, result.numbering)
        self.assertIn("7", str(result.numbering[0]))

    def test_a_style_chain_that_switches_numbering_off_is_not_participation(self):
        """``w:numId`` ``"0"`` is an override, not a list called zero."""
        result = self.clean_and_verify(db.document(
            db.text_para(NOTE, style="NoNum"),
            db.text_para("Requirement.", style="ListItem"),
            db.text_para("Anchor."),
        ), "num_off")

        self.assertEqual(result.numbering, [])

    def test_a_list_with_no_survivors_renumbers_nothing(self):
        """A notice here would be noise dressed as precision."""
        result = self.clean_and_verify(
            db.document(numbered(NOTE), db.text_para("Anchor.")), "num_alone")

        self.assertEqual(result.numbering, [])
        self.assertTrue(result.passed, result.removed)

    def test_an_unnumbered_removal_is_not_noticed(self):
        result = self.clean_and_verify(db.document(
            db.text_para(NOTE), numbered("Requirement one."), db.text_para("Anchor."),
        ), "num_plain")

        self.assertEqual(result.numbering, [])
        self.assertTrue(result.passed, result.removed)

    def test_a_clean_that_removes_nothing_is_silent(self):
        result = self.clean_and_verify(
            db.document(numbered("Requirement one."), numbered("Requirement two.")),
            "num_none")

        self.assertEqual(result.numbering, [])
        self.assertTrue(result.passed, result.removed)


if __name__ == "__main__":
    unittest.main()
