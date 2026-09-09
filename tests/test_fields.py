"""Field carriers survive editorial cleaning, and their loss is visible.

Word records the same field two ways.  A *simple* field is one ``w:fldSimple``
carrying its instruction in an attribute; a *complex* one is a run sequence
delimited by ``w:fldChar`` with the instruction in ``w:instrText`` between them.
Only the second was ever protected, so a paragraph whose one field was simple
had no carrier protection at all — the cached result reads as ordinary words, so
an editorial pattern matching the paragraph deleted the live cross-reference with
it, and nothing reported the loss.

Two things follow, and they are separate:

* the cleaner must keep the carrier, emptying the paragraph in place rather than
  deleting it, and
* verification must be able to *see* a carrier go even when the text is
  unchanged, because a stripped field leaves the same characters behind.

Keeping a wrapper is not a promise about its value: Word recalculates fields on
refresh, so a preserved-but-emptied field result may come back.  What is promised
is that the instruction and its wrapper are still there to recalculate from.
"""

import unittest

from docx_xml import field_instructions, has_embedded_content, iter_paragraphs
from verify import verify_clean

from tests import docx_builder as db
from tests.support import DocxTestCase


SIMPLE = (
    '<w:p><w:r><w:t>Note to Specifier: see </w:t></w:r>'
    '<w:fldSimple w:instr=" REF Target "><w:r><w:t>Section 21 13 13</w:t></w:r>'
    '</w:fldSimple><w:r><w:t> before issue.</w:t></w:r></w:p>'
)
#: The same field, written the long way, with the instruction split across two
#: ``w:instrText`` nodes as Word does.
COMPLEX = (
    '<w:p><w:r><w:t>Note to Specifier: see </w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> REF </w:instrText></w:r>'
    '<w:r><w:instrText xml:space="preserve">Target </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    '<w:r><w:t>Section 21 13 13</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    '<w:r><w:t> before issue.</w:t></w:r></w:p>'
)


class FieldIdentity(DocxTestCase):
    """The two ways of writing a field describe the same field."""

    def test_both_forms_yield_the_same_instruction(self):
        simple = field_instructions(self.root(self.build(db.document(SIMPLE), name="s.docx")))
        complex_ = field_instructions(self.root(self.build(db.document(COMPLEX), name="c.docx")))

        self.assertEqual(dict(simple), {"REF Target": 1})
        self.assertEqual(dict(complex_), dict(simple),
                         "a split instruction must normalise to the same string")

    def test_both_forms_protect_their_paragraph(self):
        for label, xml in (("simple", SIMPLE), ("complex", COMPLEX)):
            with self.subTest(label):
                root = self.root(self.build(db.document(xml), name=f"p{label}.docx"))
                self.assertTrue(has_embedded_content(next(iter_paragraphs(root))))

    def test_repeated_fields_are_counted_not_collapsed(self):
        """Losing one of two identical fields is still a loss."""
        twice = self.root(self.build(db.document(SIMPLE, SIMPLE), name="twice.docx"))

        self.assertEqual(dict(field_instructions(twice)), {"REF Target": 2})


def _field(instruction):
    """The opening half of a complex field: begin plus its instruction."""
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve">{instruction}</w:instrText></w:r>')


_SEPARATE = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
_END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

#: A field nested inside another field's result — an IF whose condition is a
#: PAGE field.  Two carriers, not one.
NESTED = ('<w:p>' + _field(" IF ") + _field(" PAGE ") + _SEPARATE
          + '<w:r><w:t>7</w:t></w:r>' + _END + _SEPARATE
          + '<w:r><w:t>x</w:t></w:r>' + _END + '</w:p>')

#: The same paragraph with the *inner* field's begin and end stripped: its
#: instruction text and cached result survive, so the characters are unchanged.
NESTED_DAMAGED = ('<w:p>' + _field(" IF ")
                  + '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                  + _SEPARATE + '<w:r><w:t>7</w:t></w:r>'
                  + _SEPARATE + '<w:r><w:t>x</w:t></w:r>' + _END + '</w:p>')


class NestedFields(DocxTestCase):
    """A field inside another field's result is its own carrier.

    Accumulating instructions into one buffer and emitting when the nesting
    closed merged the two, so an output that lost the inner field's begin and
    end — keeping its instruction text and cached result — produced an
    identical inventory, and the loss was invisible.
    """

    def test_nested_fields_are_counted_separately(self):
        found = field_instructions(self.root(self.build(db.document(NESTED), name="n.docx")))

        self.assertEqual(dict(found), {"IF": 1, "PAGE": 1})

    def test_losing_the_inner_field_changes_the_inventory(self):
        intact = field_instructions(self.root(self.build(db.document(NESTED), name="ni.docx")))
        damaged = field_instructions(
            self.root(self.build(db.document(NESTED_DAMAGED), name="nd.docx")))

        self.assertNotEqual(dict(intact), dict(damaged),
                            "the loss of a nested carrier must be visible")

    def test_a_stripped_nested_field_is_reported(self):
        source = self.build(db.document(NESTED, db.text_para("Anchor.")), name="nx_in.docx")
        damaged = self.build(db.document(NESTED_DAMAGED, db.text_para("Anchor.")),
                             name="nx_out.docx")

        result = verify_clean(source, damaged, engine=self.make_engine())

        self.assertTrue(any("field lost" in str(v) for v in result.structural),
                        f"not reported: {[str(v) for v in result.structural]}")
        self.assertFalse(result.passed)

    def test_an_unchanged_nested_field_document_stays_silent(self):
        """The false-alarm guard for the same code path."""
        doc = db.document(NESTED, db.text_para("Anchor."))
        source = self.build(doc, name="ns_in.docx")
        same = self.build(doc, name="ns_out.docx")

        self.assertEqual(verify_clean(source, same, engine=self.make_engine()).structural, [])


class AcceptedMoves(DocxTestCase):
    """The source half of a tracked move goes when revisions are accepted.

    ``w:moveFrom`` is the one revision whose content reaches the comparison as
    ordinary ``w:t`` — a deleted run hides its text in ``w:delText``, which no
    extractor reads.  So accepting a move was reported as an unexplained
    removal, and any field inside it as a lost carrier.
    """

    MOVED = ('<w:p><w:moveFrom w:id="1" w:author="Editor">'
             '<w:r><w:t>Moved requirement text.</w:t></w:r></w:moveFrom></w:p>')
    MOVED_FIELD = ('<w:p><w:moveFrom w:id="1" w:author="Editor">'
                   '<w:fldSimple w:instr=" REF Moved "><w:r><w:t>x</w:t></w:r>'
                   '</w:fldSimple></w:moveFrom></w:p>')

    def _accept(self, xml, name):
        from processor import DocxProcessor
        engine = self.make_engine()
        source = self.build(db.document(db.text_para("Body."), xml), name=f"{name}.docx")
        cleaned = self.temp_dir / f"{name}_out.docx"
        self.assertEqual(
            DocxProcessor(engine, strip_revisions=True).process(source, cleaned).errors, [])
        return source, verify_clean(source, cleaned, engine=engine, strip_revisions=True)

    def test_accepting_a_move_is_not_an_unexplained_removal(self):
        _, result = self._accept(self.MOVED, "move_plain")

        self.assertTrue(result.passed, result.removed)
        self.assertEqual([r.category for r in result.removed], ["tracked_deletion"])

    def test_a_field_inside_an_accepted_move_is_not_a_lost_carrier(self):
        _, result = self._accept(self.MOVED_FIELD, "move_field")

        self.assertEqual(result.structural, [])
        self.assertTrue(result.passed, result.removed)

    def test_the_same_loss_is_damage_when_revisions_are_kept(self):
        """Without the option that authorizes it, the authority does not exist."""
        source = self.build(db.document(db.text_para("Body."), self.MOVED),
                            name="movekeep_in.docx")
        damaged = self.build(db.document(db.text_para("Body.")), name="movekeep_out.docx")

        result = verify_clean(source, damaged, engine=self.make_engine())

        self.assertFalse(result.passed)
        self.assertEqual(len(result.unexpected_removals), 1, result.removed)


class CarrierSurvivesCleaning(DocxTestCase):
    """An editorial paragraph is emptied in place when it carries a field."""

    def _clean(self, doc_xml, name):
        engine = self.make_engine()
        source = self.build(doc_xml, name=f"{name}.docx")
        _, cleaned = self.clean(source, engine)
        return source, cleaned, verify_clean(source, cleaned, engine=engine)

    def test_a_simple_field_survives_removal_of_its_editorial_paragraph(self):
        """X05 — the whole paragraph matches a removal rule; the field stays."""
        source, cleaned, result = self._clean(
            db.document(SIMPLE, db.text_para("Anchor.")), "x05")

        self.assertEqual(dict(field_instructions(self.root(cleaned))), {"REF Target": 1})
        self.assertEqual(self.count_tags(cleaned, "fldSimple"), 1)
        self.assertNotIn("Note to Specifier", db.read_part(cleaned),
                         "the editorial text should still be gone")
        self.assertTrue(result.passed, result.removed)

    def test_a_simple_field_in_a_text_box_survives(self):
        """X07 — the nested paragraph is visited in its own right, and protected."""
        source, cleaned, result = self._clean(db.document(
            db.para(db.run("Outer text. "),
                    db.text_box('<w:p><w:r><w:t>Note to Specifier: </w:t></w:r>'
                                '<w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r>'
                                '</w:fldSimple></w:p>')),
            db.text_para("Anchor."),
        ), "x07")

        self.assertEqual(dict(field_instructions(self.root(cleaned))), {"PAGE": 1})
        self.assertNotIn("Note to Specifier", db.read_part(cleaned))
        self.assertTrue(result.passed, result.removed)

    def test_a_complex_field_still_survives(self):
        """The protection that already worked must keep working."""
        source, cleaned, result = self._clean(
            db.document(COMPLEX, db.text_para("Anchor.")), "cplx")

        self.assertEqual(dict(field_instructions(self.root(cleaned))), {"REF Target": 1})
        self.assertTrue(result.passed, result.removed)


class CarrierLossIsVisible(DocxTestCase):
    """Verification sees a field go even when the text does not change."""

    def test_a_stripped_simple_field_is_a_structural_violation(self):
        """X06 — same characters, no live reference: nothing else would notice."""
        source = self.build(db.document(
            db.text_para("Provide valves per the referenced section."),
            '<w:p><w:r><w:t>See </w:t></w:r>'
            '<w:fldSimple w:instr=" REF Target "><w:r><w:t>Section 21 13 13</w:t></w:r>'
            '</w:fldSimple><w:r><w:t> for details.</w:t></w:r></w:p>',
        ), name="x06_in.docx")
        damaged = self.build(db.document(
            db.text_para("Provide valves per the referenced section."),
            db.para(db.run("See "), db.run("Section 21 13 13"), db.run(" for details.")),
        ), name="x06_out.docx")

        result = verify_clean(source, damaged, engine=self.make_engine())

        self.assertTrue(
            any("field lost" in str(v) for v in result.structural),
            f"stripped field not reported: {[str(v) for v in result.structural]}")
        self.assertFalse(result.passed)

    def test_a_field_lost_from_a_header_is_not_answered_by_the_body(self):
        """Per part, so an identical field elsewhere does not cover the loss."""
        body = db.document(db.text_para("Body."), 
                           '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r>'
                           '</w:fldSimple></w:p>')
        source = self.build(body, {"word/header1.xml": db.header(
            '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>')},
            name="hdrfield_in.docx")
        damaged = self.build(body, {"word/header1.xml": db.header(db.text_para("1"))},
                             name="hdrfield_out.docx")

        result = verify_clean(source, damaged, engine=self.make_engine())

        self.assertTrue(
            any("header1.xml" in str(v) and "field lost" in str(v) for v in result.structural),
            f"header field loss not reported: {[str(v) for v in result.structural]}")

    def test_an_unchanged_document_reports_no_field_loss(self):
        """The false-alarm guard: keeping every field must stay silent."""
        doc = db.document(SIMPLE, COMPLEX, db.text_para("Anchor."))
        source = self.build(doc, name="same_in.docx")
        same = self.build(doc, name="same_out.docx")

        result = verify_clean(source, same, engine=self.make_engine())

        self.assertEqual(result.structural, [])

    def test_a_field_an_accepted_revision_removes_is_not_a_violation(self):
        """The source revision is the evidence that explains its absence."""
        source = self.build(db.document(
            db.text_para("Body."),
            db.table_of(db.deleted_row(
                '<w:p><w:fldSimple w:instr=" REF Gone "><w:r><w:t>x</w:t></w:r>'
                '</w:fldSimple></w:p>')),
        ), name="delfield_in.docx")
        engine = self.make_engine()
        from processor import DocxProcessor
        cleaned = self.temp_dir / "delfield_out.docx"
        self.assertEqual(
            DocxProcessor(engine, strip_revisions=True).process(source, cleaned).errors, [])

        result = verify_clean(source, cleaned, engine=engine, strip_revisions=True)

        self.assertFalse(
            any("field lost" in str(v) for v in result.structural),
            f"an accepted deletion should explain the field: "
            f"{[str(v) for v in result.structural]}")

    def test_the_same_deletion_is_a_violation_when_revisions_are_kept(self):
        """Without the option that authorizes it, the authority does not exist."""
        row = db.table_of(db.deleted_row(
            '<w:p><w:fldSimple w:instr=" REF Gone "><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>'))
        source = self.build(db.document(db.text_para("Body."), row), name="keepfield_in.docx")
        damaged = self.build(db.document(db.text_para("Body."),
                                         db.table_of(db.row(db.text_para("x")))),
                             name="keepfield_out.docx")

        result = verify_clean(source, damaged, engine=self.make_engine())

        self.assertTrue(
            any("field lost" in str(v) for v in result.structural),
            f"field loss not reported: {[str(v) for v in result.structural]}")


if __name__ == "__main__":
    unittest.main()
