"""A table accepted revisions emptied is removed, and an empty one is visible.

Deleting a table's last row already worked; the ``w:tbl`` around it stayed,
holding nothing.  Word does not accept a table with no rows, and neither the
structural lint nor verification could see one — so the package looked clean and
the file did not open.  That is the live demonstration that **a clean lint is not
evidence Word will accept a package**: the lint had no rule for this shape, so it
reported nothing, which is not the same as reporting that nothing is wrong.

A table removed this way is an explained structural change, not an unexplained
loss: the run was asked to accept the revision that emptied it.
"""

import unittest

from processor import DocxProcessor
from verify import lint_structure, verify_clean

from tests import docx_builder as db
from tests.support import DocxTestCase


class EmptiedTables(DocxTestCase):

    def accept(self, doc_xml, name, parts=None, strip_revisions=True):
        engine = self.make_engine()
        source = self.build(doc_xml, parts, name=f"{name}.docx")
        cleaned = self.temp_dir / f"{name}_out.docx"
        self.assertEqual(
            DocxProcessor(engine, strip_revisions=strip_revisions)
            .process(source, cleaned).errors, [])
        result = verify_clean(
            source, cleaned, engine=engine, strip_revisions=strip_revisions)
        return cleaned, result

    def assertSound(self, cleaned, result):
        """No lint, no unexplained change: the shape Word will actually open."""
        self.assertEqual(lint_structure(cleaned), [])
        self.assertTrue(result.passed, f"{result.removed} {result.structural}")

    def test_deleting_the_last_row_removes_the_table(self):
        cleaned, result = self.accept(db.document(
            db.text_para("Before."),
            db.table_of(db.deleted_row(db.text_para("Only row."))),
            db.text_para("After."),
        ), "lastrow")

        self.assertEqual(self.count_tags(cleaned, "tbl"), 0)
        self.assertEqual(self.paragraph_texts(cleaned), ["Before.", "After."])
        self.assertSound(cleaned, result)

    def test_a_table_with_a_surviving_row_is_kept(self):
        cleaned, result = self.accept(db.document(
            db.text_para("Before."),
            db.table_of(db.deleted_row(db.text_para("Gone.")),
                        db.row(db.text_para("Kept."))),
            db.text_para("After."),
        ), "onerowleft")

        self.assertEqual(self.count_tags(cleaned, "tbl"), 1)
        self.assertEqual(self.count_tags(cleaned, "tr"), 1)
        self.assertSound(cleaned, result)

    def test_deleting_every_cell_of_the_only_row_removes_the_table(self):
        """The row goes for want of cells, and then the table for want of rows."""
        cleaned, result = self.accept(db.document(
            db.text_para("Before."),
            db.table_of('<w:tr><w:tc><w:tcPr>'
                        '<w:cellDel w:id="9" w:author="Editor"/></w:tcPr>'
                        + db.text_para("Cell.") + '</w:tc></w:tr>'),
            db.text_para("After."),
        ), "allcells")

        self.assertEqual(self.count_tags(cleaned, "tbl"), 0)
        self.assertSound(cleaned, result)

    def test_an_inner_table_is_emptied_without_disturbing_the_outer(self):
        cleaned, result = self.accept(db.document(
            db.text_para("Before."),
            db.table_of(db.row(db.table_of(db.deleted_row(db.text_para("Inner.")))
                               + db.text_para("Cell tail."))),
            db.text_para("After."),
        ), "nested")

        self.assertEqual(self.count_tags(cleaned, "tbl"), 1)
        self.assertSound(cleaned, result)

    def test_an_emptied_table_that_was_a_cell_s_last_block(self):
        """A cell must still *end* with a paragraph once the table goes."""
        cleaned, result = self.accept(db.document(
            db.text_para("Before."),
            db.table_of(db.row(db.text_para("Cell head.")
                               + db.table_of(db.deleted_row(db.text_para("Inner."))))),
            db.text_para("After."),
        ), "celltail")

        self.assertEqual(self.count_tags(cleaned, "tbl"), 1)
        self.assertSound(cleaned, result)

    def test_both_levels_emptied(self):
        cleaned, result = self.accept(db.document(
            db.text_para("Before."),
            db.table_of(db.deleted_row(
                db.table_of(db.deleted_row(db.text_para("Inner.")))
                + db.text_para("Tail."))),
            db.text_para("After."),
        ), "bothlevels")

        self.assertEqual(self.count_tags(cleaned, "tbl"), 0)
        self.assertSound(cleaned, result)

    def test_a_sole_table_in_the_body_leaves_a_paragraph(self):
        """Removing the only block would empty a container Word needs filled."""
        cleaned, result = self.accept(
            db.document(db.table_of(db.deleted_row(db.text_para("Only row.")))),
            "solebody")

        self.assertEqual(self.count_tags(cleaned, "tbl"), 0)
        self.assertEqual(self.count_tags(cleaned, "p"), 1)
        self.assertSound(cleaned, result)

    def test_a_sole_table_in_a_header_leaves_a_paragraph(self):
        cleaned, result = self.accept(
            db.document(db.text_para("Body.")),
            "soleheader",
            parts={"word/header1.xml":
                   db.header(db.table_of(db.deleted_row(db.text_para("Only row."))))})

        self.assertEqual(self.count_tags(cleaned, "tbl", "word/header1.xml"), 0)
        self.assertEqual(self.count_tags(cleaned, "p", "word/header1.xml"), 1)
        self.assertSound(cleaned, result)

    def test_nothing_happens_when_revisions_are_not_accepted(self):
        """The deletion is only authorized by the option that asks for it."""
        cleaned, result = self.accept(db.document(
            db.text_para("Before."),
            db.table_of(db.deleted_row(db.text_para("Only row."))),
            db.text_para("After."),
        ), "norevisions", strip_revisions=False)

        self.assertEqual(self.count_tags(cleaned, "tbl"), 1)
        self.assertEqual(self.count_tags(cleaned, "tr"), 1)
        self.assertSound(cleaned, result)


class OnlyWhatThisRunEmptied(DocxTestCase):
    """A table that arrived rowless is the document's own problem.

    Removing it would rewrite a file that had no revisions to accept, merely
    because the option was on — a change for a reason unrelated to what the run
    was asked to do.  It is still linted, so it is visible without being
    silently repaired.
    """

    ROWLESS = '<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr></w:tbl>'

    def accept(self, doc_xml, name):
        engine = self.make_engine()
        source = self.build(doc_xml, name=f"{name}.docx")
        cleaned = self.temp_dir / f"{name}_out.docx"
        self.assertEqual(
            DocxProcessor(engine, strip_revisions=True).process(source, cleaned).errors, [])
        return cleaned

    def test_a_pre_existing_rowless_table_is_left_alone(self):
        cleaned = self.accept(db.document(
            db.text_para("Body."), self.ROWLESS, db.text_para("After.")), "prerowless")

        self.assertEqual(self.count_tags(cleaned, "tbl"), 1)

    def test_it_is_still_left_alone_beside_a_table_this_run_empties(self):
        """Only the one whose row actually went is removed."""
        cleaned = self.accept(db.document(
            db.text_para("Body."), self.ROWLESS,
            db.table_of(db.deleted_row(db.text_para("Only row."))),
            db.text_para("After.")), "prerowless_mixed")

        self.assertEqual(self.count_tags(cleaned, "tbl"), 1)
        self.assertTrue(any("no rows" in issue for issue in lint_structure(cleaned)),
                        "the surviving one is the pre-existing rowless table")

    def test_a_pre_existing_rowless_table_is_not_blamed_on_the_clean(self):
        """Present on both sides, so the comparison stays silent about it."""
        engine = self.make_engine()
        doc = db.document(db.text_para("Body."), self.ROWLESS, db.text_para("After."))
        source = self.build(doc, name="prerowless_v.docx")
        cleaned = self.accept(doc, "prerowless_v2")

        self.assertEqual(verify_clean(source, cleaned, engine=engine).structural, [])


class EmptyStructureIsVisible(DocxTestCase):
    """The lint had no rule for these shapes, so it reported nothing."""

    def test_a_table_with_no_rows_is_linted(self):
        path = self.build(db.document(
            db.text_para("Body."), '<w:tbl><w:tblPr/></w:tbl>'), name="rowless.docx")

        self.assertTrue(any("no rows" in issue for issue in lint_structure(path)),
                        lint_structure(path))

    def test_a_row_with_no_cells_is_linted(self):
        path = self.build(db.document(
            db.text_para("Body."), '<w:tbl><w:tblPr/><w:tr><w:trPr/></w:tr></w:tbl>'),
            name="cellless.docx")

        self.assertTrue(any("no cells" in issue for issue in lint_structure(path)),
                        lint_structure(path))

    def test_a_sound_table_lints_clean(self):
        path = self.build(db.document(
            db.text_para("Body."), db.table_of(db.row(db.text_para("Cell.")))),
            name="soundtable.docx")

        self.assertEqual(lint_structure(path), [])


if __name__ == "__main__":
    unittest.main()
