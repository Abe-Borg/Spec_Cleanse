"""The optional pass that accepts tracked changes and drops comments."""

import unittest
import zipfile

from processor import DocxProcessor

from tests import docx_builder as db
from tests.support import DocxTestCase

DOCUMENT = db.document(
    db.para(
        db.run("Sprinkler spacing shall not exceed "),
        db.inserted(db.run("12 feet")),
        db.deleted("15 feet"),
        db.run(" on centre."),
        db.comment_anchor("1"),
    ),
)
PARTS = {"word/comments.xml": db.comments(db.text_para("Check this with the AHJ."))}


class DefaultTests(DocxTestCase):
    """Off by default: the document's revision history is left alone."""

    def test_tracked_changes_survive_an_ordinary_clean(self):
        path = self.build(DOCUMENT, PARTS)
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertIn("<w:ins", xml)
        self.assertIn("15 feet", xml)
        self.assertIn("<w:commentRangeStart", xml)

    def test_comments_part_survives(self):
        path = self.build(DOCUMENT, PARTS)
        _, out = self.clean(path)

        with zipfile.ZipFile(out) as zf:
            self.assertIn("word/comments.xml", zf.namelist())


class StripRevisionsTests(DocxTestCase):
    """On: insertions are kept, deletions and comments go."""

    def clean_stripped(self, path):
        output_path = self.temp_dir / "stripped.docx"
        processor = DocxProcessor(self.make_engine(), strip_revisions=True)
        result = processor.process(path, output_path)
        self.assertEqual(result.errors, [])
        return output_path

    def test_insertions_are_kept_and_deletions_dropped(self):
        out = self.clean_stripped(self.build(DOCUMENT, PARTS))
        xml = self.part(out)

        self.assertNotIn("<w:ins", xml)
        self.assertNotIn("<w:del", xml)
        self.assertNotIn("15 feet", xml)
        self.assertEqual(
            self.paragraph_texts(out),
            ["Sprinkler spacing shall not exceed 12 feet on centre."],
        )

    def test_comment_anchors_and_part_are_removed(self):
        out = self.clean_stripped(self.build(DOCUMENT, PARTS))
        xml = self.part(out)

        self.assertNotIn("commentRangeStart", xml)
        self.assertNotIn("commentReference", xml)
        with zipfile.ZipFile(out) as zf:
            names = zf.namelist()
        self.assertNotIn("word/comments.xml", names)

    def test_package_bookkeeping_is_updated(self):
        """A part listed in the rels or content types after deletion breaks the file."""
        out = self.clean_stripped(self.build(DOCUMENT, PARTS))

        self.assertNotIn("comments.xml", self.part(out, "word/_rels/document.xml.rels"))
        self.assertNotIn("comments.xml", self.part(out, "[Content_Types].xml"))

    def test_a_deleted_table_row_is_removed_with_its_text(self):
        """A deleted row records the deletion in w:trPr; its text stays plain w:t."""
        path = self.build(db.document(db.table_of(
            db.deleted_row(db.text_para("Row the editor deleted.")),
            db.row(db.text_para("Row the editor kept.")),
        )))
        out = self.clean_stripped(path)

        self.assertEqual(self.paragraph_texts(out), ["Row the editor kept."])

    def test_a_deleted_table_cell_is_removed(self):
        path = self.build(db.document(db.table_of(
            '<w:tr>'
            '<w:tc><w:tcPr><w:cellDel w:id="96" w:author="E" w:date="2026-01-01T00:00:00Z"/>'
            '</w:tcPr>' + db.text_para("Deleted cell.") + '</w:tc>'
            '<w:tc><w:tcPr/>' + db.text_para("Kept cell.") + '</w:tc>'
            '</w:tr>'
        )))
        out = self.clean_stripped(path)

        self.assertEqual(self.paragraph_texts(out), ["Kept cell."])

    def test_the_comment_parts_sidecar_rels_go_too(self):
        """A comment holding an image has its own .rels; orphaning it is invalid OPC."""
        path = self.build(DOCUMENT, dict(
            PARTS, **{"word/_rels/comments.xml.rels": db.COMMENTS_RELS}
        ))
        out = self.clean_stripped(path)

        with zipfile.ZipFile(out) as zf:
            names = zf.namelist()
        self.assertNotIn("word/_rels/comments.xml.rels", names)
        self.assertIn("word/_rels/document.xml.rels", names)

    def test_verification_accounts_for_the_accepted_deletion(self):
        """The report must judge the run it was asked for, not cry wolf about it."""
        from verify import verify_clean

        path = self.build(db.document(db.table_of(
            db.deleted_row(db.text_para("Row the editor deleted.")),
            db.row(db.text_para("Row the editor kept.")),
        )))
        engine = self.make_engine()
        out = self.temp_dir / "stripped.docx"
        DocxProcessor(engine, strip_revisions=True).process(path, out)

        accepted = verify_clean(path, out, engine=engine, strip_revisions=True)
        self.assertTrue(accepted.passed, accepted.removed)
        self.assertEqual(accepted.removed[0].category, "tracked_deletion")

        # Judged as an ordinary clean, the same loss is unexplained.
        plain = verify_clean(path, out, engine=engine)
        self.assertEqual(len(plain.unexpected_removals), 1)

    def test_a_document_without_revisions_is_unharmed(self):
        path = self.build(db.document(db.text_para("Plain requirement.")))
        out = self.clean_stripped(path)

        self.assertEqual(self.paragraph_texts(out), ["Plain requirement."])


if __name__ == "__main__":
    unittest.main()
