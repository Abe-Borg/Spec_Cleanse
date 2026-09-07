"""The optional pass that accepts tracked changes and drops comments."""

import unittest
import zipfile

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
        from processor import DocxProcessor

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

    def test_a_document_without_revisions_is_unharmed(self):
        path = self.build(db.document(db.text_para("Plain requirement.")))
        out = self.clean_stripped(path)

        self.assertEqual(self.paragraph_texts(out), ["Plain requirement."])


if __name__ == "__main__":
    unittest.main()
