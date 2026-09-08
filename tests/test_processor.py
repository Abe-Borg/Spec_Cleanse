"""Removal safety: what may be deleted, and what must only be emptied."""

import unittest

from detection import ContentType
from docx_xml import W, block_children

from tests import docx_builder as db
from tests.support import DocxTestCase

NOTE = "[Specifier: delete this note before issue]"
COPYRIGHT = "© 2026 ARCOM. All rights reserved."


class ParagraphRemovalTests(DocxTestCase):
    """A plain editorial paragraph is still deleted outright."""

    def test_specifier_note_paragraph_is_deleted(self):
        path = self.build(db.document(
            db.text_para("Provide sprinklers per NFPA 13."),
            db.text_para(NOTE),
        ))
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertNotIn("Specifier", xml)
        self.assertEqual(
            self.paragraph_texts(out), ["Provide sprinklers per NFPA 13."]
        )

    def test_untouched_parts_are_not_rewritten(self):
        """A header with nothing to remove keeps its original bytes."""
        header_xml = db.header(db.text_para("SECTION 21 13 13"))
        path = self.build(
            db.document(db.text_para(NOTE)),
            {"word/header1.xml": header_xml},
        )
        _, out = self.clean(path)

        self.assertEqual(self.part(out, "word/header1.xml"), header_xml)


class ContainerTests(DocxTestCase):
    """Containers Word requires to hold block content are never emptied."""

    def test_sole_paragraph_in_footer_is_kept_as_empty_paragraph(self):
        path = self.build(
            db.document(db.text_para("Body text.")),
            {"word/footer1.xml": db.footer(db.text_para(COPYRIGHT))},
        )
        _, out = self.clean(path)
        footer = self.part(out, "word/footer1.xml")

        self.assertNotIn("ARCOM", footer)
        self.assertEqual(
            self.count_tags(out, "p", "word/footer1.xml"), 1,
            "a footer must keep at least one block-level child",
        )

    def test_sole_paragraph_in_table_cell_is_kept(self):
        path = self.build(db.document(
            db.table(db.text_para(NOTE)),
        ))
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertNotIn("Specifier", xml)
        self.assertEqual(self.count_tags(out, "tc"), 1)
        self.assertEqual(self.count_tags(out, "p"), 1, "the cell must keep a paragraph")

    def test_cell_still_ends_with_a_paragraph(self):
        """Removing the last cell paragraph beside a nested table is unsafe."""
        cell = (
            db.text_para("Keep me.")
            + db.table(db.text_para("Nested."))
            + db.text_para(NOTE)
        )
        path = self.build(db.document(db.table(cell)))
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertNotIn("Specifier", xml)
        self.assertIn("Keep me.", xml)

        # The outer cell's own content still ends with a paragraph, not a table.
        outer_cell = next(self.root(out).iter(f"{W}tc"))
        blocks = block_children(outer_cell)
        self.assertEqual(
            blocks[-1].tag, f"{W}p",
            "a table cell must end with a paragraph or Word calls the file unreadable",
        )

    def test_sole_paragraph_in_footnote_is_kept_with_its_reference_mark(self):
        note = db.para(db.run(inner="<w:footnoteRef/>"), db.run(NOTE))
        path = self.build(
            db.document(db.text_para("Body text.")),
            {"word/footnotes.xml": db.footnotes(note)},
        )
        _, out = self.clean(path)
        footnotes = self.part(out, "word/footnotes.xml")

        self.assertNotIn("Specifier", footnotes)
        self.assertIn("<w:footnoteRef/>", footnotes)
        self.assertEqual(
            self.count_tags(out, "p", "word/footnotes.xml"), 3,
            "each footnote must keep at least one paragraph",
        )

    def test_sole_paragraph_in_text_box_is_kept(self):
        path = self.build(db.document(
            db.para(db.text_box(db.text_para(NOTE))),
        ))
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertNotIn("Specifier", xml)
        text_box_content = next(self.root(out).iter(f"{W}txbxContent"))
        self.assertEqual(
            len(block_children(text_box_content)), 1,
            "a text box must keep at least one block-level child",
        )


class SectionBreakTests(DocxTestCase):
    """A paragraph carrying w:sectPr is a section boundary, not just text."""

    def test_section_break_paragraph_is_emptied_not_deleted(self):
        path = self.build(db.document(
            db.text_para("Body text."),
            db.para(db.run(NOTE), sect=True),
        ))
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertNotIn("Specifier", xml)
        self.assertEqual(self.count_tags(out, "sectPr"), 2, "section break was lost")


class EmbeddedContentTests(DocxTestCase):
    """Pictures, fields, and note anchors survive a paragraph removal."""

    def test_drawing_survives_a_hidden_caption(self):
        path = self.build(db.document(
            db.para(db.run(inner=db.DRAWING), db.run("hidden caption", vanish=True)),
        ))
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertNotIn("hidden caption", xml)
        self.assertEqual(
            self.count_tags(out, "drawing"), 1,
            "the image was deleted along with its hidden caption",
        )

    def test_unbalanced_field_keeps_its_markers(self):
        path = self.build(db.document(
            db.para(db.field_begin(), db.run(NOTE)),
            db.para(db.run("result"), db.field_end()),
        ))
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertNotIn("Specifier", xml)
        self.assertEqual(xml.count('w:fldCharType="begin"'), 1)
        self.assertEqual(xml.count('w:fldCharType="end"'), 1)

    def test_orphaned_bookmark_is_relocated_not_dropped(self):
        path = self.build(db.document(
            db.para(db.bookmark_start("1"), db.run(NOTE)),
            db.para(db.run("Real content."), db.bookmark_end("1")),
        ))
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertNotIn("Specifier", xml)
        self.assertIn('<w:bookmarkStart w:id="1"', xml)
        self.assertIn('<w:bookmarkEnd w:id="1"/>', xml)


class NestingTests(DocxTestCase):
    """Runs are visited once, through the paragraph they belong to."""

    def test_hyperlink_run_is_detected_once(self):
        path = self.build(db.document(
            db.para(
                db.run("See "),
                db.hyperlink(db.run("hidden link text", vanish=True)),
            ),
        ))
        detections = self.detect(path)
        hidden = [d for d in detections if d.content_type == ContentType.HIDDEN_TEXT]

        self.assertEqual(len(hidden), 1, f"expected one detection, got {hidden}")

    def test_text_box_paragraph_is_detected_once(self):
        path = self.build(db.document(
            db.para(db.text_box(db.text_para(NOTE))),
        ))
        detections = self.detect(path)
        notes = [
            d for d in detections
            if d.content_type == ContentType.SPECIFIER_NOTE and NOTE in d.text
        ]

        self.assertEqual(len(notes), 1, f"expected one detection, got {notes}")


class ToggleTests(DocxTestCase):
    """w:val on a toggle property means what it says."""

    def test_explicitly_unhidden_text_is_kept(self):
        """Word writes <w:vanish w:val="0"/> to un-hide inherited hidden text."""
        path = self.build(db.document(
            db.para(db.run("Visible requirement text.", vanish=False)),
        ))
        _, out = self.clean(path)

        self.assertIn("Visible requirement text.", self.part(out))

    def test_hidden_text_is_still_removed(self):
        path = self.build(db.document(
            db.para(db.run("Keep this. "), db.run("Hidden note.", vanish=True)),
        ))
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertIn("Keep this.", xml)
        self.assertNotIn("Hidden note.", xml)

    def test_empty_hidden_run_keeps_its_field(self):
        path = self.build(db.document(
            db.para(
                db.run("Page "),
                db.run(vanish=True, inner='<w:fldChar w:fldCharType="begin"/>'),
                db.run(vanish=True, inner='<w:fldChar w:fldCharType="end"/>'),
            ),
        ))
        _, out = self.clean(path)
        xml = self.part(out)

        self.assertIn('w:fldCharType="begin"', xml)
        self.assertIn('w:fldCharType="end"', xml)


if __name__ == "__main__":
    unittest.main()
