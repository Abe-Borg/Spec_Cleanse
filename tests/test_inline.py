"""Inline placeholders are cut out; the requirement around them survives."""

import unittest

from detection import ContentType

from tests import docx_builder as db
from tests.support import DocxTestCase


class InlineRedactionTests(DocxTestCase):

    def test_placeholder_is_cut_and_the_sentence_kept(self):
        path = self.build(db.document(
            db.text_para("Provide two [Verify quantity with Owner] spare sprinklers."),
        ))
        _, out = self.clean(path)

        self.assertEqual(
            self.paragraph_texts(out),
            ["Provide two spare sprinklers."],
        )

    def test_placeholder_split_across_runs(self):
        path = self.build(db.document(
            db.para(
                db.run("Manufacturer: "),
                db.run("<Insert manufacturer's name>"),
                db.run(" or approved equal."),
            ),
        ))
        _, out = self.clean(path)

        self.assertEqual(
            self.paragraph_texts(out),
            ["Manufacturer: or approved equal."],
        )

    def test_paragraph_of_nothing_but_placeholders_is_removed(self):
        path = self.build(db.document(
            db.text_para("Keep this requirement."),
            db.text_para("[Insert product name]"),
        ))
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), ["Keep this requirement."])

    def test_preserved_paragraph_is_not_redacted(self):
        path = self.build(db.document(
            db.text_para("SECTION 23 05 00 - [Verify with Owner]"),
        ))
        _, out = self.clean(path)

        self.assertEqual(
            self.paragraph_texts(out),
            ["SECTION 23 05 00 - [Verify with Owner]"],
        )

    def test_inline_detection_does_not_remove_the_paragraph(self):
        path = self.build(db.document(
            db.text_para("Pipe shall be Schedule 40 [___] black steel."),
        ))
        detections = self.detect(path)
        inline = [d for d in detections if d.content_type == ContentType.INLINE_PLACEHOLDER]

        self.assertEqual(len(inline), 1)
        self.assertEqual(len(inline[0].spans), 1)


class PatternPrecisionTests(DocxTestCase):
    """Patterns that used to fire on real specification prose."""

    KEPT = [
        "Retain records of all tests required in Paragraph 1.6.",
        "Submit one copy of each report described in Paragraph 1.5.",
        "Select one-piece molded fittings for changes in direction.",
        "Owner proprietary information shall remain confidential.",
        "Sprinkler system shall comply with NFPA 13, latest edition.",
    ]

    REMOVED = [
        "Retain subparagraph below for wet-pipe systems.",
        "Copy paragraphs above for each additional riser.",
        "Select one of the two options below.",
        "Retain or delete the following as applicable.",
    ]

    def test_real_spec_prose_is_kept(self):
        path = self.build(db.document(*[db.text_para(t) for t in self.KEPT]))
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), self.KEPT)

    def test_editorial_instructions_are_removed(self):
        path = self.build(db.document(
            db.text_para("Real requirement."),
            *[db.text_para(t) for t in self.REMOVED],
        ))
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), ["Real requirement."])

    def test_masterformat_section_heading_is_preserved(self):
        path = self.build(db.document(
            db.text_para("SECTION 23 05 00 - COMMON WORK RESULTS FOR HVAC"),
            db.text_para("SECTION 238126"),
            db.text_para("PART 2 - PRODUCTS"),
        ))
        detections = self.detect(path)
        preserved = [d for d in detections if d.content_type == ContentType.PRESERVE]

        self.assertEqual(len(preserved), 3)


class FormattingOnlySwitchTests(DocxTestCase):
    """specifier_notes.formatting_only_removal decides whether looks are enough."""

    DOC = None

    def _document(self):
        return db.document(
            db.text_para("Real requirement text."),
            db.para(db.run("Coordinate hangers with structural.", italic=True, color="FF0000")),
        )

    def test_on_by_default_removes_italic_editorial_colour(self):
        path = self.build(self._document())
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), ["Real requirement text."])

    def test_off_keeps_it(self):
        path = self.build(self._document())
        _, out = self.clean(
            path, specifier_notes={"formatting_only_removal": False}
        )

        self.assertEqual(
            self.paragraph_texts(out),
            ["Real requirement text.", "Coordinate hangers with structural."],
        )

    def test_off_still_removes_pattern_matches(self):
        path = self.build(db.document(
            db.para(db.run("[Specifier: delete before issue]", italic=True, color="FF0000")),
            db.text_para("Real requirement text."),
        ))
        _, out = self.clean(
            path, specifier_notes={"formatting_only_removal": False}
        )

        self.assertEqual(self.paragraph_texts(out), ["Real requirement text."])

    def test_removal_is_labelled_formatting_only(self):
        path = self.build(self._document())
        detections = self.detect(path)
        flagged = [d for d in detections if d.formatting_only]

        self.assertEqual(len(flagged), 1)
        self.assertIn("formatting-only", flagged[0].reason)


if __name__ == "__main__":
    unittest.main()
