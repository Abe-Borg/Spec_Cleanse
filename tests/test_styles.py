"""Style-aware detection resolved against word/styles.xml."""

import unittest

from detection import ContentType

from tests import docx_builder as db
from tests.support import DocxTestCase

STYLES = db.styles(
    db.style_def("CMT", name="Comment"),
    db.style_def("FirmNote", name="Firm Note", based_on="CMT"),
    db.style_def("SpecNote", name="Specifier Note"),
    db.style_def("HiddenNote", name="Hidden Note", hidden=True),
    db.style_def("PRT", name="Part Heading"),
    db.style_def("SCT", name="Section Title"),
    db.style_def("BodyText", name="Body Text"),
)


class EditorialStyleTests(DocxTestCase):

    def clean_with_styles(self, *paragraphs):
        path = self.build(db.document(*paragraphs), {"word/styles.xml": STYLES})
        return self.clean(path)[1]

    def test_style_id_is_matched(self):
        out = self.clean_with_styles(
            db.text_para("Real requirement."),
            db.text_para("Editorial aside.", style="CMT"),
        )

        self.assertEqual(self.paragraph_texts(out), ["Real requirement."])

    def test_display_name_with_a_space_is_matched(self):
        """"Specifier Note" can never equal a style ID, but it is a real name."""
        out = self.clean_with_styles(
            db.text_para("Real requirement."),
            db.text_para("Editorial aside.", style="SpecNote"),
        )

        self.assertEqual(self.paragraph_texts(out), ["Real requirement."])

    def test_a_style_based_on_an_editorial_style_is_matched(self):
        out = self.clean_with_styles(
            db.text_para("Real requirement."),
            db.text_para("Firm's own note.", style="FirmNote"),
        )

        self.assertEqual(self.paragraph_texts(out), ["Real requirement."])

    def test_an_unrelated_style_is_left_alone(self):
        out = self.clean_with_styles(
            db.text_para("Real requirement.", style="BodyText"),
        )

        self.assertEqual(self.paragraph_texts(out), ["Real requirement."])


class HiddenStyleTests(DocxTestCase):
    """MasterSpec hides its notes through the style, not run by run."""

    def test_text_hidden_by_its_style_is_removed(self):
        path = self.build(
            db.document(
                db.text_para("Real requirement."),
                db.text_para("Hidden editorial note.", style="HiddenNote"),
            ),
            {"word/styles.xml": STYLES},
        )
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), ["Real requirement."])

    def test_a_run_that_un_hides_itself_survives(self):
        """<w:vanish w:val="0"/> beats the style it inherits hidden from."""
        path = self.build(
            db.document(
                db.para(db.run("Visible in a hidden style.", vanish=False), style="HiddenNote"),
            ),
            {"word/styles.xml": STYLES},
        )
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), ["Visible in a hidden style."])

    def test_hidden_style_detection_needs_the_styles_part(self):
        """Without styles.xml there is nothing to inherit from — and no false hit."""
        path = self.build(db.document(
            db.text_para("Ordinary text.", style="HiddenNote"),
        ))
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), ["Ordinary text."])


class PreserveStyleTests(DocxTestCase):
    """Headings whose text alone gives nothing away."""

    def test_auto_numbered_part_heading_is_preserved(self):
        """MasterSpec numbers parts automatically: "PART 1 - GENERAL" extracts as "GENERAL"."""
        path = self.build(
            db.document(db.text_para("GENERAL", style="PRT")),
            {"word/styles.xml": STYLES},
        )
        detections = self.detect(path)

        self.assertEqual(
            [d.content_type for d in detections], [ContentType.PRESERVE]
        )

    def test_section_title_style_is_preserved_over_a_removal_pattern(self):
        path = self.build(
            db.document(db.text_para("Retain paragraphs below", style="SCT")),
            {"word/styles.xml": STYLES},
        )
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), ["Retain paragraphs below"])

    def test_body_styles_are_not_preserved(self):
        """Requirement styles stay cleanable — that is where placeholders live."""
        path = self.build(
            db.document(
                db.text_para("Provide [Verify quantity] spare heads.", style="BodyText"),
            ),
            {"word/styles.xml": STYLES},
        )
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), ["Provide spare heads."])


if __name__ == "__main__":
    unittest.main()
