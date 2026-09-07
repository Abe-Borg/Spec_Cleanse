"""The GUI's file-level workers, exercised without opening a window."""

import unittest

from tests import docx_builder as db
from tests.support import DocxTestCase

try:
    import gui
    GUI_IMPORTABLE = True
except ImportError:  # pragma: no cover - environments without Tk
    GUI_IMPORTABLE = False


DOCUMENT = db.document(
    db.text_para("SECTION 21 13 13 - WET-PIPE SPRINKLER SYSTEMS"),
    db.text_para("PART 1 - GENERAL"),
    db.text_para("Retain subparagraph below for wet-pipe systems."),
    db.text_para("Provide two [Verify quantity with Owner] spare sprinklers."),
    db.para(db.run("Coordinate hangers with structural.", italic=True, color="FF0000")),
    db.text_para("END OF SECTION 21 13 13"),
)


@unittest.skipUnless(GUI_IMPORTABLE, "gui.py needs tkinter, which is unavailable")
class WorkerTests(DocxTestCase):

    def setUp(self):
        super().setUp()
        self.lines: list[str] = []

    def log(self, text: str):
        self.lines.append(text)

    @property
    def output(self) -> str:
        return "\n".join(self.lines)

    def test_preview_reports_every_category(self):
        path = self.build(DOCUMENT)

        self.assertTrue(gui._preview_one(path, self.make_engine(), self.log))
        self.assertIn("REMOVALS — editorial_artifact", self.output)
        self.assertIn("INLINE REDACTIONS", self.output)
        self.assertIn("PRESERVED", self.output)
        self.assertIn("formatting-only", self.output)

    def test_clean_reports_a_pass(self):
        path = self.build(DOCUMENT)
        out = self.temp_dir / "out.docx"

        self.assertTrue(gui._clean_one(path, out, self.make_engine(), self.log))
        self.assertIn("PASS", self.output)
        self.assertIn("Paragraphs modified: 1", self.output)
        self.assertTrue(out.exists())

    def test_clean_reports_a_missing_file_without_raising(self):
        out = self.temp_dir / "out.docx"

        self.assertFalse(
            gui._clean_one(self.temp_dir / "nope.docx", out, self.make_engine(), self.log)
        )
        self.assertIn("ERROR", self.output)

    def test_build_engine_reports_a_broken_config(self):
        bad = self.temp_dir / "patterns.yaml"
        bad.write_text("specifier_notes:\n  text_patterns:\n    - '[unclosed'\n", encoding="utf-8")
        original = gui.CONFIG_PATH
        gui.CONFIG_PATH = bad
        try:
            with self.assertRaises(ValueError):
                gui.build_engine()
        finally:
            gui.CONFIG_PATH = original


if __name__ == "__main__":
    unittest.main()
