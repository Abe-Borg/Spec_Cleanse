"""Shared helpers for the SpecCleanse test suite."""

import shutil
import tempfile
import unittest
from pathlib import Path

from lxml import etree

from detection import DetectionEngine
from docx_xml import W, iter_paragraphs, load_config, paragraph_text
from processor import DocxProcessor

from tests import docx_builder as db

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "patterns.yaml"


class DocxTestCase(unittest.TestCase):
    """Base case that builds documents in a temp dir and cleans them."""

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_test_"))
        self.addCleanup(shutil.rmtree, self.temp_dir, True)
        self.config = load_config(CONFIG_PATH)

    def make_engine(self, **overrides) -> DetectionEngine:
        """Build an engine from patterns.yaml, with optional config overrides.

        ``overrides`` are applied as ``{"section": {"key": value}}`` updates.
        """
        config = self._copy_config()
        for section, values in overrides.items():
            config.setdefault(section, {}).update(values)
        return DetectionEngine(config)

    def _copy_config(self) -> dict:
        return {
            key: (dict(value) if isinstance(value, dict) else value)
            for key, value in self.config.items()
        }

    def build(self, document_xml: str, extra_parts: dict | None = None, name: str = "in.docx") -> Path:
        return db.build_docx(self.temp_dir / name, document_xml, extra_parts)

    def clean(self, input_path: Path, engine: DetectionEngine | None = None, **engine_overrides):
        """Run a real clean and return ``(result, output_path)``."""
        engine = engine or self.make_engine(**engine_overrides)
        output_path = self.temp_dir / f"{input_path.stem}_cleaned.docx"
        processor = DocxProcessor(engine)
        result = processor.process(input_path, output_path)
        self.assertEqual(result.errors, [], "processing reported errors")
        return result, output_path

    def detect(self, input_path: Path, engine: DetectionEngine | None = None, **engine_overrides):
        """Run a dry-run pass and return the detections."""
        engine = engine or self.make_engine(**engine_overrides)
        processor = DocxProcessor(engine, dry_run=True)
        result = processor.process(input_path, self.temp_dir / "unused.docx")
        self.assertEqual(result.errors, [], "processing reported errors")
        return result.detections

    def part(self, docx_path: Path, part_name: str = "word/document.xml") -> str:
        return db.read_part(docx_path, part_name)

    def root(self, docx_path: Path, part_name: str = "word/document.xml"):
        """Parse one part of a .docx and return its root element."""
        return etree.fromstring(self.part(docx_path, part_name).encode("utf-8"))

    def paragraph_texts(self, docx_path: Path, part_name: str = "word/document.xml") -> list[str]:
        """Every non-empty paragraph string in one part, in document order."""
        root = self.root(docx_path, part_name)
        texts = [paragraph_text(para).strip() for para in iter_paragraphs(root, True)]
        return [text for text in texts if text]

    def count_tags(self, docx_path: Path, tag: str, part_name: str = "word/document.xml") -> int:
        """Count elements with a local ``w:`` tag name in one part."""
        return sum(1 for _ in self.root(docx_path, part_name).iter(f"{W}{tag}"))
