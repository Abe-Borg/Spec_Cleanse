"""Configuration loading: encoding, validation, and error reporting."""

import unittest
from pathlib import Path

import yaml

from detection import DetectionEngine
from docx_xml import load_config

from tests.support import CONFIG_PATH, DocxTestCase


class ConfigEncodingTests(DocxTestCase):
    """patterns.yaml is UTF-8 and must be read as UTF-8 on every platform."""

    def _write(self, text: str) -> Path:
        path = self.temp_dir / "patterns.yaml"
        path.write_text(text, encoding="utf-8")
        return path

    def test_non_ascii_patterns_survive_the_load(self):
        path = self._write(
            "copyright_notices:\n"
            "  enabled: true\n"
            "  text_patterns:\n"
            "    - '©'\n"
            "preserve_patterns:\n"
            "  enabled: true\n"
            "  text_patterns:\n"
            "    - '^\\s*part\\s+\\d+\\s*[-–—:]'\n"
        )
        config = load_config(path)
        engine = DetectionEngine(config)

        copyright_pattern = engine.detectors[1].compiled_patterns[0]
        self.assertTrue(copyright_pattern.search("© 2026 ARCOM"))

        preserve_pattern = engine.preserve_detector.compiled_patterns[0]
        self.assertTrue(preserve_pattern.search("PART 1 – GENERAL"))

    def test_platform_default_encoding_would_mangle_the_file(self):
        """Guards the fix: cp1252 (the Windows default) corrupts these patterns."""
        path = self._write(
            "copyright_notices:\n"
            "  text_patterns:\n"
            "    - '©'\n"
        )
        correct = load_config(path)["copyright_notices"]["text_patterns"][0]
        mangled = yaml.safe_load(
            path.read_bytes().decode("cp1252")
        )["copyright_notices"]["text_patterns"][0]

        self.assertEqual(correct, "©")
        self.assertNotEqual(mangled, correct)
        self.assertIsNone(__import__("re").compile(mangled).search("© 2026 ARCOM"))


class ConfigValidationTests(DocxTestCase):
    """A broken config must fail loudly, once, naming what is wrong."""

    def _write(self, text: str) -> Path:
        path = self.temp_dir / "patterns.yaml"
        path.write_text(text, encoding="utf-8")
        return path

    def test_missing_file(self):
        with self.assertRaises(FileNotFoundError):
            load_config(self.temp_dir / "nope.yaml")

    def test_empty_file(self):
        with self.assertRaisesRegex(ValueError, "empty"):
            load_config(self._write(""))

    def test_not_a_mapping(self):
        with self.assertRaisesRegex(ValueError, "mapping"):
            load_config(self._write("- just\n- a\n- list\n"))

    def test_invalid_yaml(self):
        with self.assertRaisesRegex(ValueError, "valid YAML"):
            load_config(self._write("specifier_notes: [unclosed\n"))

    def test_invalid_regex_names_section_and_index(self):
        path = self._write(
            "specifier_notes:\n"
            "  text_patterns:\n"
            "    - 'fine'\n"
            "    - '[unclosed'\n"
        )
        with self.assertRaises(ValueError) as caught:
            load_config(path)
        message = str(caught.exception)
        self.assertIn("specifier_notes.text_patterns[1]", message)

    def test_shipped_config_is_valid(self):
        DetectionEngine(load_config(CONFIG_PATH))


if __name__ == "__main__":
    unittest.main()
