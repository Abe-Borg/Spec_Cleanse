"""Configuration loading: encoding, validation, and error reporting."""

import unittest
from pathlib import Path

import yaml

from detection import SUPERSEDED_PATTERNS, DetectionEngine, config_notices
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


class ConfigNoticeTests(unittest.TestCase):
    """What a user running an older patterns.yaml is told about it.

    apppaths prefers an existing executable-adjacent or per-user file over the
    bundled default, so updating the application does not update the rules. A
    configuration copied before a rule was narrowed keeps the broad version,
    and nothing else would say so.
    """

    def test_the_shipped_configuration_says_nothing(self):
        self.assertEqual(config_notices(load_config(CONFIG_PATH)), [])

    def test_a_superseded_copyright_rule_is_named(self):
        notices = config_notices({
            "copyright_notices": {
                "text_patterns": [r"may\s+not\s+be\s+reproduced",
                                  r"all\s+rights\s+reserved"],
            },
        })

        self.assertEqual(len(notices), 1)
        self.assertIn("may", notices[0])
        self.assertIn("Shop Drawings", notices[0], "says what the rule took")
        self.assertIn("never overwritten", notices[0], "says edits are safe")

    def test_a_superseded_editorial_rule_is_named(self):
        notices = config_notices({
            "editorial_artifacts": {
                "text_patterns": [r"^\s*(?:select|choose)\s+one\b(?!-)"],
            },
        })

        self.assertEqual(len(notices), 1)
        self.assertIn("listed manufacturers", notices[0])

    def test_an_edited_rule_is_left_alone(self):
        # Matched on the exact prior string, so someone who tuned the rule
        # themselves is not told their own work is out of date.
        notices = config_notices({
            "copyright_notices": {
                "text_patterns": [r"may\s+not\s+be\s+reproduced\s+at\s+all"],
            },
        })

        self.assertEqual(notices, [])

    def test_formatting_only_removal_being_on_is_reported(self):
        notices = config_notices({
            "specifier_notes": {"formatting_only_removal": True},
        })

        self.assertEqual(len(notices), 1)
        self.assertIn("formatting_only_removal", notices[0])
        self.assertIn("census_formatting", notices[0], "points at the measurement")

    def test_formatting_only_removal_being_off_is_not_worth_saying(self):
        self.assertEqual(
            config_notices({"specifier_notes": {"formatting_only_removal": False}}),
            [],
        )

    def test_an_empty_configuration_says_nothing(self):
        self.assertEqual(config_notices({}), [])

    def test_every_superseded_pattern_is_one_the_project_used_to_ship(self):
        # The check is worthless if the recorded strings drift from what was
        # actually shipped, and a typo here would silently stop warning.
        for section, superseded in SUPERSEDED_PATTERNS.items():
            notices = config_notices({section: {"text_patterns": list(superseded)}})
            self.assertEqual(
                len(notices), len(superseded),
                f"{section}: not every recorded pattern was recognised",
            )
