"""Configuration loading: encoding, validation, and error reporting."""

import unittest
from pathlib import Path

import yaml
from lxml import etree

from detection import SUPERSEDED_PATTERNS, DetectionEngine, config_notices
from docx_xml import W_NS, load_config

from tests import docx_builder as db
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

    # -- Boolean switches ---------------------------------------------------

    def test_a_quoted_false_is_refused_rather_than_read_as_true(self):
        # The one that matters most: every reader asks a plain truthiness
        # question, so the string 'false' switches formatting-only removal ON —
        # the only path that removes text on no content evidence — for someone
        # who wrote that it should be off.
        with self.assertRaisesRegex(ValueError, "must be true or false"):
            load_config(self._write(
                "specifier_notes:\n  formatting_only_removal: 'false'\n"
            ))

    def test_a_quoted_false_enabled_is_refused_too(self):
        with self.assertRaisesRegex(ValueError, "must be true or false"):
            load_config(self._write("specifier_notes:\n  enabled: 'false'\n"))

    def test_a_number_is_not_a_boolean(self):
        with self.assertRaisesRegex(ValueError, "must be true or false"):
            load_config(self._write("specifier_notes:\n  formatting_only_removal: 1\n"))

    def test_a_real_boolean_is_accepted(self):
        # The guard against the cases above being satisfied by refusing
        # everything.
        config = load_config(self._write(
            "specifier_notes:\n  enabled: false\n  formatting_only_removal: true\n"
        ))
        self.assertIs(config["specifier_notes"]["enabled"], False)
        self.assertIs(config["specifier_notes"]["formatting_only_removal"], True)

    def test_the_message_says_why_quoting_broke_it(self):
        with self.assertRaises(ValueError) as caught:
            load_config(self._write("specifier_notes:\n  enabled: 'no'\n"))
        self.assertIn("string", str(caught.exception))

    # -- Editorial colours --------------------------------------------------

    def test_a_leading_hash_is_normalised_rather_than_rejected(self):
        # One possible meaning, and it is how every other tool writes a hex
        # colour.  Rejecting it would be defensible; silently matching nothing
        # is not, which is what happened before.
        config = load_config(self._write(
            "specifier_notes:\n  formatting_signals:\n    colors: ['#FF0000']\n"
        ))
        self.assertEqual(
            config["specifier_notes"]["formatting_signals"]["colors"], ["FF0000"]
        )

    def test_a_normalised_colour_actually_matches_a_run(self):
        # Normalising the config and never checking the effect would leave the
        # original defect in place with a passing test over it.
        config = load_config(self._write(
            "specifier_notes:\n  enabled: true\n  formatting_only_removal: true\n"
            "  text_patterns: []\n  formatting_signals:\n    colors: ['#FF0000']\n"
        ))
        run = etree.fromstring(
            db.run("Coordinate with Division 26.", italic=True, color="FF0000")
            .replace("<w:r>", f'<w:r xmlns:w="{W_NS}">', 1).encode("utf-8")
        )

        detections = DetectionEngine(config).detect_in_element(
            run, "Coordinate with Division 26."
        )

        self.assertTrue(any("Color" in (d.reason or "") for d in detections))

    def test_lower_case_hex_is_normalised_to_what_word_writes(self):
        config = load_config(self._write(
            "specifier_notes:\n  formatting_signals:\n    colors: ['ff0000']\n"
        ))
        self.assertEqual(
            config["specifier_notes"]["formatting_signals"]["colors"], ["FF0000"]
        )

    def test_prose_is_not_a_colour(self):
        # Was accepted and matched nothing, forever.
        with self.assertRaisesRegex(ValueError, "not a colour"):
            load_config(self._write(
                "specifier_notes:\n  formatting_signals:\n    colors: ['bright red']\n"
            ))

    def test_a_number_is_not_a_colour(self):
        # Was an AttributeError per file at detection time, naming no section,
        # key or line — the file simply failed.
        with self.assertRaisesRegex(ValueError, "expected a colour"):
            load_config(self._write(
                "specifier_notes:\n  formatting_signals:\n    colors: [255]\n"
            ))

    def test_the_wrong_number_of_digits_is_refused(self):
        with self.assertRaisesRegex(ValueError, "six hexadecimal digits"):
            load_config(self._write(
                "specifier_notes:\n  formatting_signals:\n    colors: ['FF00']\n"
            ))

    def test_the_colour_error_names_the_section_key_and_index(self):
        with self.assertRaises(ValueError) as caught:
            load_config(self._write(
                "specifier_notes:\n  formatting_signals:\n"
                "    colors: ['FF0000', 'nope']\n"
            ))
        self.assertIn("specifier_notes.formatting_signals.colors[1]", str(caught.exception))

    def test_colours_must_be_a_list(self):
        with self.assertRaisesRegex(ValueError, "must be a list of colours"):
            load_config(self._write(
                "specifier_notes:\n  formatting_signals:\n    colors: 'FF0000'\n"
            ))

    def test_a_style_name_absent_from_any_one_document_is_not_an_error(self):
        # §15.2 is explicit: styles differ across templates, so a name that no
        # sample document happens to use says nothing about the config.
        load_config(self._write(
            "specifier_notes:\n  styles: ['NoSuchStyleAnywhere', 'CMT']\n"
        ))


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

    def test_a_disabled_section_raises_no_notice(self):
        # Every detector short-circuits on `enabled`, so a superseded pattern
        # under a disabled section cannot reach any output.  A notice about it
        # would warn about something that provably did not happen — and since
        # a notice makes every file in a run need review, it would be that
        # warning repeated over the whole batch.
        section, superseded = next(iter(SUPERSEDED_PATTERNS.items()))
        pattern = next(iter(superseded))

        self.assertEqual(
            config_notices({section: {"enabled": False, "text_patterns": [pattern]}}),
            [],
        )

    def test_the_same_pattern_in_an_enabled_section_still_raises_one(self):
        # The guard against the case above being satisfied by a check that
        # never reports anything.
        section, superseded = next(iter(SUPERSEDED_PATTERNS.items()))
        pattern = next(iter(superseded))

        self.assertEqual(
            len(config_notices({section: {"text_patterns": [pattern]}})), 1
        )

    def test_formatting_only_removal_under_a_disabled_section_raises_nothing(self):
        # The setting only reaches SpecifierNoteDetector's scoring, which the
        # engine never runs when the section is off.
        self.assertEqual(
            config_notices({
                "specifier_notes": {"enabled": False, "formatting_only_removal": True}
            }),
            [],
        )

    def test_formatting_only_removal_in_an_enabled_section_still_raises_one(self):
        self.assertEqual(
            len(config_notices({"specifier_notes": {"formatting_only_removal": True}})), 1
        )


if __name__ == "__main__":
    unittest.main()
