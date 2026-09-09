"""The shared walker: what a build would actually do, one row per action.

Both censuses and the corpus harness were first built on ``result.detections``
— what the detectors *found* — and both were wrong in the same way. A paragraph
can carry several detections and still be one action. These pin the difference.
"""

import unittest

from docx_xml import load_config
from tools.actions import (
    PRESERVED,
    REDACTED,
    REMOVED,
    RUN_REMOVED,
    content_digest,
    iter_actions,
)

from tests import docx_builder as db
from tests.support import CONFIG_PATH, DocxTestCase


class ActionWalkerTests(DocxTestCase):

    def actions(self, document_xml):
        return iter_actions(self.build(document_xml), load_config(CONFIG_PATH))

    def test_two_rules_on_one_paragraph_are_one_action(self):
        # The bug this file exists to prevent: a note matching both a specifier
        # rule and a copyright rule produced two detections and one removed
        # paragraph, and counting detections inflated every total built on it.
        actions = self.actions(db.document(
            db.text_para("[Specifier: Copyright 2026 ARCOM]"),
        ))

        self.assertEqual(len(actions), 1)
        self.assertEqual(actions[0].action, REMOVED)
        self.assertEqual(actions[0].categories, ("copyright", "specifier_note"))
        self.assertEqual(len(actions[0].rules), 2, "both rules kept as evidence")

    def test_a_placeholder_only_paragraph_is_removed_not_redacted(self):
        # _redaction_spans returns an empty list to say "cutting every
        # placeholder leaves nothing, remove the paragraph".  Reading that as a
        # redaction misstates a whole-paragraph deletion.
        actions = self.actions(db.document(
            db.text_para("[Verify quantity with Owner]"),
        ))

        self.assertEqual([a.action for a in actions], [REMOVED])

    def test_a_placeholder_inside_a_requirement_is_redacted(self):
        actions = self.actions(db.document(
            db.text_para("Provide two [Verify quantity with Owner] spare sprinklers."),
        ))

        self.assertEqual([a.action for a in actions], [REDACTED])

    def test_a_preserved_paragraph_is_recorded_as_preserved(self):
        actions = self.actions(db.document(
            db.text_para("SECTION 21 13 13 - WET-PIPE SPRINKLER SYSTEMS"),
        ))

        self.assertEqual([a.action for a in actions], [PRESERVED])

    def test_a_hidden_run_in_a_surviving_paragraph_is_a_run_removal(self):
        actions = self.actions(db.document(db.para(
            db.run("Sprinkler piping shall be tested at 200 psi. "),
            db.run("check this later", vanish=True),
        )))

        self.assertEqual([a.action for a in actions], [RUN_REMOVED])
        self.assertEqual(actions[0].text, "check this later")

    def test_untouched_content_produces_no_action(self):
        self.assertEqual(self.actions(db.document(
            db.text_para("Provide sprinklers throughout the data hall."),
        )), [])

    def test_actions_carry_their_part(self):
        actions = self.actions(db.document(
            db.text_para("[Specifier: delete this note before issue]"),
        ))

        self.assertEqual(actions[0].part, "document.xml")

    def test_an_unreadable_package_raises_for_the_caller_to_handle(self):
        broken = self.temp_dir / "broken.docx"
        broken.write_bytes(b"not a zip")

        with self.assertRaises(Exception):
            iter_actions(broken, load_config(CONFIG_PATH))


class DigestTests(unittest.TestCase):

    def test_the_same_text_digests_the_same(self):
        self.assertEqual(content_digest("A note."), content_digest("A note."))

    def test_different_text_digests_differently(self):
        self.assertNotEqual(content_digest("One."), content_digest("Two."))

    def test_whitespace_is_normalised_first(self):
        # Extraction can differ in whitespace where the meaning does not.
        self.assertEqual(content_digest("A  note."), content_digest("A note."))

    def test_the_digest_does_not_contain_the_text(self):
        self.assertNotIn("note", content_digest("A note."))


if __name__ == "__main__":
    unittest.main()
