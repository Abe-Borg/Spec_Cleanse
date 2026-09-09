"""The corpus harness: recording decisions, and diffing two recordings."""

import copy
import unittest

from docx_xml import load_config
from tools.actions import content_digest
from tools.corpus_compare import (
    Decision,
    diff,
    format_diff,
    load,
    record_one,
    save,
)

from tests import docx_builder as db
from tests.support import CONFIG_PATH, DocxTestCase


DOCUMENT = db.document(
    db.text_para("SECTION 21 13 13 - WET-PIPE SPRINKLER SYSTEMS"),
    db.text_para("Provide sprinklers throughout the data hall."),
    db.text_para("Retain subparagraph below for wet-pipe systems."),
    db.text_para("Provide two [Verify quantity with Owner] spare sprinklers."),
)


class RecordingTests(DocxTestCase):

    def record(self, document_xml=DOCUMENT, **kwargs):
        return record_one(self.build(document_xml), load_config(CONFIG_PATH), **kwargs)

    def test_every_action_is_recorded_with_the_rule_behind_it(self):
        decisions = self.record()

        actions = {d.action for d in decisions}
        self.assertEqual(actions, {"preserved", "removed", "redacted"})
        self.assertTrue(all(d.rules for d in decisions), "a decision without a rule")

    def test_two_rules_on_one_paragraph_are_one_row(self):
        # Recording per detection meant narrowing one of two rules removed a
        # baseline row, and the diff reported "no longer acted on" for content
        # the candidate still deletes.
        decisions = self.record(db.document(
            db.text_para("[Specifier: Copyright 2026 ARCOM]"),
        ))

        self.assertEqual(len(decisions), 1)
        self.assertEqual(decisions[0].categories, "copyright; specifier_note")

    def test_a_placeholder_only_paragraph_is_recorded_as_removed(self):
        decisions = self.record(db.document(
            db.text_para("[Verify quantity with Owner]"),
        ))

        self.assertEqual([d.action for d in decisions], ["removed"])

    def test_untouched_content_is_not_a_decision(self):
        # Only what the build acted on.  Recording every paragraph would bury
        # the differences that matter in a diff of the whole document.
        previews = [d.preview for d in self.record()]

        self.assertFalse(any("Provide sprinklers throughout" in p for p in previews))

    def test_no_text_omits_previews_but_still_distinguishes_decisions(self):
        # The privacy mode has to stay comparable.  Without a digest every
        # decision of one category in a document collapsed to one entry, and a
        # recording that lost two of three diffed as no change at all.
        decisions = self.record(db.document(
            db.text_para("[Specifier: delete note one]"),
            db.text_para("[Specifier: delete note two]"),
            db.text_para("[Specifier: delete note three]"),
        ), keep_text=False)

        self.assertEqual(len(decisions), 3)
        self.assertTrue(all(d.preview == "" for d in decisions))
        self.assertEqual(len({d.content for d in decisions}), 3)
        self.assertEqual(len(diff(decisions, decisions[:1])), 2)

    def test_a_long_paragraph_is_previewed_not_stored_whole(self):
        long_note = "Retain subparagraph below " + ("x" * 400)
        decisions = self.record(db.document(db.text_para(long_note)))

        self.assertTrue(decisions)
        self.assertLess(len(decisions[0].preview), 120)
        self.assertTrue(decisions[0].preview.endswith("..."))

    def test_an_unreadable_file_is_recorded_as_an_error_row(self):
        broken = self.temp_dir / "broken.docx"
        broken.write_bytes(b"not a zip")

        decisions = record_one(broken, load_config(CONFIG_PATH))

        self.assertEqual(len(decisions), 1)
        self.assertEqual(decisions[0].action, "error")

    def test_repeated_boilerplate_keeps_its_multiplicity(self):
        note = db.text_para("[Specifier: delete this note before issue]")
        decisions = self.record(db.document(note, note, note))

        self.assertEqual(len(decisions), 3)
        self.assertEqual(len(diff(decisions, decisions[:1])), 2)

    def test_a_recording_survives_a_round_trip_through_json(self):
        decisions = self.record()
        path = self.temp_dir / "baseline.json"

        save(decisions, path)

        self.assertEqual(load(path), decisions)


class DiffTests(unittest.TestCase):

    def decision(self, preview="A note.", action="removed", rules="pattern X",
                 categories="specifier_note", document="a.docx"):
        return Decision(document, "document.xml", action, categories, rules,
                        content_digest(preview), preview)

    def test_an_unchanged_recording_diffs_to_nothing(self):
        rows = [self.decision(), self.decision(preview="Another.")]

        self.assertEqual(diff(rows, list(rows)), [])
        self.assertEqual(format_diff([]), "No decisions changed.")

    def test_a_decision_the_candidate_no_longer_makes(self):
        # What narrowing a rule looks like: content that used to be removed
        # is not acted on at all any more.
        differences = diff([self.decision()], [])

        self.assertEqual(len(differences), 1)
        self.assertEqual(differences[0].kind, "no longer acted on")
        self.assertIn("was: removed", differences[0].describe())

    def test_a_decision_the_candidate_newly_makes(self):
        differences = diff([], [self.decision()])

        self.assertEqual(differences[0].kind, "newly acted on")
        self.assertIn("now: removed", differences[0].describe())

    def test_the_same_content_acted_on_differently(self):
        before = self.decision(action="removed")
        after = self.decision(action="redacted")

        differences = diff([before], [after])

        self.assertEqual(differences[0].kind, "action changed")
        self.assertIn("was: removed", differences[0].describe())
        self.assertIn("now: redacted", differences[0].describe())

    def test_the_same_content_removed_by_a_different_rule(self):
        differences = diff([self.decision(rules="pattern X")],
                           [self.decision(rules="pattern Y")])

        self.assertEqual(differences[0].kind, "action changed")

    def test_losing_one_of_several_identical_decisions_is_one_change(self):
        rows = [self.decision(), self.decision(), self.decision()]

        differences = diff(rows, rows[:2])

        self.assertEqual(len(differences), 1)
        self.assertEqual(differences[0].kind, "no longer acted on")

    def test_gaining_one_of_several_identical_decisions_is_one_change(self):
        rows = [self.decision(), self.decision()]

        differences = diff(rows[:1], rows)

        self.assertEqual(len(differences), 1)
        self.assertEqual(differences[0].kind, "newly acted on")

    def test_identity_ignores_position_so_a_shift_is_not_a_change(self):
        # Narrowing a rule changes how many paragraphs go, shifting every
        # later one.  A position-keyed diff would call the whole document
        # changed and bury the real differences.
        rows = [self.decision(preview="First."), self.decision(preview="Second.")]

        self.assertEqual(diff(rows, list(reversed(rows))), [])

    def test_the_same_text_in_two_documents_is_two_decisions(self):
        differences = diff(
            [self.decision(document="a.docx")],
            [self.decision(document="b.docx")],
        )

        self.assertEqual(len(differences), 2)

    def test_the_summary_counts_each_kind_and_asks_for_a_verdict(self):
        text = format_diff(diff([self.decision()], [self.decision(preview="New.")]))

        self.assertIn("2 changed decision(s)", text)
        self.assertIn("1 newly acted on", text)
        self.assertIn("1 no longer acted on", text)
        self.assertIn("needs a verdict", text)


class EndToEndDiffTests(DocxTestCase):

    def test_narrowing_a_rule_shows_up_as_a_decision_no_longer_made(self):
        # A miniature of what W02 will do, run through the whole harness.
        path = self.build(db.document(
            db.text_para("Retain subparagraph below for wet-pipe systems."),
            db.text_para("[Specifier: delete this note before issue]"),
        ))
        config = load_config(CONFIG_PATH)

        baseline = record_one(path, config)

        narrowed = copy.deepcopy(config)
        narrowed["editorial_artifacts"]["text_patterns"] = [
            p for p in narrowed["editorial_artifacts"]["text_patterns"]
            if "retain" not in p
        ]
        candidate = record_one(path, narrowed)

        differences = diff(baseline, candidate)

        self.assertEqual(len(differences), 1)
        self.assertEqual(differences[0].kind, "no longer acted on")
        self.assertIn("Retain subparagraph below", differences[0].describe())


if __name__ == "__main__":
    unittest.main()
