"""Comparison happens inside a location, and duplicate text does not mask a loss.

Two defects motivated this, both measured against a real clean before anything
was changed:

* A requirement deleted from the body was reported as **verified**, because an
  identical hidden paragraph in a header carried authority the body's copy did
  not.  ``extract_paragraphs()`` returned one flat list across every part, so
  the two were interchangeable entries in a sequence diff.
* A **correct** clean was reported as damage.  Where a location holds the same
  text twice, ``difflib`` aligns on the longest matching block, not on evidence,
  so removing a hidden note and keeping the plain requirement was blamed on the
  plain one.

Both are the same underlying mistake — treating a text match as an identity —
and both are pinned here in the direction that matters: a loss must be caught,
and a correct clean must still pass.
"""

import unittest

from docx_xml import iter_paragraphs, paragraph_text
from verify import verify_clean

from tests import docx_builder as db
from tests.support import DocxTestCase


NOTE = "Delete before issue."
REQ = "Provide fire pumps rated 1500 gpm at 100 psi."


class LocationIdentity(DocxTestCase):
    """A part or note is compared against itself, never against another."""

    def compare(self, before, after, before_parts=None, after_parts=None, name="loc"):
        source = self.build(before, before_parts, name=f"{name}_in.docx")
        output = self.build(after, after_parts, name=f"{name}_out.docx")
        return verify_clean(source, output, engine=self.make_engine())

    def test_a_header_copy_does_not_authorize_losing_the_body_copy(self):
        """The measured false negative: authority is not transferable between parts.

        The header's copy is hidden, so policy would permit losing *it*.  The
        body's copy is plain and was the one deleted.  Judged across a flat
        list, the header's authority excused the body's loss and the run
        verified clean.
        """
        header = {"word/header1.xml": db.header(db.para(db.run(NOTE, vanish=True)))}
        result = self.compare(
            db.document(db.text_para("Anchor."), db.text_para(NOTE)),
            db.document(db.text_para("Anchor.")),
            header, header, name="crosspart",
        )

        self.assertEqual(len(result.unexpected_removals), 1, result.removed)
        self.assertFalse(result.passed)

    def test_one_note_does_not_account_for_another(self):
        """``footnotes.xml`` is one part holding many independent stories."""
        result = self.compare(
            db.document(db.text_para("Body.")),
            db.document(db.text_para("Body.")),
            {"word/footnotes.xml": db.footnotes(db.text_para(REQ), db.text_para(REQ))},
            {"word/footnotes.xml": db.footnotes(db.text_para(REQ))},
            name="notes",
        )

        self.assertEqual(len(result.unexpected_removals), 1, result.removed)
        self.assertFalse(result.passed)

    def test_the_glossary_is_a_different_part_from_the_document(self):
        """Both are named ``document.xml``; the full part path keeps them apart."""
        glossary = {"word/glossary/document.xml": db.document(db.text_para(REQ))}
        result = self.compare(
            db.document(db.text_para(REQ), db.text_para("Other.")),
            db.document(db.text_para("Other.")),
            glossary, glossary, name="glossary",
        )

        self.assertEqual(len(result.unexpected_removals), 1, result.removed)
        self.assertFalse(result.passed)

    def test_a_part_missing_from_the_output_is_reported(self):
        result = self.compare(
            db.document(db.text_para("Body.")),
            db.document(db.text_para("Body.")),
            {"word/header1.xml": db.header(db.text_para("Header requirement."))},
            None, name="lostpart",
        )

        self.assertEqual(len(result.unexpected_removals), 1, result.removed)
        self.assertFalse(result.passed)

    def test_a_part_only_in_the_output_is_reported_as_added(self):
        result = self.compare(
            db.document(db.text_para("Body.")),
            db.document(db.text_para("Body.")),
            None,
            {"word/header1.xml": db.header(db.text_para("Invented header."))},
            name="newpart",
        )

        self.assertEqual(result.added, ["Invented header."])
        self.assertFalse(result.passed)

    def test_an_unchanged_document_passes(self):
        body = db.document(db.text_para("Body."), db.text_para(REQ))
        parts = {"word/header1.xml": db.header(db.text_para("Header.")),
                 "word/footnotes.xml": db.footnotes(db.text_para("Note."))}
        result = self.compare(body, body, parts, parts, name="unchanged")

        self.assertTrue(result.passed, result.removed)
        self.assertEqual(result.removed, [])
        self.assertEqual(result.modified, [])
        self.assertEqual(result.added, [])

    def test_text_the_output_invented_is_reported(self):
        """Insert-only output: nothing was lost, and the document still changed."""
        result = self.compare(
            db.document(db.text_para("Body.")),
            db.document(db.text_para("Body."), db.text_para("Invented requirement.")),
            name="inserted",
        )

        self.assertEqual(result.added, ["Invented requirement."])
        self.assertEqual(result.removed, [])
        self.assertFalse(result.passed)

    def test_an_empty_document_passes(self):
        empty = db.document(db.text_para(" "))
        result = self.compare(empty, empty, name="empty")

        self.assertTrue(result.passed, result.removed)
        self.assertEqual(result.input_paragraph_count, 0)


class ParagraphStyleAuthority(DocxTestCase):
    """Identical runs can still differ in what policy permits."""

    STYLES = db.styles(db.style_def("CMT", name="CMT"))
    TEXT = "Coordinate hangers with the structural drawings."

    def compare(self, after, name):
        before = db.document(
            db.text_para(self.TEXT), db.text_para(self.TEXT, style="CMT"),
            db.text_para("Anchor."))
        parts = {"word/styles.xml": self.STYLES}
        source = self.build(before, parts, name=f"{name}_in.docx")
        output = self.build(after, parts, name=f"{name}_out.docx")
        return verify_clean(source, output, engine=self.make_engine())

    def test_removing_the_styled_copy_is_a_correct_clean(self):
        """An editorial paragraph style is authority the runs do not carry.

        Both copies hold character-for-character identical runs, so a run-level
        identity cannot tell them apart, and removing the styled one was blamed
        on the plain one.
        """
        result = self.compare(
            db.document(db.text_para(self.TEXT), db.text_para("Anchor.")), "style_ok")

        self.assertTrue(result.passed, result.removed)
        self.assertEqual(len(result.expected_removals), 1)

    def test_removing_the_plain_copy_is_still_damage(self):
        result = self.compare(
            db.document(db.text_para(self.TEXT, style="CMT"), db.text_para("Anchor.")),
            "style_bad")

        self.assertEqual(len(result.unexpected_removals), 1, result.removed)
        self.assertFalse(result.passed)


class CombinedRemovalAndRedaction(DocxTestCase):
    """A paragraph can carry both an authorized run and a placeholder.

    Expecting only the placeholder cut gives a text no correct clean produces,
    so the case fell through to a differ, which then blamed the wrong copy of a
    repeated run.  These are cleaned for real: the point is what the cleaner
    actually writes.
    """

    CASES = {
        "hidden twin beside a redacted requirement": db.para(
            db.run("Keep. "), db.run("Keep. ", vanish=True),
            db.run("Provide [Verify quantity] units.")),
        "hidden run whose text repeats, plus a placeholder": db.para(
            db.run("Isolate riser. "), db.run("Isolate riser. ", vanish=True),
            db.run("Provide [Insert model] valves.")),
        "placeholder inside one of two identical runs": db.para(
            db.run("Provide [Verify quantity] valves. "),
            db.run("Provide [Verify quantity] valves.", vanish=True)),
    }

    def test_each_case_verifies(self):
        for label, xml in self.CASES.items():
            with self.subTest(label):
                engine = self.make_engine()
                source = self.build(
                    db.document(xml, db.text_para("Anchor.")),
                    name=f"combined{abs(hash(label))}.docx")
                _, cleaned = self.clean(source, engine)
                result = verify_clean(source, cleaned, engine=engine)

                self.assertTrue(
                    result.passed,
                    f"{label}: {result.modified} {result.removed} {result.added}")


class RedactionShapes(DocxTestCase):
    """Repeated words, split runs, several placeholders, edge whitespace."""

    CASES = {
        "repeated word around a placeholder":
            "Provide valves [Verify quantity] and valves for the riser.",
        "two placeholders":
            "Provide [Verify quantity] valves and [Insert model] actuators.",
        "adjacent placeholders":
            "Provide [Verify quantity][Insert model] valves.",
        "repeated identical placeholders":
            "Provide [Verify quantity] valves and [Verify quantity] actuators.",
        "placeholder at the very start":
            "[Verify quantity] valves are required.",
        "placeholder at the very end":
            "Provide valves per schedule [Verify quantity]",
        "long redaction leaving a short survivor":
            "Provide [Verify quantity with the Owner and the AHJ prior to bid] units.",
        "text repeated verbatim twice":
            "Isolate the riser. Isolate the riser. [Verify quantity]",
    }

    def test_each_shape_verifies(self):
        for label, text in self.CASES.items():
            with self.subTest(label):
                engine = self.make_engine()
                source = self.build(
                    db.document(db.text_para(text), db.text_para("Anchor.")),
                    name=f"shape{abs(hash(label))}.docx")
                _, cleaned = self.clean(source, engine)
                result = verify_clean(source, cleaned, engine=engine)

                self.assertTrue(result.passed, f"{label}: {result.modified}")

    def test_a_placeholder_split_across_runs_verifies(self):
        engine = self.make_engine()
        source = self.build(db.document(
            db.para(db.run("Provide two ["), db.run("Verify quan"),
                    db.run("tity] spare filters.")),
            db.text_para("Anchor."),
        ), name="split.docx")
        _, cleaned = self.clean(source, engine)
        result = verify_clean(source, cleaned, engine=engine)

        self.assertTrue(result.passed, result.modified)
        self.assertEqual(
            self.paragraph_texts(cleaned), ["Provide two spare filters.", "Anchor."])

    def test_leading_and_trailing_whitespace_is_kept(self):
        """Read unstripped: the padding is the thing under test.

        ``paragraph_texts`` trims, which is right for comparing content and
        wrong here — cutting a placeholder must not take the paragraph's own
        leading or trailing space with it.
        """
        engine = self.make_engine()
        source = self.build(db.document(
            db.para(db.run("   Provide [Verify quantity] valves.   ")),
            db.text_para("Anchor."),
        ), name="edgews.docx")
        _, cleaned = self.clean(source, engine)
        result = verify_clean(source, cleaned, engine=engine)
        raw = [
            paragraph_text(para)
            for para in iter_paragraphs(self.root(cleaned), True)
            if paragraph_text(para).strip()
        ]

        self.assertTrue(result.passed, result.modified)
        self.assertEqual(raw, ["   Provide valves.   ", "Anchor."])


class DuplicateTextAttribution(DocxTestCase):
    """Which occurrence went is decided by evidence, not by the alignment."""

    def build_pair(self, before, after, name):
        source = self.build(before, name=f"{name}_in.docx")
        output = self.build(after, name=f"{name}_out.docx")
        return verify_clean(source, output, engine=self.make_engine())

    def test_removing_the_hidden_twin_of_a_requirement_is_a_correct_clean(self):
        """The measured false alarm.

        ``difflib`` anchors on the longest matching block, so with the plain
        copy first it called *that* one deleted and reported an unexplained
        removal for a clean that did exactly what policy asked.
        """
        result = self.build_pair(
            db.document(db.text_para(NOTE), db.para(db.run(NOTE, vanish=True)),
                        db.text_para("Anchor.")),
            db.document(db.text_para(NOTE), db.text_para("Anchor.")),
            "dup_ok",
        )

        self.assertTrue(result.passed, result.removed)
        self.assertEqual(len(result.expected_removals), 1)
        self.assertEqual(result.expected_removals[0].category, "hidden_text")

    def test_removing_the_plain_twin_is_still_damage(self):
        """The discriminating case: same text, and the wrong copy went."""
        result = self.build_pair(
            db.document(db.text_para(NOTE), db.para(db.run(NOTE, vanish=True)),
                        db.text_para("Anchor.")),
            db.document(db.para(db.run(NOTE, vanish=True)), db.text_para("Anchor.")),
            "dup_bad",
        )

        self.assertEqual(len(result.unexpected_removals), 1, result.removed)
        self.assertFalse(result.passed)

    def test_duplicate_unchanged_paragraphs_keep_their_multiplicity(self):
        result = self.build_pair(
            db.document(db.text_para(REQ), db.text_para("Between."), db.text_para(REQ)),
            db.document(db.text_para(REQ), db.text_para("Between.")),
            "dupcount",
        )

        self.assertEqual(len(result.unexpected_removals), 1, result.removed)
        self.assertFalse(result.passed)

    def test_adjacent_similar_paragraphs_cannot_exchange_survivors(self):
        """A near-identical neighbour must not stand in for a lost paragraph."""
        result = self.build_pair(
            db.document(
                db.text_para("Provide valves of type A for the standpipe system."),
                db.text_para("Provide valves of type B for the standpipe system."),
                db.text_para("Anchor."),
            ),
            db.document(
                db.text_para("Provide valves of type B for the standpipe system."),
                db.text_para("Anchor."),
            ),
            "adjacent",
        )

        self.assertEqual(len(result.unexpected_removals), 1, result.removed)
        self.assertFalse(result.passed)

    def test_several_changed_paragraphs_between_two_anchors(self):
        """A run of changes between unchanged anchors is resolved deterministically."""
        source = db.document(
            db.text_para("Anchor one."),
            db.text_para("Note to Specifier: first."),
            db.text_para("Note to Specifier: second."),
            db.text_para("Anchor two."),
        )
        engine = self.make_engine()
        path = self.build(source, name="anchors.docx")
        _, cleaned = self.clean(path, engine)
        result = verify_clean(path, cleaned, engine=engine)

        self.assertTrue(result.passed, result.removed)
        self.assertEqual(len(result.expected_removals), 2)
        self.assertEqual(self.paragraph_texts(cleaned), ["Anchor one.", "Anchor two."])


if __name__ == "__main__":
    unittest.main()
