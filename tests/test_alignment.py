"""The pure-deletion fast path: when it is taken, and that it changes no verdict.

§16.4 asks for a deterministic regression for the problematic pattern, using
work counts rather than wall-clock assertions.  The work count here is the
number of `difflib.SequenceMatcher` alignments performed over paragraph
sequences: on a duplicate-heavy pure deletion it must be zero, because that is
the quadratic search the fast path exists to skip.

The equivalence tests matter more than the work count.  A fast path that is
fast and wrong is worse than no fast path, and the way this one was wrong when
first written — pairing on text, so a surviving hidden note stood in for the
plain requirement that had actually been deleted — is not something a timing
test would ever have caught.
"""

import difflib
import unittest
from unittest.mock import patch

import verify
from processor import DocxProcessor
from verify import verify_clean

from tests import docx_builder as db
from tests.support import DocxTestCase

#: A paragraph the cleaner removes, and one it must keep.
NOTE = "Retain or delete manufacturers below."
REQUIREMENT = "Provide listed sprinklers."
HEADING = "PART 1 - GENERAL"


class _CountingMatcher(difflib.SequenceMatcher):
    """A SequenceMatcher that records how many times one was built."""

    built = 0

    def __init__(self, *args, **kwargs):
        type(self).built += 1
        super().__init__(*args, **kwargs)


class FastPathTests(DocxTestCase):

    def _alignments(self, source_body, output_body) -> tuple[int, object]:
        """Verify a hand-built pair, returning (alignments performed, result).

        Source and output are built independently, never by running the
        cleaner: a fast path validated against the cleaner that produced its
        input would agree with whatever the cleaner did.
        """
        engine = self.make_engine()
        source = db.build_docx(self.temp_dir / "in.docx", db.document(*source_body))
        output = db.build_docx(self.temp_dir / "out.docx", db.document(*output_body))

        _CountingMatcher.built = 0
        with patch.object(verify.difflib, "SequenceMatcher", _CountingMatcher):
            result = verify_clean(source, output, engine=engine)
        return _CountingMatcher.built, result

    def _repeated(self, count: int) -> list[str]:
        """A document with three distinct paragraph texts and nothing unique."""
        body = []
        for index in range(count):
            body.append(db.text_para(
                (HEADING, NOTE, REQUIREMENT)[index % 3]
            ))
        return body

    def test_a_duplicate_heavy_pure_deletion_needs_no_alignment(self):
        # The problematic pattern, as a work count: not "this was quick" but
        # "the quadratic search did not run".
        source = self._repeated(60)
        output = [p for index, p in enumerate(source) if index % 3 != 1]

        alignments, result = self._alignments(source, output)

        self.assertEqual(alignments, 0)
        self.assertTrue(result.passed, result.removed)
        self.assertEqual(len(result.removed), 20)

    def test_a_modified_paragraph_falls_back_to_the_ordinary_path(self):
        # The guard against the fast path swallowing everything: it must not
        # apply where a paragraph was rewritten rather than removed.
        source = [db.text_para(HEADING),
                  db.text_para("Provide two [Verify quantity] spare sprinklers.")]
        output = [db.text_para(HEADING),
                  db.text_para("Provide two spare sprinklers.")]

        alignments, result = self._alignments(source, output)

        self.assertGreater(alignments, 0)
        self.assertEqual(len(result.modified), 1)

    def test_an_invented_paragraph_falls_back(self):
        source = [db.text_para(HEADING)]
        output = [db.text_para(HEADING), db.text_para("Provide gold sprinklers.")]

        alignments, result = self._alignments(source, output)

        self.assertGreater(alignments, 0)
        self.assertEqual(result.added, ["Provide gold sprinklers."])

    def test_a_reordering_falls_back_and_is_still_reported(self):
        # Reordering is not an in-order embedding, so the scan rejects it and
        # the ordinary comparison reports it — as V12 requires.
        source = [db.text_para(HEADING), db.text_para(REQUIREMENT),
                  db.text_para("PART 3 - EXECUTION")]
        output = [db.text_para("PART 3 - EXECUTION"), db.text_para(HEADING),
                  db.text_para(REQUIREMENT)]

        alignments, result = self._alignments(source, output)

        self.assertGreater(alignments, 0)
        self.assertFalse(result.passed)

    def test_a_reshaped_run_falls_back(self):
        # Text identical, formatting not: the signature differs, so this is not
        # a pure deletion and the scan declines it.
        source = [db.text_para(HEADING), db.text_para(REQUIREMENT)]
        output = [db.text_para(HEADING),
                  db.para(db.run(REQUIREMENT, italic=True))]

        alignments, _ = self._alignments(source, output)

        self.assertGreater(alignments, 0)


class EquivalenceTests(DocxTestCase):
    """Every verdict must be the one the ordinary comparison would have given."""

    def _both_paths(self, source_body, output_body, strip_revisions=False):
        engine = self.make_engine()
        source = db.build_docx(self.temp_dir / "in.docx", db.document(*source_body))
        output = db.build_docx(self.temp_dir / "out.docx", db.document(*output_body))

        fast = verify_clean(source, output, engine=engine,
                            strip_revisions=strip_revisions)
        with patch.object(verify, "_pure_deletion_pairing", return_value=None):
            slow = verify_clean(source, output, engine=engine,
                                strip_revisions=strip_revisions)
        return fast, slow

    def _cleaned(self, source_body, strip_revisions=False):
        """Compare both paths on an output the cleaner really produced.

        The independent-fixture rule applies to *damage* cases, where agreeing
        with the cleaner proves nothing.  Here the question is the opposite —
        whether a correct clean is reported as correct — so the clean itself is
        the fixture.
        """
        engine = self.make_engine()
        source = db.build_docx(self.temp_dir / "in.docx", db.document(*source_body))
        output = self.temp_dir / "out.docx"
        result = DocxProcessor(engine, strip_revisions=strip_revisions).process(
            source, output
        )
        self.assertEqual(result.errors, [])

        fast = verify_clean(source, output, engine=engine,
                            strip_revisions=strip_revisions)
        with patch.object(verify, "_pure_deletion_pairing", return_value=None):
            slow = verify_clean(source, output, engine=engine,
                                strip_revisions=strip_revisions)
        return fast, slow

    def _same(self, fast, slow) -> None:
        self.assertEqual(fast.passed, slow.passed)
        self.assertEqual(sorted(r.text for r in fast.removed),
                         sorted(r.text for r in slow.removed))
        self.assertEqual(sorted(r.category or "" for r in fast.removed),
                         sorted(r.category or "" for r in slow.removed))
        self.assertEqual(fast.added, slow.added)
        self.assertEqual([str(v) for v in fast.structural],
                         [str(v) for v in slow.structural])
        self.assertEqual(fast.review_categories(), slow.review_categories())

    def test_a_plain_removal_agrees(self):
        self._same(*self._both_paths(
            [db.text_para(HEADING), db.text_para(NOTE), db.text_para(REQUIREMENT)],
            [db.text_para(HEADING), db.text_para(REQUIREMENT)],
        ))

    def test_a_deleted_requirement_agrees_and_is_damage_on_both(self):
        fast, slow = self._both_paths(
            [db.text_para(HEADING), db.text_para(REQUIREMENT)],
            [db.text_para(HEADING)],
        )
        self._same(fast, slow)
        self.assertEqual(len(fast.unexpected_removals), 1)

    def test_duplicate_text_agrees(self):
        body = [db.text_para(HEADING), db.text_para(REQUIREMENT),
                db.text_para(REQUIREMENT), db.text_para(NOTE)]
        self._same(*self._both_paths(body, body[:3]))

    def test_the_hidden_twin_case_agrees_and_is_damage_on_both(self):
        # W04's discriminating case, and the one that caught this fast path
        # pairing on text: the output kept the *hidden* copy, so the visible
        # requirement is what went.
        hidden = db.para(db.run(REQUIREMENT, vanish=True))
        plain = db.text_para(REQUIREMENT)

        fast, slow = self._both_paths(
            [db.text_para(HEADING), hidden, plain],
            [db.text_para(HEADING), hidden],
        )

        self._same(fast, slow)
        self.assertEqual(len(fast.unexpected_removals), 1,
                         "losing the visible requirement must still be damage")

    def test_the_correct_clean_of_the_hidden_twin_agrees(self):
        # The other direction, so neither path is satisfied by distrusting
        # everything: removing the hidden copy is a correct clean.
        hidden = db.para(db.run(REQUIREMENT, vanish=True))
        plain = db.text_para(REQUIREMENT)

        fast, slow = self._both_paths(
            [db.text_para(HEADING), hidden, plain],
            [db.text_para(HEADING), plain],
        )

        self._same(fast, slow)
        self.assertTrue(fast.passed, fast.removed)

    def test_a_preserve_violation_agrees(self):
        self._same(*self._both_paths(
            [db.text_para(HEADING), db.text_para(REQUIREMENT)],
            [db.text_para(REQUIREMENT)],
        ))

    def test_a_tracked_deleted_row_before_its_plain_twin_agrees(self):
        # Identical signatures, different authority.  Accepting revisions
        # removes the tracked row and nothing authorizes losing the plain one,
        # so pairing on the signature alone reserved the *deleted* row as the
        # survivor and called a correct clean damage.
        fast, slow = self._cleaned([
            db.text_para(HEADING),
            db.table_of(
                db.deleted_row(db.text_para(REQUIREMENT)),
                db.row(db.text_para(REQUIREMENT)),
            ),
        ], strip_revisions=True)

        self._same(fast, slow)
        self.assertTrue(fast.passed, fast.removed)
        self.assertEqual([r.category for r in fast.removed], ["tracked_deletion"])

    def test_the_plain_row_before_its_tracked_twin_agrees(self):
        # The other order, so the fix is not just an artefact of which came
        # first.
        fast, slow = self._cleaned([
            db.text_para(HEADING),
            db.table_of(
                db.row(db.text_para(REQUIREMENT)),
                db.deleted_row(db.text_para(REQUIREMENT)),
            ),
        ], strip_revisions=True)

        self._same(fast, slow)
        self.assertTrue(fast.passed, fast.removed)

    def test_losing_the_plain_row_is_still_damage(self):
        # The guard: the fix must not make every tracked-deletion document
        # verify.  Here the output kept the *deleted* row and lost the plain
        # one, which is the damage the pairing key exists to keep visible.
        deleted = db.deleted_row(db.text_para(REQUIREMENT))
        plain = db.row(db.text_para(REQUIREMENT))

        fast, slow = self._both_paths(
            [db.text_para(HEADING), db.table_of(deleted, plain)],
            [db.text_para(HEADING), db.table_of(deleted)],
            strip_revisions=True,
        )

        self._same(fast, slow)
        self.assertFalse(fast.passed)

    def test_numbering_notices_agree(self):
        numbered = ('<w:pPr><w:numPr><w:ilvl w:val="0"/>'
                    '<w:numId w:val="1"/></w:numPr></w:pPr>')
        source = [db.text_para(HEADING),
                  "<w:p>" + numbered + db.run(NOTE) + "</w:p>",
                  "<w:p>" + numbered + db.run(REQUIREMENT) + "</w:p>"]
        output = [db.text_para(HEADING),
                  "<w:p>" + numbered + db.run(REQUIREMENT) + "</w:p>"]

        fast, slow = self._both_paths(source, output)

        self._same(fast, slow)
        self.assertEqual(len(fast.numbering), 1)


if __name__ == "__main__":
    unittest.main()
