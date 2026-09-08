"""Batch planning: destinations must be distinct, and must not be inputs.

These run everywhere.  The rules they cover used to live in ``gui.py``, where
no Linux run could reach them.
"""

import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from batch import (
    BatchItem,
    FileOutcome,
    output_for,
    plan_batch,
    same_file,
    summarise,
)


class OutputPathTests(unittest.TestCase):

    def test_output_lands_beside_the_input_by_default(self):
        self.assertEqual(
            output_for(Path("/specs/230500 Fire Suppression.docx")),
            Path("/specs/230500 Fire Suppression_cleaned.docx"),
        )

    def test_output_lands_in_the_chosen_folder(self):
        self.assertEqual(
            output_for(Path("/specs/a.docx"), Path("/out")),
            Path("/out/a_cleaned.docx"),
        )

    def test_a_file_is_never_its_own_destination(self):
        already = Path("/specs/a_cleaned.docx")
        self.assertNotEqual(output_for(already), already)


class PlanBatchTests(unittest.TestCase):

    def test_same_basename_from_two_folders_is_rejected(self):
        plan = plan_batch(
            [Path("/proj/A/230500.docx"), Path("/proj/B/230500.docx")],
            Path("/out"),
        )

        self.assertFalse(plan.ok)
        self.assertEqual(len(plan.conflicts), 1)
        conflict = plan.conflicts[0]
        self.assertEqual(conflict.kind, "shared_destination")
        self.assertEqual(len(conflict.sources), 2)
        self.assertIn("230500.docx", conflict.describe())
        self.assertIn("230500_cleaned.docx", conflict.describe())

    def test_three_way_collision_is_reported_once_with_every_source(self):
        plan = plan_batch(
            [Path(f"/proj/{folder}/s.docx") for folder in ("A", "B", "C")],
            Path("/out"),
        )

        self.assertEqual(len(plan.conflicts), 1)
        self.assertEqual(len(plan.conflicts[0].sources), 3)

    def test_distinct_names_are_accepted(self):
        files = [Path("/proj/A/a.docx"), Path("/proj/B/b.docx")]

        plan = plan_batch(files, Path("/out"))

        self.assertTrue(plan.ok)
        self.assertEqual(plan.items, [
            BatchItem(files[0], Path("/out/a_cleaned.docx")),
            BatchItem(files[1], Path("/out/b_cleaned.docx")),
        ])

    def test_same_basename_in_place_does_not_collide(self):
        # No common destination folder, so the two outputs stay apart.
        plan = plan_batch([Path("/proj/A/s.docx"), Path("/proj/B/s.docx")], None)

        self.assertTrue(plan.ok)

    def test_case_only_collision_is_rejected_where_case_does_not_distinguish(self):
        # normcase folds case on Windows and is identity on POSIX, so this
        # patch runs the Windows rule on any host.  Two files that differ only
        # in case are one file there, and one of the cleans would be lost.
        with patch("os.path.normcase", str.lower):
            plan = plan_batch(
                [Path("/proj/A/Spec.docx"), Path("/proj/B/SPEC.docx")],
                Path("/out"),
            )

        self.assertFalse(plan.ok)
        self.assertEqual(plan.conflicts[0].kind, "shared_destination")

    def test_case_only_names_are_distinct_where_case_distinguishes(self):
        # The same two files on a case-sensitive filesystem are two files.
        with patch("os.path.normcase", lambda value: value):
            plan = plan_batch(
                [Path("/proj/A/Spec.docx"), Path("/proj/B/SPEC.docx")],
                Path("/out"),
            )

        self.assertTrue(plan.ok)

    def test_destination_that_is_another_selected_input_is_rejected(self):
        # Cleaning foo.docx writes foo_cleaned.docx, which is also selected.
        # Processing is sequential, so it would be destroyed before its turn.
        plan = plan_batch(
            [Path("/x/foo.docx"), Path("/x/foo_cleaned.docx")], None
        )

        self.assertFalse(plan.ok)
        conflict = next(c for c in plan.conflicts if c.kind == "destination_is_input")
        self.assertEqual(conflict.sources, [Path("/x/foo_cleaned.docx")])
        self.assertIn("overwrite a document waiting to be processed",
                      conflict.describe())

    def test_a_lone_already_cleaned_file_is_fine(self):
        plan = plan_batch([Path("/x/foo_cleaned.docx")], None)

        self.assertTrue(plan.ok)


class ExistingPathTests(unittest.TestCase):
    """Cases that need real files: equivalent spellings and existing outputs."""

    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix="speccleanse_batch_"))
        self.addCleanup(shutil.rmtree, self.dir, True)

    def _touch(self, name: str) -> Path:
        path = self.dir / name
        path.write_bytes(b"")
        return path

    def test_equivalent_spelling_of_an_input_is_still_that_input(self):
        self._touch("foo.docx")
        equivalent = self._touch("foo_cleaned.docx")
        spelled_differently = self.dir / "sub" / ".." / "foo_cleaned.docx"
        (self.dir / "sub").mkdir()

        plan = plan_batch([self.dir / "foo.docx", spelled_differently], None)

        self.assertFalse(plan.ok)
        self.assertTrue(any(c.kind == "destination_is_input" for c in plan.conflicts))
        self.assertTrue(equivalent.exists())

    def test_existing_outputs_are_listed_for_confirmation(self):
        source = self._touch("a.docx")
        stale = self._touch("a_cleaned.docx")

        plan = plan_batch([source], None)

        self.assertTrue(plan.ok)
        self.assertEqual(plan.existing_outputs, [stale])

    def test_a_rejected_batch_does_not_ask_about_overwriting(self):
        # Collision first: there is nothing to confirm about a run that will
        # not happen, and asking would imply it was going to.
        self._touch("s.docx")
        (self.dir / "B").mkdir()
        (self.dir / "B" / "s.docx").write_bytes(b"")
        (self.dir / "s_cleaned.docx").write_bytes(b"")

        plan = plan_batch(
            [self.dir / "s.docx", self.dir / "B" / "s.docx"], self.dir
        )

        self.assertFalse(plan.ok)
        self.assertEqual(plan.existing_outputs, [])

    def test_same_file_sees_through_a_symlink(self):
        real = self._touch("real.docx")
        link = self.dir / "link.docx"
        try:
            link.symlink_to(real)
        except (OSError, NotImplementedError):
            self.skipTest("symlinks unavailable on this host")

        self.assertTrue(same_file(link, real))
        self.assertFalse(same_file(real, self.dir / "other.docx"))

    def test_same_file_is_false_for_two_missing_paths(self):
        self.assertFalse(same_file(self.dir / "no.docx", self.dir / "nope.docx"))


class SummariseTests(unittest.TestCase):

    def test_every_outcome_that_occurred_is_named(self):
        self.assertEqual(
            summarise({
                FileOutcome.VERIFIED: 8,
                FileOutcome.NEEDS_REVIEW: 2,
                FileOutcome.FAILED: 1,
            }),
            "Done: 8 verified, 2 need review, 1 failed",
        )

    def test_zero_counts_are_left_out(self):
        self.assertEqual(
            summarise({FileOutcome.VERIFIED: 3, FileOutcome.FAILED: 0}),
            "Done: 3 verified",
        )

    def test_one_file_needing_review_reads_correctly(self):
        self.assertEqual(
            summarise({FileOutcome.NEEDS_REVIEW: 1}), "Done: 1 needs review"
        )

    def test_an_empty_run_says_so(self):
        self.assertEqual(
            summarise({o: 0 for o in FileOutcome}), "Done: no files processed"
        )

    def test_needs_review_is_never_folded_into_a_success_total(self):
        line = summarise({FileOutcome.VERIFIED: 1, FileOutcome.NEEDS_REVIEW: 1})

        self.assertIn("1 verified", line)
        self.assertIn("1 needs review", line)
        self.assertNotIn("succeeded", line)
        self.assertNotIn("2", line)


class FileOutcomeTests(unittest.TestCase):

    def test_the_three_outcomes_are_distinct(self):
        self.assertEqual(len({o for o in FileOutcome}), 3)

    def test_no_outcome_is_falsy(self):
        # The point of the enum is that a caller cannot test it for truth and
        # accidentally count a needs-review file as a success.
        for outcome in FileOutcome:
            self.assertTrue(outcome)


if __name__ == "__main__":
    unittest.main()
