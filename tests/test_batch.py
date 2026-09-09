"""Batch planning: destinations must be distinct, and must not be inputs.

These run everywhere.  The rules they cover used to live in ``gui.py``, where
no Linux run could reach them.
"""

import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import batch
from batch import (
    BatchItem,
    FileOutcome,
    FileReport,
    ReviewCategory,
    run_batch,
    output_for,
    plan_batch,
    same_file,
    summarise,
    volume_ignores_case,
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

    # The volume decides whether case matters, so these read the same on every
    # host.  Patching os.path.normcase would not: it is no longer a case knob,
    # and asserting anything about it fails wherever the platform disagrees.

    def test_case_only_collision_is_rejected_on_a_case_insensitive_volume(self):
        # Two files differing only in case are one file there, and one of the
        # cleans would be lost.  A default macOS APFS volume is this case, and
        # posixpath.normcase cannot see it — which is why the volume is asked.
        with patch.object(batch, "volume_ignores_case", return_value=True):
            plan = plan_batch(
                [Path("/proj/A/Spec.docx"), Path("/proj/B/SPEC.docx")],
                Path("/out"),
            )

        self.assertFalse(plan.ok)
        self.assertEqual(plan.conflicts[0].kind, "shared_destination")

    def test_the_same_pair_is_allowed_on_a_case_sensitive_volume(self):
        with patch.object(batch, "volume_ignores_case", return_value=False):
            plan = plan_batch(
                [Path("/proj/A/Spec.docx"), Path("/proj/B/SPEC.docx")],
                Path("/out"),
            )

        self.assertTrue(plan.ok)

    def test_a_destination_that_is_an_input_in_another_case_is_rejected(self):
        with patch.object(batch, "volume_ignores_case", return_value=True):
            plan = plan_batch(
                [Path("/x/foo.docx"), Path("/x/FOO_CLEANED.docx")], None
            )

        self.assertFalse(plan.ok)
        self.assertTrue(
            any(c.kind == "destination_is_input" for c in plan.conflicts)
        )

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

    def test_the_volume_probe_answers_for_a_real_directory(self):
        # Whatever this host's filesystem does, the probe must answer without
        # raising and without writing anything into the directory.
        before = sorted(entry.name for entry in self.dir.iterdir())

        answer = volume_ignores_case(self.dir)

        self.assertIsInstance(answer, bool)
        self.assertEqual(sorted(e.name for e in self.dir.iterdir()), before)

    def test_the_volume_probe_answers_for_a_directory_yet_to_be_created(self):
        # The output folder need not exist when the batch is planned; the
        # nearest existing ancestor decides.
        self.assertIsInstance(
            volume_ignores_case(self.dir / "not" / "created" / "yet"), bool
        )

    def test_the_volume_probe_falls_back_to_the_platform_when_nothing_exists(self):
        # Nothing to probe, so the platform's convention is the answer.
        answer = volume_ignores_case(Path("/nonexistent-root-xyz/deep/path"))

        self.assertEqual(answer, batch._PLATFORM_IGNORES_CASE)


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


class RunBatchTests(unittest.TestCase):
    """Driving a validated plan, where no window is needed to watch it.

    This loop used to live in ``gui.py``, so "the batch kept going after one
    file failed" was a claim no Linux run could check.
    """

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_runbatch_"))
        self.addCleanup(shutil.rmtree, self.temp_dir, True)
        self.lines: list[str] = []

    def item(self, name: str) -> BatchItem:
        source = self.temp_dir / f"{name}.docx"
        source.write_bytes(b"not really a docx")
        return BatchItem(source=source, destination=self.temp_dir / f"{name}_cleaned.docx")

    @property
    def output(self) -> str:
        return "\n".join(self.lines)

    def test_counts_and_categories_are_tallied(self):
        reports = {
            "a": FileReport(FileOutcome.VERIFIED, output_written=True),
            "b": FileReport(
                FileOutcome.NEEDS_REVIEW,
                frozenset({ReviewCategory.DETECTED_DAMAGE}),
                output_written=True,
            ),
            "c": FileReport(FileOutcome.FAILED),
        }
        items = [self.item(name) for name in reports]

        tally = run_batch(items, lambda i: reports[i.source.stem], self.lines.append)

        self.assertEqual(tally.counts, {
            FileOutcome.VERIFIED: 1,
            FileOutcome.NEEDS_REVIEW: 1,
            FileOutcome.FAILED: 1,
        })
        self.assertEqual(tally.categories, {ReviewCategory.DETECTED_DAMAGE: 1})
        self.assertEqual(
            tally.summary(),
            "Done: 1 verified, 1 needs review (1 detected damage), 1 failed",
        )

    def test_a_file_in_two_categories_is_counted_in_both(self):
        item = self.item("a")
        report = FileReport(
            FileOutcome.NEEDS_REVIEW,
            frozenset({ReviewCategory.DETECTED_DAMAGE, ReviewCategory.REFERENCE_NUMBERING}),
            output_written=True,
        )

        tally = run_batch([item], lambda _: report, self.lines.append)

        self.assertEqual(tally.categories, {
            ReviewCategory.DETECTED_DAMAGE: 1,
            ReviewCategory.REFERENCE_NUMBERING: 1,
        })

    def test_an_exception_fails_one_file_and_the_batch_continues(self):
        # A single bad file used to end the run, leaving every remaining file
        # unprocessed with nothing in the log to say why.
        items = [self.item(name) for name in ("a", "b", "c")]

        def clean(item: BatchItem) -> FileReport:
            if item.source.stem == "b":
                raise RuntimeError("unreadable")
            return FileReport(FileOutcome.VERIFIED, output_written=True)

        tally = run_batch(items, clean, self.lines.append)

        self.assertEqual(
            tally.counts, {FileOutcome.VERIFIED: 2, FileOutcome.FAILED: 1}
        )
        self.assertIn("unreadable", self.output)
        self.assertIn("[3/3] c.docx", self.output)

    def test_a_failure_that_wrote_nothing_is_not_counted_as_unverified(self):
        tally = run_batch(
            [self.item("a")],
            lambda _: (_ for _ in ()).throw(RuntimeError("boom")),
            self.lines.append,
        )

        self.assertEqual(tally.unverified_outputs, 0)
        self.assertEqual(tally.summary(), "Done: 1 failed")

    def test_only_the_validated_manifest_is_ever_cleaned(self):
        # The selection and output folder are widgets the user can change
        # while the run is under way.  A destination worked out mid-run could
        # collide with one already written — the loss plan_batch exists to
        # prevent, reintroduced after its check had passed.
        items = [self.item("a"), self.item("b")]
        seen: list[BatchItem] = []

        def clean(item: BatchItem) -> FileReport:
            seen.append(item)
            # Whatever the caller does to its own state afterwards, the
            # manifest this loop is walking cannot change.
            items.append(self.item("c"))
            return FileReport(FileOutcome.VERIFIED, output_written=True)

        tally = run_batch(list(items), clean, self.lines.append)

        self.assertEqual([i.source.stem for i in seen], ["a", "b"])
        self.assertEqual(tally.counts, {FileOutcome.VERIFIED: 2})

    def test_the_destination_is_named_only_when_something_was_written(self):
        items = [self.item("a"), self.item("b")]
        reports = {
            "a": FileReport(FileOutcome.VERIFIED, output_written=True),
            "b": FileReport(FileOutcome.FAILED),
        }

        run_batch(items, lambda i: reports[i.source.stem], self.lines.append)

        self.assertIn("-> a_cleaned.docx", self.output)
        self.assertNotIn("-> b_cleaned.docx", self.output)

    def test_an_empty_plan_runs_nothing_and_says_so(self):
        tally = run_batch([], lambda _: self.fail("should not be called"), self.lines.append)

        self.assertEqual(tally.summary(), "Done: no files processed")
        self.assertEqual(self.lines, [])

    def test_a_pre_existing_destination_is_not_called_this_run_s_output(self):
        # No result to ask when the callback raises, so the destination is
        # checked before the call as well as after.  A file that was already
        # there is an earlier run's, and calling it unverified would invite the
        # user to delete a good document.
        item = self.item("a")
        item.destination.write_bytes(b"a perfectly good earlier output")

        tally = run_batch(
            [item],
            lambda _: (_ for _ in ()).throw(RuntimeError("boom")),
            self.lines.append,
        )

        self.assertEqual(tally.unverified_outputs, 0)
        self.assertNotIn("UNVERIFIED", self.output)
        self.assertIn("was already there before this run", self.output)
        self.assertEqual(tally.summary(), "Done: 1 failed")

    def test_a_file_that_appeared_during_a_failed_run_is_reported(self):
        # The guard against the case above being satisfied by never reporting
        # an unverified output at all.
        item = self.item("a")

        def clean(i: BatchItem) -> FileReport:
            i.destination.write_bytes(b"half a document")
            raise RuntimeError("boom")

        tally = run_batch([item], clean, self.lines.append)

        self.assertEqual(tally.unverified_outputs, 1)
        self.assertIn("UNVERIFIED", self.output)
        self.assertIn(str(item.destination), self.output)
        self.assertEqual(tally.summary(), "Done: 1 failed (1 wrote an unverified file)")

