"""Planning a clean run: which input writes to which output, and is that safe.

These rules live outside ``gui.py`` deliberately.  That module imports tkinter
at the top, so anything defined there cannot be exercised where Tk is absent —
which is every Linux test run and every Linux CI job.  Deciding whether a batch
is safe to write is exactly the kind of rule that has to be tested, so it is
kept where a test can reach it.

The rule the planner exists to enforce: two selected documents that share a
basename map to one destination when a common output folder is chosen, and the
second clean silently overwrites the first.  Nothing on disk shows the loss,
because the surviving file is a perfectly valid cleaned document — of the wrong
source.
"""

import os
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path


class FileOutcome(Enum):
    """What became of one file in a run.

    A successful write and a passing verification are separate facts.  One
    Boolean cannot carry both, which is how a file whose verification reported
    a preserve violation used to be counted in the "succeeded" total.
    """

    VERIFIED = "verified"
    NEEDS_REVIEW = "needs review"
    FAILED = "failed"


#: Appended to an input's stem to name its cleaned output.
OUTPUT_SUFFIX = "_cleaned"


def output_for(input_path: Path, output_dir: Path | None = None) -> Path:
    """The cleaned-file destination for one input."""
    parent = output_dir if output_dir else input_path.parent
    return parent / (input_path.stem + OUTPUT_SUFFIX + ".docx")


def _key(path: Path) -> str:
    """A comparison key under which two spellings of one file are equal.

    ``normcase`` folds case and separators on Windows and does nothing on
    POSIX.  ``realpath`` resolves symlinks, ``.`` and ``..`` segments, and
    relative spellings; for a path that does not exist yet it still normalises
    the part of it that does.
    """
    return os.path.normcase(os.path.realpath(path))


def same_file(left: Path, right: Path) -> bool:
    """True if two paths denote one file.

    String inequality does not prove two paths are different files: they can
    differ by case on Windows, by a symlink, or by a relative spelling.  Where
    both exist the filesystem is asked directly, which also catches hard links
    and junctions that ``realpath`` does not collapse.
    """
    if _key(left) == _key(right):
        return True
    try:
        return os.path.samefile(left, right)
    except OSError:
        # One of them does not exist, or cannot be stat'd.  The key comparison
        # above is then the best available answer, and it said no.
        return False


@dataclass(frozen=True)
class BatchItem:
    """One input and the output it will be written to."""

    source: Path
    destination: Path


@dataclass(frozen=True)
class BatchConflict:
    """A reason the batch must not run, with the paths that caused it."""

    #: ``shared_destination`` — several inputs map to one output.
    #: ``destination_is_input`` — an output would overwrite a selected input.
    kind: str
    destination: Path
    sources: list[Path]

    def describe(self) -> str:
        """A message naming the sources and the destination they collide on."""
        listed = "\n".join(f"    {source}" for source in self.sources)
        if self.kind == "shared_destination":
            return (
                f"These {len(self.sources)} selected files would all be written "
                f"to one destination:\n{listed}\n  destination: {self.destination}"
            )
        return (
            f"This destination is itself a selected input, so cleaning would "
            f"overwrite a document waiting to be processed:\n{listed}\n"
            f"  destination: {self.destination}"
        )


@dataclass
class BatchPlan:
    """The validated set of input/output pairs for one run."""

    items: list[BatchItem] = field(default_factory=list)
    conflicts: list[BatchConflict] = field(default_factory=list)
    #: Destinations that already exist, for the overwrite confirmation.  Only
    #: meaningful once ``ok`` is true: a batch that collides with itself is
    #: rejected before anyone is asked about replacing an earlier run's output.
    existing_outputs: list[Path] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        """True if every destination is distinct and none is a selected input."""
        return not self.conflicts


def plan_batch(files: list[Path], output_dir: Path | None = None) -> BatchPlan:
    """Map every selected input to its destination and check the whole set.

    Built once, on the caller's thread, before any file is opened.  The worker
    is handed the result rather than recomputing destinations later from state
    the user may have changed in the meantime.

    This is preflight protection, not a lock: another process can still write
    into the same folder while the run is under way.
    """
    items = [BatchItem(source=path, destination=output_for(path, output_dir))
             for path in files]

    conflicts: list[BatchConflict] = []

    # Several inputs sharing one destination.  Grouped by key rather than by
    # the Path itself so that two spellings of one destination still collide.
    by_destination: dict[str, list[BatchItem]] = {}
    for item in items:
        by_destination.setdefault(_key(item.destination), []).append(item)

    for group in by_destination.values():
        if len(group) > 1:
            conflicts.append(BatchConflict(
                kind="shared_destination",
                destination=group[0].destination,
                sources=[item.source for item in group],
            ))

    # A destination that is itself a selected input.  Its own source does not
    # count: an input named "x.docx" produces "x_cleaned.docx", and re-cleaning
    # a file already called "x_cleaned.docx" produces "x_cleaned_cleaned.docx",
    # so a file can never be its own destination.  Another *selected* file
    # landing on it is the real hazard — processing is sequential, so that
    # input would be overwritten before its turn came.
    for item in items:
        clashing = [
            other.source for other in items
            if other is not item and same_file(item.destination, other.source)
        ]
        if clashing:
            conflicts.append(BatchConflict(
                kind="destination_is_input",
                destination=item.destination,
                sources=clashing,
            ))

    plan = BatchPlan(items=items, conflicts=conflicts)
    if plan.ok:
        plan.existing_outputs = [
            item.destination for item in items if item.destination.exists()
        ]
    return plan
