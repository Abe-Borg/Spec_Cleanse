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


def summarise(counts: dict[FileOutcome, int]) -> str:
    """One line describing how a finished run turned out.

    Every outcome that actually occurred is named.  There is deliberately no
    "succeeded" total: that word used to cover both a verified file and one
    whose verification reported a preserve violation, which is the confusion
    :class:`FileOutcome` exists to end.
    """
    verified = counts.get(FileOutcome.VERIFIED, 0)
    review = counts.get(FileOutcome.NEEDS_REVIEW, 0)
    failed = counts.get(FileOutcome.FAILED, 0)

    parts = []
    if verified:
        parts.append(f"{verified} verified")
    if review:
        parts.append(f"{review} need{'s' if review == 1 else ''} review")
    if failed:
        parts.append(f"{failed} failed")

    return "Done: " + (", ".join(parts) if parts else "no files processed")


#: Appended to an input's stem to name its cleaned output.
OUTPUT_SUFFIX = "_cleaned"


def output_for(input_path: Path, output_dir: Path | None = None) -> Path:
    """The cleaned-file destination for one input."""
    parent = output_dir if output_dir else input_path.parent
    return parent / (input_path.stem + OUTPUT_SUFFIX + ".docx")


#: True where the platform's own convention ignores case — the fallback when a
#: volume cannot be probed.
_PLATFORM_IGNORES_CASE = os.path.normcase("A") == "a"


def volume_ignores_case(path: Path) -> bool:
    """Whether the volume holding ``path`` treats two spellings of a name as one.

    ``normcase`` answers for the *platform*, which is a different question. A
    default macOS APFS volume ignores case while ``posixpath.normcase`` is the
    identity, so two destinations differing only in case read as two files when
    they are one — and the second clean replaces the first, which is the exact
    loss this module exists to prevent. Windows volumes can be case-sensitive
    per directory, so the converse holds too.

    Probed read-only. The nearest existing ancestor has its own name respelled
    in the opposite case; if that still resolves, the volume ignores case.
    Where there is nothing to probe — no existing ancestor, or a name with no
    letters in it — the platform's convention is the fallback.

    A case-sensitive volume that genuinely holds both spellings answers
    "ignores case" here. That is wrong, and it is wrong in the safe direction:
    it can only cause a batch to be rejected, never a file to be overwritten.
    """
    for candidate in (path, *path.parents):
        if not candidate.exists():
            continue
        flipped = candidate.name.swapcase()
        if not flipped or flipped == candidate.name:
            continue  # a root, or a name with no letters to flip
        return candidate.with_name(flipped).exists()

    return _PLATFORM_IGNORES_CASE


def _key(path: Path, fold_case: bool) -> str:
    """A comparison key under which two spellings of one file are equal.

    ``normcase`` folds separators on Windows (and case, there, already);
    ``realpath`` resolves symlinks, ``.`` and ``..`` segments, and relative
    spellings, and for a path that does not exist yet it still normalises the
    part of it that does. ``fold_case`` supplies what ``normcase`` cannot: the
    answer for the volume this path is actually on.
    """
    normalised = os.path.normcase(os.path.realpath(path))
    return normalised.lower() if fold_case else normalised


class _Keyer:
    """Builds comparison keys, remembering each directory's case semantics.

    Probing touches the filesystem, and a batch asks about the same handful of
    directories repeatedly.
    """

    def __init__(self) -> None:
        self._folds: dict[Path, bool] = {}

    def folds_case(self, path: Path) -> bool:
        parent = path.parent
        if parent not in self._folds:
            self._folds[parent] = volume_ignores_case(parent)
        return self._folds[parent]

    def __call__(self, path: Path) -> str:
        return _key(path, self.folds_case(path))


def same_file(left: Path, right: Path, keyer: "_Keyer | None" = None) -> bool:
    """True if two paths denote one file.

    String inequality does not prove two paths are different files: they can
    differ by case, by a symlink, or by a relative spelling. Where both exist
    the filesystem is asked directly, which also catches hard links and
    junctions that ``realpath`` does not collapse. Case is folded when *either*
    side sits on a volume that ignores it, since one of the two may not exist
    yet and so cannot be probed on its own.
    """
    keyer = keyer or _Keyer()
    fold = keyer.folds_case(left) or keyer.folds_case(right)
    if _key(left, fold) == _key(right, fold):
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
    keyer = _Keyer()

    # Several inputs sharing one destination.  Grouped by key rather than by
    # the Path itself so that two spellings of one destination still collide.
    by_destination: dict[str, list[BatchItem]] = {}
    for item in items:
        by_destination.setdefault(keyer(item.destination), []).append(item)

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
            if other is not item and same_file(item.destination, other.source, keyer)
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
