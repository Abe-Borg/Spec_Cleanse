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
from collections.abc import Callable
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


class ReviewCategory(Enum):
    """Why a file needs review — the four kinds, never pooled.

    §10.6 criterion 3 requires these measured separately, and carries one
    non-gating requirement into this package: the same split has to reach the
    user, not only the acceptance measurement.  The reason is specific.  A
    Needs-review rate dominated by ``AMBIGUOUS_ALIGNMENT`` would say the
    verifier cannot follow its own reasoning on real documents, which is a
    reason to revisit the comparison; the same rate dominated by
    ``DETECTED_DAMAGE`` would say the cleaner is losing content, which is a
    different problem with a different fix.  One pooled number cannot tell
    those apart, and a verdict that cannot be acted on will be ignored.

    They are deliberately not ordered by severity.  A file can be in several
    at once, and picking one to show would be the pooling this exists to end.
    """

    #: The verifier could not establish *which* source paragraph a difference
    #: belongs to, or could not reason about the paragraph's offsets at all.
    #: Something is unexplained; where it came from is a guess.
    AMBIGUOUS_ALIGNMENT = "ambiguous alignment"

    #: Content is gone, or content the output invented is present, and the
    #: alignment that says so was exact.  A claim about the document.
    DETECTED_DAMAGE = "detected damage"

    #: The rules that produced this file are not the shipped defaults in a way
    #: worth knowing about.  Nothing is wrong with the output as such.
    CONFIGURATION = "configuration notice"

    #: A cross-reference this run broke, or numbering a removal may have
    #: shifted.  Nothing was lost; what a reader sees may differ.
    REFERENCE_NUMBERING = "reference/numbering warning"


def summarise(
    counts: dict[FileOutcome, int],
    categories: dict[ReviewCategory, int] | None = None,
    unverified_outputs: int = 0,
) -> str:
    """One line describing how a finished run turned out.

    Every outcome that actually occurred is named.  There is deliberately no
    "succeeded" total: that word used to cover both a verified file and one
    whose verification reported a preserve violation, which is the confusion
    :class:`FileOutcome` exists to end.

    ``categories`` names why the Needs-review files need it.  The counts are
    per *category*, not per file — a file in two categories is counted in
    both — so they can exceed the Needs-review total.  That is the honest
    shape: forcing one category per file would mean choosing which of two
    real concerns to hide.

    ``unverified_outputs`` is how many failures nonetheless left a file on
    disk.  §14.1 asks the disposition to be explicit either way: "failed" on
    its own leaves a reader unable to tell whether there is something in the
    output folder to delete.
    """
    verified = counts.get(FileOutcome.VERIFIED, 0)
    review = counts.get(FileOutcome.NEEDS_REVIEW, 0)
    failed = counts.get(FileOutcome.FAILED, 0)

    parts = []
    if verified:
        parts.append(f"{verified} verified")
    if review:
        detail = describe_categories(categories or {})
        parts.append(
            f"{review} need{'s' if review == 1 else ''} review"
            + (f" ({detail})" if detail else "")
        )
    if failed:
        parts.append(
            f"{failed} failed"
            + (f" ({unverified_outputs} wrote an unverified file)"
               if unverified_outputs else "")
        )

    return "Done: " + (", ".join(parts) if parts else "no files processed")


def describe_categories(categories: dict[ReviewCategory, int]) -> str:
    """The Needs-review reasons, in a fixed order, naming only what occurred.

    Declaration order, not frequency: a summary whose wording reshuffles
    between runs is harder to read than one that does not, and there is no
    severity ranking to sort by.
    """
    return ", ".join(
        f"{categories[category]} {category.value}"
        for category in ReviewCategory
        if categories.get(category)
    )


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

    ``realpath`` resolves symlinks, ``.`` and ``..`` segments and relative
    spellings, normalises separators, and for a path that does not exist yet
    still normalises the part of it that does.

    Case is folded on ``fold_case`` alone. ``normcase`` is deliberately not
    used here: it would fold case a second time, from the platform's
    convention, and the two can disagree — a case-sensitive directory on
    Windows would still have been folded, and no volume answer could have
    stopped it. One rule, and the volume states it (see
    :func:`volume_ignores_case`, whose fallback is the platform's convention
    when there is nothing to probe).
    """
    resolved = os.path.realpath(path)
    return resolved.lower() if fold_case else resolved


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


# ---------------------------------------------------------------------------
# Running a validated plan
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class FileReport:
    """What became of one file, and why.

    The outcome alone cannot be acted on.  "Needs review" says a file is not
    ready to hand on; it does not say whether to go and read the document, fix
    the configuration, or check a cross-reference — three different pieces of
    work.  ``categories`` is what makes the verdict actionable, and it is a
    set because a file can genuinely be in more than one.

    ``output_written`` is separate from the outcome because a failure can
    still leave a file on disk: verification raising after a successful write
    is exactly that case, and §14.1 requires the disposition stated rather
    than left for the reader to guess.
    """

    outcome: FileOutcome
    categories: frozenset[ReviewCategory] = frozenset()
    output_written: bool = False


@dataclass
class RunTally:
    """Running counts for a batch, and the line that describes it."""

    counts: dict[FileOutcome, int] = field(default_factory=dict)
    categories: dict[ReviewCategory, int] = field(default_factory=dict)
    #: Files that failed but left something on disk anyway.
    unverified_outputs: int = 0

    def add(self, report: FileReport) -> None:
        self.counts[report.outcome] = self.counts.get(report.outcome, 0) + 1
        for category in report.categories:
            self.categories[category] = self.categories.get(category, 0) + 1
        if report.outcome is FileOutcome.FAILED and report.output_written:
            self.unverified_outputs += 1

    def summary(self) -> str:
        return summarise(self.counts, self.categories, self.unverified_outputs)


def run_batch(
    items: list[BatchItem],
    clean: Callable[[BatchItem], FileReport],
    log: Callable[[str], None],
    announce: Callable[[int, int, BatchItem], None] | None = None,
) -> RunTally:
    """Clean every item in a validated plan, and tally what happened.

    The loop lives here rather than in ``gui.py`` for the reason this module
    exists: what it does has to be testable where Tk is absent, and "the batch
    kept going after one file failed" is not a claim worth making untested.

    ``items`` is the manifest :func:`plan_batch` validated before the worker
    started, and destinations are never recomputed from anything live.  The
    file selection and the output folder are widgets the user can change while
    the run is under way; a destination worked out mid-run could collide with
    one already written, which is the loss :func:`plan_batch` exists to
    prevent and would have been reintroduced after the check had passed.

    One file's unexpected failure is that file's, not the run's.  ``process()``
    already turns its own exceptions into errors, so this guard is for
    everything around it — without it a single bad file ended the batch, and
    every remaining file went unprocessed with no record of why.

    When the callback raises there is no result to ask, so the destination is
    checked *before* the call as well as after.  Only a file that appeared is
    this run's; one that was already there is an earlier run's output and is
    reported as such.  Inferring a write from the file merely existing
    afterwards would tell the user their good document is unverified, which is
    an invitation to delete it.
    """
    tally = RunTally()
    total = len(items)

    for index, item in enumerate(items, 1):
        if announce is not None:
            announce(index, total, item)
        log(f"[{index}/{total}] {item.source.name}")

        existed = item.destination.exists()
        try:
            report = clean(item)
        except Exception as exc:
            # Nothing below the callback is trusted to have reported this.
            appeared = not existed and item.destination.exists()
            report = FileReport(FileOutcome.FAILED, output_written=appeared)
            log(f"  ERROR: {item.source.name} could not be processed: {exc}")
            if appeared:
                log(f"  A file was left behind and is UNVERIFIED: {item.destination}")
            elif existed:
                # Deliberately not counted as this run's output.  Whether the
                # run got as far as overwriting it is unknowable from here, and
                # the safe reading is the one that does not call an existing
                # document unverified.
                log(f"  {item.destination.name} was already there before this run;"
                    " it may be an earlier output and was not checked.")

        tally.add(report)
        if report.output_written and report.outcome is not FileOutcome.FAILED:
            log(f"  -> {item.destination.name}")
        log("")

    return tally
