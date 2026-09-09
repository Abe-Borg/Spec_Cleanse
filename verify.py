"""
SpecCleanse Verification Module

Compares an input DOCX with its cleaned output and reports what the clean
actually did to the document.  Three questions are asked:

  1. Which paragraphs disappeared, and does each one match a rule that was
     meant to remove it?
  2. Which surviving paragraphs lost text, and was the lost fragment
     something a rule asked for?  (Inline placeholder redaction lands here.)
  3. Is the output still a document Word will open?

Question 1 shares its compiled patterns with the detection engine, so it
agrees with the cleaner by construction instead of through a second, drifting
copy of the same regexes.  That makes it a consistency check, not an
independent one: a rule that removes the wrong thing is reported as expected.
The two genuinely independent checks are the formatting signals read from the
source DOCX and the structural inspection, which is the only layer that can
see damage no pattern describes — an emptied footer, a lost section break, a
field left half-open.
"""

import difflib
import shutil
import tempfile
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from functools import cached_property
from pathlib import Path

from lxml import etree

from apppaths import resolve_config_path
from batch import ReviewCategory
from detection import DetectionEngine, ParagraphEvidence
from docx_xml import (
    P_TAG,
    StyleIndex,
    TBL_TAG,
    TC_TAG,
    W,
    block_children,
    collect_content_parts,
    bookmark_names,
    field_chars_balanced,
    field_instructions,
    in_tracked_deletion,
    iter_own_runs,
    iter_paragraphs,
    note_identity,
    numbering_id,
    paragraph_signature,
    reference_consumers,
    run_profile,
    run_text,
    load_config,
    load_styles,
    orphaned_range_markers,
    paragraph_text,
    parse_xml,
)

#: Containers whose emptiness makes Word call a file unreadable.  Inline
#: ``w:sdtContent`` legitimately holds runs rather than blocks, so it is not
#: inspected here even though paragraph removal treats it as a block container.
LINTED_CONTAINERS = (
    f"{W}body",
    f"{W}hdr",
    f"{W}ftr",
    f"{W}footnote",
    f"{W}endnote",
    f"{W}txbxContent",
    TC_TAG,
)

#: A paired paragraph must still resemble the one it came from; below this
#: similarity the "modification" is really a removal that happened to share a
#: few characters with an unrelated survivor.
MIN_PAIR_SIMILARITY = 0.5

PRESERVE_VIOLATION = "preserve_violation"
FORMATTING_BASED = "formatting_based"
TRACKED_DELETION = "tracked_deletion"


# =============================================================================
# Data structures
# =============================================================================

@dataclass
class ParagraphInfo:
    """A source paragraph, and what the configured policy permits losing from it.

    ``raw_text`` is the paragraph as the document holds it; ``text`` is the
    trimmed form the comparison aligns on.  Both are kept deliberately: offsets
    belong to the raw text, and an offset computed against a trimmed copy
    addresses the wrong characters.
    """
    raw_text: str
    evidence: ParagraphEvidence
    #: Signatures of this paragraph's text-carrying runs, as document fact.
    #: Read from both sides, so the output can be compared run-for-run when
    #: its extracted text is identical to more than one source arrangement.
    run_profile: tuple = ()
    #: The paragraph's full identity — its style as well as its runs.  Used to
    #: decide *which* of several identical texts actually disappeared, where the
    #: style is authority the runs do not carry.
    signature: tuple = ()
    #: Package-relative part name, e.g. ``word/document.xml``.  Full, so that
    #: ``word/document.xml`` and ``word/glossary/document.xml`` stay distinct.
    part: str = ""
    #: Footnote or endnote identity within a shared part, or None in body text.
    story: str | None = None
    #: The automatic-numbering list this paragraph belongs to, if any.  Read
    #: once during extraction rather than re-derived per removal.
    numbering: str | None = None
    #: True if a tracked change marks this paragraph's container as deleted —
    #: a deleted table row keeps its text in plain w:t, so nothing else shows it.
    in_tracked_deletion: bool = False

    @property
    def location(self) -> tuple[str, str | None]:
        """What this paragraph is compared *within*.

        Comparison happens inside a location, never across them.  A header and
        the body are different documents as far as content identity goes, and an
        identical heading in one cannot account for the other's loss.
        """
        return (self.part, self.story)

    @cached_property
    def text(self) -> str:
        """The trimmed text the comparison aligns on.

        Cached because ``raw_text`` never changes after construction and this
        was read 31 million times in one 4,000-paragraph verification — 8.7s of
        ``str.strip`` on identical input.
        """
        return self.raw_text.strip()

    @cached_property
    def lead(self) -> int:
        """Characters trimmed from the front, mapping ``text`` offsets to raw."""
        return len(self.raw_text) - len(self.raw_text.lstrip())

    @property
    def preserve_reason(self) -> str | None:
        """Why this paragraph is protected outright, pattern or style."""
        return self.evidence.preserve_reason

    def authority_for(self, start: int, end: int) -> tuple[str, str, bool] | None:
        """The rule authorizing loss of ``text[start:end]``, or None.

        Positional, not textual.  Asking whether a lost fragment resembles
        something removable cannot tell two identical occurrences apart, and a
        paragraph may well hold the same words in an editorial run and in a
        requirement.
        """
        return self.evidence.reason_for_span(start + self.lead, end + self.lead)

    def expected_after_removal(self) -> str | None:
        """What this paragraph becomes under the whole intended removal.

        Every authorized interval, not only the placeholders: a paragraph can
        carry both a run policy permits losing and a placeholder inside another
        run, and expecting only one of the two leaves the differ to guess.  The
        intervals come from evaluating the configured patterns against this
        source paragraph — never from anything the processor reports.
        """
        return self.evidence.expected_text()

    def authorized_intervals(self) -> list[tuple[int, int]]:
        """The authorized intervals in ``text`` coordinates, not raw ones.

        Used where the output *is* the expected text: the intervals are then
        known outright, and deriving them by diffing would reintroduce exactly
        the guess this path exists to avoid.
        """
        limit = len(self.text)
        intervals = []
        for start, end in self.evidence.authorized_spans():
            lo, hi = max(0, start - self.lead), min(limit, end - self.lead)
            if lo < hi:
                intervals.append((lo, hi))
        return intervals


@dataclass
class RemovedParagraph:
    """A paragraph present in input but absent in output."""
    text: str
    category: str | None = None       # e.g. "specifier_note", or None if unexpected
    pattern_matched: str | None = None # the regex or signal that matched
    #: True when the paragraph's own offsets could not be trusted, so no
    #: interval question about it could be answered and everything it lost
    #: reads as unexplained.  That is the verifier declining to reason, not a
    #: finding about the document — see :class:`ReviewCategory`.
    ambiguous: bool = False


@dataclass
class ModifiedParagraph:
    """A paragraph that survived but lost some of its text."""
    before: str
    after: str
    fragments: list[str] = field(default_factory=list)
    category: str | None = None
    pattern_matched: str | None = None
    #: True when this paragraph was paired with its survivor by similarity
    #: rather than by an exact answer.  The fragments such a pairing reports
    #: are alignment artefacts — a real case produced ``'nd hangers a'`` — so
    #: reporting them as content someone lost overstates what is known.
    ambiguous: bool = False

    @property
    def text(self) -> str:
        """The text that went missing, for reporting alongside removals."""
        return " / ".join(self.fragments)


@dataclass
class StructuralViolation:
    """Damage to the document's structure that the output has and the input did not."""
    issue: str
    count: int = 1
    #: ``"structure"`` for a shape the package should not have, ``"reference"``
    #: for a cross-reference this run broke.  A field, not a substring of
    #: ``issue``: deciding a category by matching the sentence shown to the
    #: user makes the wording load-bearing, and the wording is meant to be
    #: free to change.
    kind: str = "structure"

    def __str__(self) -> str:
        return f"{self.issue}" + (f" (x{self.count})" if self.count > 1 else "")


@dataclass
class NumberingNotice:
    """A removed paragraph that took part in automatic numbering.

    Deliberately not a structural violation and not an unexplained removal: no
    text integrity claim is being made.  What is being said is narrower — the
    numbers a reader sees, and any reference written against them, *may* now
    read differently.  Nothing here asserts that a reference broke; §13.1's
    check is what says that, and it says it about a specific named bookmark.
    """
    part: str
    numbering: str
    preview: str

    def __str__(self) -> str:
        return (
            f"{self.part}: removed a paragraph in numbering list {self.numbering} — "
            f"displayed numbering or references to it may change: \"{self.preview}\""
        )


@dataclass
class StructureReport:
    """What an inspection of one DOCX found."""
    issues: Counter = field(default_factory=Counter)
    section_breaks: int = 0
    #: Field carriers keyed by ``(part, instruction)``.  Per part, because a
    #: field lost from the body is not answered by an identical one in a
    #: header; by instruction rather than by a count, because a document-wide
    #: total hides one field going while another arrives.
    fields: Counter = field(default_factory=Counter)
    #: Bookmark names defined anywhere in the package, and the names something
    #: still points at.  Document-wide, not per part, because a reference in a
    #: header legitimately names a bookmark in the body — the opposite of the
    #: field inventory above, and for the opposite reason.
    bookmarks: set = field(default_factory=set)
    #: ``(part, consumer)`` for everything that points at a bookmark.  Kept per
    #: consumer rather than collapsed by target name: one missing bookmark can
    #: break references in the body, a header and a note at once, and each is a
    #: separate place someone has to go and repair.
    references: list = field(default_factory=list)

    def broken_references(self) -> set:
        """Folded names something points at that no bookmark defines."""
        return {
            consumer.folded for _, consumer in self.references
            if consumer.folded not in self.bookmarks
        }


@dataclass
class VerificationResult:
    """Structured result of input-vs-output comparison."""
    input_path: Path
    output_path: Path
    input_paragraph_count: int = 0
    output_paragraph_count: int = 0
    removed: list[RemovedParagraph] = field(default_factory=list)
    modified: list[ModifiedParagraph] = field(default_factory=list)
    structural: list[StructuralViolation] = field(default_factory=list)
    added: list[str] = field(default_factory=list)
    #: Its own category, counted apart from damage — §10.6 criterion 3.  These
    #: make a run need review without claiming anything was lost.
    numbering: list[NumberingNotice] = field(default_factory=list)

    @property
    def expected_removals(self) -> list[RemovedParagraph]:
        """Removals that matched a known bloat pattern or formatting signal."""
        return [
            r for r in self.removed
            if r.category is not None and r.category != PRESERVE_VIOLATION
        ]

    @property
    def unexpected_removals(self) -> list[RemovedParagraph]:
        """Removals that did NOT match any known pattern — red flags."""
        return [r for r in self.removed if r.category is None]

    @property
    def expected_modifications(self) -> list[ModifiedParagraph]:
        """Paragraphs trimmed by a rule that asked for exactly that."""
        return [
            m for m in self.modified
            if m.category is not None and m.category != PRESERVE_VIOLATION
        ]

    @property
    def unexpected_modifications(self) -> list[ModifiedParagraph]:
        """Paragraphs that lost text no rule accounts for."""
        return [m for m in self.modified if m.category is None]

    @property
    def preserve_violations(self) -> list:
        """Content that matched a preserve pattern and was removed anyway."""
        return (
            [r for r in self.removed if r.category == PRESERVE_VIOLATION]
            + [m for m in self.modified if m.category == PRESERVE_VIOLATION]
        )

    @property
    def removed_characters(self) -> int:
        """Characters of text the clean took out, removals and trims together.

        A more honest measure of what changed than the file size, which mostly
        reflects how the ZIP recompressed.
        """
        return (
            sum(len(r.text) for r in self.removed)
            + sum(len(fragment) for m in self.modified for fragment in m.fragments)
        )

    @property
    def reference_violations(self) -> list[StructuralViolation]:
        """Cross-references this run broke — a warning, not lost content."""
        return [v for v in self.structural if v.kind == "reference"]

    @property
    def structural_damage(self) -> list[StructuralViolation]:
        """Structural problems the output has that are not broken references."""
        return [v for v in self.structural if v.kind != "reference"]

    @property
    def passed(self) -> bool:
        return not (
            self.unexpected_removals
            or self.unexpected_modifications
            or self.preserve_violations
            or self.structural
            or self.added
            or self.numbering
        )

    def review_categories(self) -> set[ReviewCategory]:
        """Why this file needs review — every reason that applies, never one.

        A file can be in several categories at once and picking one to show
        would be the pooling §10.6 criterion 3 exists to prevent.

        The split between the first two turns on *how the verdict was reached*,
        not on how bad it sounds.  A preserve violation, an invented paragraph
        or a structural problem is a claim about the document: the comparison
        knew what it was looking at.  An unexplained removal or modification is
        only such a claim when the alignment behind it was exact — where it was
        a guess, what the report actually establishes is that the verifier
        could not follow the change, and calling that damage would overstate
        it in exactly the direction that trains a user to ignore the verdict.

        Configuration is deliberately absent here: nothing in a comparison of
        two documents can see it.  ``verdict_for`` adds it, from the rules the
        run was given.
        """
        categories: set[ReviewCategory] = set()

        unexplained = self.unexpected_removals + self.unexpected_modifications
        if any(finding.ambiguous for finding in unexplained):
            categories.add(ReviewCategory.AMBIGUOUS_ALIGNMENT)
        if (
            self.preserve_violations
            or self.added
            or self.structural_damage
            or any(not finding.ambiguous for finding in unexplained)
        ):
            categories.add(ReviewCategory.DETECTED_DAMAGE)
        if self.numbering or self.reference_violations:
            categories.add(ReviewCategory.REFERENCE_NUMBERING)

        return categories


# =============================================================================
# Text extraction
# =============================================================================

def _unpack(docx_path: Path) -> Path:
    """Extract a DOCX to a fresh temp directory; caller removes it."""
    temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_verify_"))
    with zipfile.ZipFile(docx_path, "r") as zf:
        zf.extractall(temp_dir)
    return temp_dir


def extract_paragraphs(
    docx_path: Path,
    engine: DetectionEngine,
    *,
    evidence: bool = True,
) -> list[ParagraphInfo]:
    """Extract every non-empty paragraph from a DOCX, with its source evidence.

    Both sides of the comparison go through this one walk, so input and output
    are always measured the same way.  Text-box paragraphs are counted once —
    through the paragraph that owns them, not again through the run that
    contains it.

    ``evidence=False`` skips policy evaluation, which is what the output side
    wants: the only question there is what text survived.  The extraction
    itself is shared either way, so the two sides cannot drift.

    Evaluating evidence binds ``engine`` to *this* document's styles, the same
    way the processor binds them before cleaning it.  Verification runs against
    the document just processed, so the binding it needs is the one already in
    place; rebinding states that rather than relying on it.
    """
    temp_dir = _unpack(docx_path)
    try:
        # Read once for the whole package: numbering is resolved against the
        # style chain, and re-parsing styles.xml per removal would be the same
        # answer computed again for every paragraph.
        styles = StyleIndex(load_styles(temp_dir / "word"))
        if evidence:
            engine.bind_styles(load_styles(temp_dir / "word"))
        paragraphs: list[ParagraphInfo] = []
        for xml_path in collect_content_parts(temp_dir / "word"):
            part = xml_path.relative_to(temp_dir).as_posix()
            root = parse_xml(xml_path).getroot()
            for para in iter_paragraphs(root, skip_alternate_fallback=True):
                raw_text = paragraph_text(para)
                if not raw_text.strip():
                    continue
                paragraphs.append(ParagraphInfo(
                    raw_text=raw_text,
                    evidence=(
                        engine.paragraph_evidence(para) if evidence
                        else ParagraphEvidence(raw_text=raw_text)
                    ),
                    run_profile=run_profile(para),
                    signature=paragraph_signature(para),
                    part=part,
                    story=note_identity(para),
                    numbering=numbering_id(para, styles),
                    in_tracked_deletion=_in_tracked_deletion(para),
                ))
        return paragraphs
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _in_tracked_deletion(para: etree._Element) -> bool:
    """True if accepting revisions would take this paragraph's text.

    Two shapes reach the comparison, and only one of them used to be checked.

    The *container* may be marked deleted — a table row records it in
    ``w:trPr``, a cell in ``w:tcPr`` — and its text stays ordinary ``w:t``,
    which is why it arrives here at all.

    Or every run carrying text may sit inside a revision whose content goes.
    ``w:del`` never shows up this way, because a deleted run holds ``w:delText``
    that no extractor reads; ``w:moveFrom`` does, because the source half of a
    move keeps real ``w:t`` until the move is accepted.  So accepting a tracked
    move made its paragraph look like an unexplained removal.
    """
    if in_tracked_deletion(para):
        return True
    runs = [run for run in iter_own_runs(para) if run_text(run).strip()]
    return bool(runs) and all(in_tracked_deletion(run) for run in runs)


# =============================================================================
# Structural inspection
# =============================================================================

def inspect_structure(docx_path: Path, strip_revisions: bool = False) -> StructureReport:
    """Inspect a DOCX for structure Word will refuse to open.

    Pattern matching cannot see any of this: a footer emptied of block
    content, a table cell that no longer ends with a paragraph, a field or
    bookmark range left half-open.
    """
    report = StructureReport()
    temp_dir = _unpack(docx_path)
    try:
        for xml_path in collect_content_parts(temp_dir / "word"):
            part = xml_path.name
            root = parse_xml(xml_path).getroot()

            for container in root.iter(*LINTED_CONTAINERS):
                blocks = block_children(container)
                tag = etree.QName(container).localname
                if not blocks:
                    report.issues[f"{part}: <w:{tag}> left with no block-level content"] += 1
                elif container.tag == TC_TAG and blocks[-1].tag != P_TAG:
                    report.issues[f"{part}: table cell does not end with a paragraph"] += 1

            for table in root.iter(TBL_TAG):
                if not any(child.tag == f"{W}tr" for child in table):
                    report.issues[f"{part}: <w:tbl> left with no rows"] += 1
            for row in root.iter(f"{W}tr"):
                if not any(child.tag == TC_TAG for child in row):
                    report.issues[f"{part}: <w:tr> left with no cells"] += 1

            if not field_chars_balanced(root):
                report.issues[f"{part}: unbalanced field characters"] += 1

            orphans = len(orphaned_range_markers(root))
            if orphans:
                report.issues[f"{part}: unmatched bookmark/comment range markers"] += orphans

            report.section_breaks += sum(1 for _ in root.iter(f"{W}sectPr"))

            for instruction, count in field_instructions(root, strip_revisions).items():
                report.fields[(part, instruction)] += count

            report.bookmarks |= bookmark_names(root)
            report.references.extend(
                (part, consumer) for consumer in reference_consumers(root))

        return report
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def lint_structure(docx_path: Path) -> list[str]:
    """Return a flat list of structural problems found in one DOCX."""
    report = inspect_structure(docx_path)
    return [
        str(StructuralViolation(issue, count))
        for issue, count in sorted(report.issues.items())
    ]


def _compare_structure(
    input_path: Path, output_path: Path, strip_revisions: bool = False
) -> list[StructuralViolation]:
    """Report structural damage the output has and the input did not.

    Documents arrive with oddities of their own; only what the clean added is
    the clean's fault.
    """
    before = inspect_structure(input_path, strip_revisions)
    after = inspect_structure(output_path)

    violations = [
        StructuralViolation(issue, count - before.issues[issue])
        for issue, count in sorted(after.issues.items())
        if count > before.issues[issue]
    ]

    lost_sections = before.section_breaks - after.section_breaks
    if lost_sections > 0:
        violations.append(
            StructuralViolation("section break(s) lost from the document", lost_sections)
        )

    # A reference the clean broke: something still names a bookmark that the
    # output no longer defines, and the input did define.  Only what this run
    # broke is reported — a reference already dangling on the way in is the
    # document's own problem, the same rule the issue counts above follow.
    #
    # There is deliberately no tracked-deletion exemption here, unlike the field
    # inventory.  Accepting a revision that deletes a referenced target is a
    # requested text deletion with an unrequested consequence, and the
    # consequence is what needs review.
    newly_broken = after.broken_references() - before.broken_references()
    for part, consumer in after.references:
        if consumer.folded in newly_broken and consumer.folded in before.bookmarks:
            violations.append(StructuralViolation(
                f"{part}: reference broken — {consumer} names {{{consumer.name}}}, "
                "which no bookmark defines",
                kind="reference",
            ))

    # A field carrier can vanish while the text stays identical — the cached
    # result reads as ordinary words, so nothing else in the comparison sees it
    # go.  Losing one turns a live cross-reference into a frozen string.
    for (part, instruction), count in sorted(before.fields.items()):
        lost = count - after.fields[(part, instruction)]
        if lost > 0:
            shown = instruction or "(no instruction)"
            violations.append(
                StructuralViolation(f"{part}: field lost — {{{shown}}}", lost)
            )

    return violations


# =============================================================================
# Classification
# =============================================================================

def _removed_intervals(before: str, after: str) -> list[tuple[int, int]] | None:
    """Intervals of ``before`` cut to produce ``after``.

    Returns None if the change was not a pure deletion — anything inserted or
    substituted means the text was mutated, which is never something the
    cleaner is supposed to do.

    Intervals, not strings.  Substring membership was the defect: a paragraph
    can hold the same words twice, once in an editorial run and once in a
    requirement, and a fragment compared against a list of editorial texts
    cannot say which occurrence actually went.
    """
    intervals: list[tuple[int, int]] = []
    matcher = difflib.SequenceMatcher(None, before, after, autojunk=False)
    for tag, i1, i2, _j1, _j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag != "delete":
            return None
        intervals.append((i1, i2))
    return intervals


def _substantive(text: str, intervals: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Drop intervals that are nothing but whitespace."""
    return [(i, j) for i, j in intervals if text[i:j].strip()]


# =============================================================================
# Main verification logic
# =============================================================================

def verify_clean(
    input_path: Path,
    output_path: Path,
    config_path: Path | None = None,
    engine: DetectionEngine | None = None,
    strip_revisions: bool = False,
) -> VerificationResult:
    """Compare input and output DOCX files, classifying every difference.

    Args:
        input_path:  Original DOCX before cleaning.
        output_path: Cleaned DOCX after processing.
        config_path: Path to patterns.yaml.  Consulted only when no engine is
                     supplied; resolved through ``apppaths`` when both are
                     absent.
        engine:      The engine that did the cleaning.  Passing it keeps
                     verification and detection on one set of patterns; if it
                     is omitted an equivalent engine is built from the config.
        strip_revisions: Whether the clean was asked to accept tracked changes.
                     Text lost to a tracked deletion is expected only then.

    Returns:
        VerificationResult with every removal, modification, and structural
        violation classified.
    """
    # Precedence, unchanged and now stated: an engine wins outright, an
    # explicit path comes next, and only with neither is the location resolved.
    #
    # The resolver is called *lazily* on purpose.  It can seed a per-user
    # patterns.yaml on first run, and a caller that supplied its own engine or
    # its own path has asked for neither that file nor that side effect.
    #
    # `Path(__file__).parent` was the old fallback and is wrong in a frozen
    # build: it names PyInstaller's extraction directory, which is temporary and
    # is not where the user's file lives.  An engine-less caller — a test, a
    # census tool, a future harness — would have verified against the shipped
    # defaults while the cleaner used the user's edits, and reported the
    # difference as damage.  The GUI always passes an engine, so this governs no
    # production run today; it is fixed because the next caller will not know
    # that.
    if engine is None:
        engine = DetectionEngine(load_config(config_path or resolve_config_path()))

    # Evidence is read from the input only.  The output side is asked one
    # question — what text survived — and policy is never re-derived from it:
    # judging a document by its own contents is how damage explains itself.
    input_paras = extract_paragraphs(input_path, engine)
    output_paras = extract_paragraphs(output_path, engine, evidence=False)
    input_texts = [p.text for p in input_paras]
    output_texts = [p.text for p in output_paras]

    result = VerificationResult(
        input_path=input_path,
        output_path=output_path,
        input_paragraph_count=len(input_texts),
        output_paragraph_count=len(output_texts),
        structural=_compare_structure(input_path, output_path, strip_revisions),
    )

    # Numbering lists are document-wide, so what survives is asked once across
    # every part rather than within each location.
    surviving_numbering = {p.numbering for p in output_paras if p.numbering}

    for location in _locations(input_paras, output_paras):
        _compare_location(
            [p for p in input_paras if p.location == location],
            [p for p in output_paras if p.location == location],
            result,
            strip_revisions,
            surviving_numbering,
        )

    return result


def _locations(
    input_paras: list[ParagraphInfo], output_paras: list[ParagraphInfo]
) -> list[tuple[str, str | None]]:
    """Every location either side holds, in input order then output-only order.

    A location is a package part plus, inside a part that holds several
    independent stories, the note it belongs to.  Comparison never crosses one:
    ``footnotes.xml`` is a single part containing many notes, and an identical
    paragraph in another note is not evidence about this one.
    """
    ordered = list(dict.fromkeys(p.location for p in input_paras))
    ordered += [
        location for location in dict.fromkeys(p.location for p in output_paras)
        if location not in set(ordered)
    ]
    return ordered


def _compare_location(
    input_paras: list[ParagraphInfo],
    output_paras: list[ParagraphInfo],
    result: VerificationResult,
    strip_revisions: bool,
    surviving_numbering: set[str],
) -> None:
    """Classify every difference within one location.

    A location missing from the output leaves ``output_paras`` empty, so all of
    its paragraphs are removals — and a location the output invented has no
    input side, so all of its paragraphs are additions.  Both fall out of the
    ordinary comparison rather than needing a case of their own.
    """
    kept = _pure_deletion_pairing(input_paras, output_paras)
    if kept is not None:
        _classify_pure_deletion(
            input_paras, output_paras, kept, result,
            strip_revisions, surviving_numbering,
        )
        return

    input_texts = [p.text for p in input_paras]
    output_texts = [p.text for p in output_paras]

    # Pairing first, classification second.  Attribution must know every
    # established pairing before it runs: a paragraph that survived in
    # shortened form is not missing, and offering it as the explanation for
    # some *other* paragraph's loss counts it twice and leaves the real loss
    # unclassified — a deleted requirement reported as a verified clean.
    removals: list[int] = []
    modifications: list[tuple[int, int, list[tuple[int, int]], bool]] = []
    survived: set[int] = set()

    matcher = difflib.SequenceMatcher(None, input_texts, output_texts, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            # Deliberately not reserved.  Inside an equal block the differ has
            # matched on text, which for repeated text says nothing about which
            # paragraph is which — reserving on that alignment would hand a
            # correct clean's survivor the identity of the copy that went.
            continue
        if tag == "insert":
            result.added.extend(output_texts[j1:j2])
            continue

        paired: set[int] = set()
        for idx in range(i1, i2):
            match = _pair_with_survivor(input_paras[idx], output_paras, j1, j2, paired)

            if match is None:
                removals.append(idx)
                continue

            jdx, intervals, guessed = match
            paired.add(jdx)
            survived.add(idx)
            if intervals:
                modifications.append((idx, jdx, intervals, guessed))

        # Anything in the replacement block that no input paragraph explains
        # is text the clean invented.
        result.added.extend(
            output_texts[jdx] for jdx in range(j1, j2) if jdx not in paired
        )

    for idx, jdx, intervals, guessed in modifications:
        result.modified.append(
            _classify_modification(
                input_paras[idx], output_paras[jdx], intervals, guessed
            )
        )

    lost = _lost_signatures(input_paras, output_paras)
    by_text = _index_by_text(input_paras)
    attributed = set(survived)
    for idx in removals:
        info = _attribute_removal(input_paras[idx], by_text, lost, attributed)
        result.removed.append(_classify_removal(info, strip_revisions))

        # Only when the list still has members: a list whose every paragraph
        # went renumbers nothing, and a notice about it would be noise
        # dressed as precision.
        if info.numbering and info.numbering in surviving_numbering:
            result.numbering.append(NumberingNotice(
                part=info.part, numbering=info.numbering,
                preview=info.text[:60] + ("…" if len(info.text) > 60 else ""),
            ))


def _pure_deletion_pairing(
    input_paras: list[ParagraphInfo], output_paras: list[ParagraphInfo]
) -> list[int] | None:
    """Input indices the output kept, when the output is a pure deletion.

    ``difflib`` costs O(n²) inside ``find_longest_match`` when text repeats,
    because every occurrence of a value is a candidate for every other.
    Measured on the adversarial family — one of three strings per paragraph —
    verification took 0.39s, 2.7s and 21.2s at 500, 1000 and 2000 paragraphs,
    with 94% of it inside that one function.

    A *pure deletion* needs none of that search.  Where every output paragraph
    matches an input paragraph exactly, in order, nothing was modified and
    nothing was invented, so every verdict is either "this survived unchanged"
    or "this went".  A greedy left-to-right scan finds that alignment in O(n),
    and greedy is not a heuristic here: if an in-order embedding exists at all,
    matching each output paragraph to the earliest unused input paragraph finds
    one.

    Matching is on the paragraph **signature**, not its text, and that is the
    difference between a fast path and a broken one.  Text equality is not
    identity: a hidden note beside an identical visible requirement extracts
    the same characters, so a text-only scan pairs the surviving *hidden* copy
    with the plain paragraph and reports the note as the removal — turning the
    loss of a requirement into a verified clean.  Both W04 discriminating cases
    caught exactly that when this was written on text.

    Signatures are stricter than text, which is the safe direction: a document
    this rejects simply takes the ordinary path.

    **Which** occurrence of a repeated signature it consumes is arbitrary, and
    that is precisely the arbitrariness ``_attribute_removal`` already exists
    to resolve.  The multiset of lost paragraphs is fixed by the two documents,
    so the verdict does not depend on the scan's choice.

    Returns ``None`` the moment the output is *not* a pure deletion — one
    modified paragraph, one invented one, one reordering, one run reshaped —
    and the ordinary comparison then runs unchanged.  This adds a fast path; it
    removes no reasoning.
    """
    if len(output_paras) > len(input_paras):
        return None            # something was added; not a deletion

    signatures = [para.signature for para in input_paras]
    kept: list[int] = []
    index = 0
    for para in output_paras:
        wanted = para.signature
        while index < len(signatures) and signatures[index] != wanted:
            index += 1
        if index == len(signatures):
            return None        # not an in-order embedding
        kept.append(index)
        index += 1
    return kept


def _classify_pure_deletion(
    input_paras: list[ParagraphInfo],
    output_paras: list[ParagraphInfo],
    kept: list[int],
    result: VerificationResult,
    strip_revisions: bool,
    surviving_numbering: set[str],
) -> None:
    """Judge a pure deletion: every unpaired input paragraph is a removal.

    Attribution runs exactly as it does on the general path, against the real
    output.  Substituting the paired input paragraphs for it would be the
    cleaner grading its own work with the evidence removed — the surviving
    hidden copy and the plain one it stood in for have different signatures,
    and that difference is the whole answer.
    """
    survived = set(kept)
    lost = _lost_signatures(input_paras, output_paras)
    by_text = _index_by_text(input_paras)
    attributed = set(survived)

    for index, para in enumerate(input_paras):
        if index in survived:
            continue
        info = _attribute_removal(para, by_text, lost, attributed)
        result.removed.append(_classify_removal(info, strip_revisions))
        if info.numbering and info.numbering in surviving_numbering:
            result.numbering.append(NumberingNotice(
                part=info.part, numbering=info.numbering,
                preview=info.text[:60] + ("…" if len(info.text) > 60 else ""),
            ))


def _lost_signatures(
    input_paras: list[ParagraphInfo], output_paras: list[ParagraphInfo]
) -> dict[str, Counter]:
    """Per text, which paragraph signatures the output no longer has.

    A multiset difference, so two identical paragraphs that both survive are
    not mistaken for one.  This is what says *which* occurrence of a repeated
    text actually disappeared — a question the text alone cannot answer.

    Grouped in one pass per side rather than rescanning both for every distinct
    text.  The rescan was O(distinct x n), which on a document of mostly unique
    paragraphs is O(n^2): once the paragraph matcher was no longer the
    bottleneck this became 88% of verification time at 4,000 paragraphs, and it
    was simply hidden behind the matcher before.  Same multisets, same answer.
    """
    before: dict[str, Counter] = defaultdict(Counter)
    for para in input_paras:
        before[para.text][para.signature] += 1

    after: dict[str, Counter] = defaultdict(Counter)
    for para in output_paras:
        after[para.text][para.signature] += 1

    lost: dict[str, Counter] = {}
    for text, counts in before.items():
        missing = counts - after[text] if text in after else counts
        if missing:
            lost[text] = missing
    return lost


def _index_by_text(
    paragraphs: list[ParagraphInfo],
) -> dict[str, list[tuple[int, ParagraphInfo]]]:
    """Where each text occurs, built once per location.

    ``_attribute_removal`` used to rescan every paragraph for every removal,
    which is O(removals x n) — 0.95s of a 1.79s verification at 4,000
    paragraphs once the earlier hotspots were gone.  Grouping first is the same
    lookup, computed once.
    """
    grouped: dict[str, list[tuple[int, ParagraphInfo]]] = defaultdict(list)
    for index, para in enumerate(paragraphs):
        grouped[para.text].append((index, para))
    return grouped


def _attribute_removal(
    info: ParagraphInfo,
    by_text: dict[str, list[tuple[int, ParagraphInfo]]],
    lost: dict[str, Counter],
    attributed: set[int],
) -> ParagraphInfo:
    """Decide which source paragraph actually went, when several could have.

    Where a location holds the same text more than once, the differ's choice of
    which occurrence to call deleted is arbitrary: it aligns on the longest
    matching block, not on evidence.  So a document holding a plain requirement
    and an identical hidden note had a *correct* clean reported as damage — the
    note was removed, and the plain copy was blamed.

    The signatures say which paragraph is genuinely absent from the output, so
    the verdict is taken against that one.  This only ever re-attributes among
    paragraphs whose text is already identical, and only to a signature the
    output really is missing; where the text occurs once there is nothing to
    choose and ``info`` is returned unchanged.
    """
    remaining = lost.get(info.text)
    if remaining is None:
        return info

    same_text = by_text.get(info.text, ())
    if len(same_text) < 2:
        return info

    for index, para in same_text:
        if index in attributed:
            continue  # already paired with a survivor, or already blamed
        if remaining.get(para.signature, 0) > 0:
            remaining[para.signature] -= 1
            attributed.add(index)
            return para
    return info


def _pair_with_survivor(
    info: ParagraphInfo,
    output_paras: list[ParagraphInfo],
    j1: int,
    j2: int,
    paired: set[int],
) -> tuple[int, list[tuple[int, int]]] | None:
    """Find the output paragraph this input paragraph turned into, if any.

    Exact answers are taken first: an unchanged paragraph, or one that matches
    exactly what cutting every authorized placeholder out of the source would
    produce.  Both are certainties.  ``MIN_PAIR_SIMILARITY`` is a guess, and a
    guess that fails on precisely the redactions that worked: for a pure
    deletion the ratio falls below 0.5 once more than two-thirds of the
    characters go, so "Provide [Verify quantity with the Owner and the AHJ
    prior to bid] units." correctly cleaned to "Provide units." was reported
    as an unexplained removal plus an invented paragraph.

    Recognising the exact result does not widen what counts as permitted: an
    output that is anything other than that exact text still has to satisfy
    the similarity path below, unchanged, and whatever it lost still has to be
    covered by an authorized interval before it counts as expected.

    Returns the survivor's index, the intervals of ``info.text`` it lost, and
    whether that pairing was reached by the guess rather than by an exact
    answer.  The caller needs the third value to say *why* a file needs
    review: an unexplained loss found on an exact pairing is a claim about the
    document, and the same loss found on a guessed one is the verifier saying
    it could not follow what happened.  Case A of the W07 probe shows the
    difference plainly — a similarity pairing reported the fragments ``'nd
    hangers a'`` and ``' on drawings'``, which are artefacts of where the
    differ happened to align, not text anyone edited out.
    """
    expected = info.expected_after_removal()

    for jdx in range(j1, j2):
        if jdx in paired:
            continue
        after = output_paras[jdx].text
        if after == info.text:
            return jdx, [], False
        if expected is not None and after == expected:
            # Matching text is not on its own evidence that the *right* text
            # went: a paragraph holding a visible requirement and an identical
            # hidden copy produces this same string whichever one was lost.  So
            # the runs have to agree as well before the intervals are taken as
            # known; otherwise the ordinary reasoning below decides, and the
            # loss of the visible copy is still reported.
            if output_paras[jdx].run_profile == info.evidence.expected_profile():
                return jdx, _substantive(info.text, info.authorized_intervals()), False
            # The text is what a correct clean would produce but the runs are
            # not, so which occurrence survived is exactly what is unknown.
            # Intervals re-derived by diffing here are a guess about *which*
            # characters went, and are marked as one.
            intervals = _removed_intervals(info.text, after) or []
            return jdx, _substantive(info.text, intervals), True

    for jdx in range(j1, j2):
        if jdx in paired:
            continue
        after = output_paras[jdx].text

        intervals = _removed_intervals(info.text, after)
        if intervals is None:
            continue
        similarity = difflib.SequenceMatcher(
            None, info.text, after, autojunk=False).ratio()
        if similarity < MIN_PAIR_SIMILARITY:
            continue
        return jdx, _substantive(info.text, intervals), True

    return None


def _classify_removal(
    info: ParagraphInfo,
    strip_revisions: bool = False,
) -> RemovedParagraph:
    """Decide whether a vanished paragraph was meant to vanish.

    The order is the plan's precedence: an explicit tracked deletion is
    separate authority and outranks everything, protection outranks every
    removal rule, and only then does policy get to explain the loss.

    Two things authorize losing a whole paragraph, and a signal *somewhere
    inside it* is neither.  Either a rule qualified against the paragraph
    itself, or the authorized intervals account for every substantive
    character in it.  One hidden run in a paragraph of requirements permits
    losing that run's own text, and nothing beside it.
    """
    if strip_revisions and info.in_tracked_deletion:
        # An explicit tracked deletion is separate authority from editorial
        # matching, and it outranks protection: the author deleted this row or
        # cell on purpose, and accepting revisions is what the run was asked to
        # do.  A preserved heading inside a deleted row is that deletion working,
        # not a violation.  The extent is validated — _in_tracked_deletion walks
        # to the enclosing w:tr or w:tc and requires the marker there — so a
        # revision somewhere nearby is not blanket permission.
        return RemovedParagraph(
            info.text, TRACKED_DELETION, "accepted a tracked deletion"
        )

    if info.preserve_reason is not None:
        return RemovedParagraph(info.text, PRESERVE_VIOLATION, info.preserve_reason)

    evidence = info.evidence
    if evidence.whole_category is not None:
        return RemovedParagraph(
            info.text,
            FORMATTING_BASED if evidence.whole_formatting_only else evidence.whole_category,
            evidence.whole_reason,
        )

    if evidence.covers_all_text():
        authorities = evidence.authorities()
        if authorities:
            return RemovedParagraph(info.text, *_verdict(authorities))

    # Unexplained.  A removal is never paired, so there is no similarity guess
    # behind it; the one way this verdict can rest on the verifier's own limits
    # is unreliable offsets, which make covers_all_text() answer False whatever
    # the paragraph holds.
    return RemovedParagraph(
        info.text, None, None, ambiguous=not evidence.offsets_reliable
    )


def _classify_modification(
    info: ParagraphInfo,
    survivor: ParagraphInfo,
    intervals: list[tuple[int, int]],
    guessed_pairing: bool = False,
) -> ModifiedParagraph:
    """Decide whether the text a surviving paragraph lost was meant to go.

    Every lost interval has to be *covered by* an authorized one — not merely
    to contain something that matches a rule.  ``[Verify quantity]`` sitting
    next to ``spare`` authorizes cutting the placeholder; it says nothing about
    the word beside it, and the containment test could not tell the difference.

    A protected paragraph is not touched by the cleaner at all, so any loss
    inside one is a violation whatever the lost text looks like — judged on the
    original paragraph, which is the only place the protection is visible.

    One exact answer comes before the intervals, because the intervals rest on
    an alignment that repeated text can make arbitrary.  A paragraph holding a
    hidden note and then an identical visible requirement extracts the same
    characters twice; when the cleaner removes the note, the survivor is equally
    consistent with either occurrence having gone, and ``SequenceMatcher``
    simply picks one.  Comparing the output's own runs against the runs a
    correct clean would leave settles it on evidence rather than on which
    alignment the differ happened to choose.

    That check only ever accepts.  When the signatures do not match, the
    interval reasoning below runs unchanged — so an output keeping the *hidden*
    copy while the visible requirement disappeared is still reported, which is
    the case textual comparison alone cannot distinguish from a correct clean.
    """
    after = survivor.text
    fragments = [info.text[start:end] for start, end in intervals]

    if info.preserve_reason is not None:
        return ModifiedParagraph(
            info.text, after, fragments, PRESERVE_VIOLATION, info.preserve_reason
        )

    expected = info.evidence.expected_profile()
    if expected and survivor.run_profile == expected:
        authorized = info.evidence.authorities()
        if authorized:
            category, pattern = _verdict(authorized)
            return ModifiedParagraph(info.text, after, fragments, category, pattern)

    authorities: list[tuple[str, str, bool]] = []
    for start, end in intervals:
        authority = info.authority_for(start, end)
        if authority is None:
            # Unexplained.  Whether that is a claim about the document or the
            # verifier declining to follow it depends on how this paragraph was
            # paired, and on whether its offsets could be trusted at all — with
            # unreliable offsets every interval question answers "no authority",
            # so the absence of one says nothing.
            return ModifiedParagraph(
                info.text, after, fragments,
                ambiguous=guessed_pairing or not info.evidence.offsets_reliable,
            )
        authorities.append(authority)

    category, pattern = _verdict(authorities)
    return ModifiedParagraph(info.text, after, fragments, category, pattern)


def _verdict(authorities: list[tuple[str, str, bool]]) -> tuple[str, str | None]:
    """Reduce the rules behind a loss to one category and its reasons.

    The first category is reported, with every distinct reason behind it, so a
    loss several rules jointly account for is not labelled with only the first
    one that happened to match.
    """
    category, _, formatting_only = authorities[0]
    return (
        FORMATTING_BASED if formatting_only else category,
        "; ".join(dict.fromkeys(reason for _, reason, _ in authorities)) or None,
    )
