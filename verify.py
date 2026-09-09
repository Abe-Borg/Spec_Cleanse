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
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

from detection import DetectionEngine, ParagraphEvidence
from docx_xml import (
    P_TAG,
    TC_TAG,
    W,
    block_children,
    collect_content_parts,
    field_chars_balanced,
    cut_spans,
    iter_paragraphs,
    paragraph_signature,
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
    signature: tuple = ()
    #: True if a tracked change marks this paragraph's container as deleted —
    #: a deleted table row keeps its text in plain w:t, so nothing else shows it.
    in_tracked_deletion: bool = False

    @property
    def text(self) -> str:
        return self.raw_text.strip()

    @property
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

    def expected_after_redaction(self) -> str | None:
        """What this paragraph becomes when every authorized placeholder is cut.

        ``None`` when no placeholder is authorized, so there is no permitted
        transformation to expect.  The intervals come from evaluating the
        configured patterns against this source paragraph — never from anything
        the processor reports about what it did.
        """
        spans = list(self.evidence.inline_spans)
        if not spans:
            return None
        return cut_spans(self.raw_text, spans).strip()


@dataclass
class RemovedParagraph:
    """A paragraph present in input but absent in output."""
    text: str
    category: str | None = None       # e.g. "specifier_note", or None if unexpected
    pattern_matched: str | None = None # the regex or signal that matched


@dataclass
class ModifiedParagraph:
    """A paragraph that survived but lost some of its text."""
    before: str
    after: str
    fragments: list[str] = field(default_factory=list)
    category: str | None = None
    pattern_matched: str | None = None

    @property
    def text(self) -> str:
        """The text that went missing, for reporting alongside removals."""
        return " / ".join(self.fragments)


@dataclass
class StructuralViolation:
    """Damage to the document's structure that the output has and the input did not."""
    issue: str
    count: int = 1

    def __str__(self) -> str:
        return f"{self.issue}" + (f" (x{self.count})" if self.count > 1 else "")


@dataclass
class StructureReport:
    """What an inspection of one DOCX found."""
    issues: Counter = field(default_factory=Counter)
    section_breaks: int = 0


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
    def passed(self) -> bool:
        return not (
            self.unexpected_removals
            or self.unexpected_modifications
            or self.preserve_violations
            or self.structural
            or self.added
        )


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
        if evidence:
            engine.bind_styles(load_styles(temp_dir / "word"))
        paragraphs: list[ParagraphInfo] = []
        for xml_path in collect_content_parts(temp_dir / "word"):
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
                    signature=paragraph_signature(para),
                    in_tracked_deletion=_in_tracked_deletion(para),
                ))
        return paragraphs
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _in_tracked_deletion(para: etree._Element) -> bool:
    """True if a tracked change marks the row or cell holding this paragraph.

    Run-level deletions hide their text in ``w:delText``, which no extractor
    reads, so they never reach the comparison.  A deleted *row* is different:
    its text stays ordinary ``w:t`` and only ``w:trPr`` records the deletion.
    """
    node = para
    while node is not None:
        if node.tag == f"{W}tr" and node.find(f"{W}trPr/{W}del") is not None:
            return True
        if node.tag == TC_TAG and node.find(f"{W}tcPr/{W}cellDel") is not None:
            return True
        node = node.getparent()
    return False


# =============================================================================
# Structural inspection
# =============================================================================

def inspect_structure(docx_path: Path) -> StructureReport:
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

            if not field_chars_balanced(root):
                report.issues[f"{part}: unbalanced field characters"] += 1

            orphans = len(orphaned_range_markers(root))
            if orphans:
                report.issues[f"{part}: unmatched bookmark/comment range markers"] += orphans

            report.section_breaks += sum(1 for _ in root.iter(f"{W}sectPr"))

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


def _compare_structure(input_path: Path, output_path: Path) -> list[StructuralViolation]:
    """Report structural damage the output has and the input did not.

    Documents arrive with oddities of their own; only what the clean added is
    the clean's fault.
    """
    before = inspect_structure(input_path)
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
        config_path: Path to patterns.yaml (auto-detected if None).
        engine:      The engine that did the cleaning.  Passing it keeps
                     verification and detection on one set of patterns; if it
                     is omitted an equivalent engine is built from the config.
        strip_revisions: Whether the clean was asked to accept tracked changes.
                     Text lost to a tracked deletion is expected only then.

    Returns:
        VerificationResult with every removal, modification, and structural
        violation classified.
    """
    if engine is None:
        if config_path is None:
            config_path = Path(__file__).parent / "patterns.yaml"
        engine = DetectionEngine(load_config(config_path))

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
        structural=_compare_structure(input_path, output_path),
    )

    matcher = difflib.SequenceMatcher(None, input_texts, output_texts, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag == "insert":
            result.added.extend(output_texts[j1:j2])
            continue

        paired: set[int] = set()
        for idx in range(i1, i2):
            info = input_paras[idx]
            match = _pair_with_survivor(info, output_texts, j1, j2, paired)

            if match is None:
                result.removed.append(_classify_removal(info, strip_revisions))
                continue

            jdx, intervals = match
            paired.add(jdx)
            if intervals:
                result.modified.append(
                    _classify_modification(info, output_paras[jdx], intervals)
                )

        # Anything in the replacement block that no input paragraph explains
        # is text the clean invented.
        result.added.extend(
            output_texts[jdx] for jdx in range(j1, j2) if jdx not in paired
        )

    return result


def _pair_with_survivor(
    info: ParagraphInfo,
    output_texts: list[str],
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

    Returns the survivor's index and the intervals of ``info.text`` it lost.
    """
    expected = info.expected_after_redaction()

    for jdx in range(j1, j2):
        if jdx in paired:
            continue
        after = output_texts[jdx]
        if after == info.text:
            return jdx, []
        if expected is not None and after == expected:
            intervals = _removed_intervals(info.text, after) or []
            return jdx, _substantive(info.text, intervals)

    for jdx in range(j1, j2):
        if jdx in paired:
            continue
        after = output_texts[jdx]

        intervals = _removed_intervals(info.text, after)
        if intervals is None:
            continue
        similarity = difflib.SequenceMatcher(
            None, info.text, after, autojunk=False).ratio()
        if similarity < MIN_PAIR_SIMILARITY:
            continue
        return jdx, _substantive(info.text, intervals)

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

    return RemovedParagraph(info.text, None, None)


def _classify_modification(
    info: ParagraphInfo,
    survivor: ParagraphInfo,
    intervals: list[tuple[int, int]],
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

    expected = info.evidence.surviving_signature()
    if expected and survivor.signature == expected:
        authorized = info.evidence.authorities()
        if authorized:
            category, pattern = _verdict(authorized)
            return ModifiedParagraph(info.text, after, fragments, category, pattern)

    authorities: list[tuple[str, str, bool]] = []
    for start, end in intervals:
        authority = info.authority_for(start, end)
        if authority is None:
            return ModifiedParagraph(info.text, after, fragments)
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
