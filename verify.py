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

from detection import DetectionEngine
from docx_xml import (
    P_TAG,
    StyleIndex,
    TC_TAG,
    W,
    block_children,
    collect_content_parts,
    field_chars_balanced,
    fold_style_names,
    is_on,
    iter_own_runs,
    iter_paragraphs,
    load_config,
    load_styles,
    orphaned_range_markers,
    paragraph_text,
    parse_xml,
    run_text,
    toggle_on,
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


# =============================================================================
# Data structures
# =============================================================================

@dataclass
class ParagraphInfo:
    """Paragraph text with formatting metadata from the source DOCX."""
    text: str
    has_editorial_style: bool = False
    is_italic: bool = False
    has_editorial_color: bool = False
    is_hidden: bool = False
    editorial_run_texts: list[str] = field(default_factory=list)

    def formatting_signals(self, trust_formatting_only: bool) -> list[str]:
        """Names of the editorial formatting signals this paragraph carries.

        ``trust_formatting_only`` mirrors ``specifier_notes.formatting_only_removal``:
        when the cleaner removes text on italic + editorial colour alone, so
        must verification accept it, and when it does not, such a removal is
        worth a warning.
        """
        signals: list[str] = []
        if self.has_editorial_style:
            signals.append("editorial style")
        if self.is_hidden:
            signals.append("hidden text")
        if trust_formatting_only and self.is_italic and self.has_editorial_color:
            signals.append("italic + editorial colour")
        return signals


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


def extract_paragraphs(docx_path: Path, config: dict) -> list[ParagraphInfo]:
    """Extract every non-empty paragraph from a DOCX, with formatting metadata.

    Used for both sides of the comparison, so input and output are always
    measured the same way.  Text-box paragraphs are counted once — through
    the paragraph that owns them, not again through the run that contains it.
    """
    style_section = config.get("style_based_detection", {})
    editorial_styles = (
        fold_style_names(
            style_section.get("paragraph_styles", [])
            + style_section.get("character_styles", [])
        )
        if style_section.get("enabled", True) else frozenset()
    )

    fmt_signals = config.get("specifier_notes", {}).get("formatting_signals", {})
    editorial_colors: set[str] = {c.upper() for c in fmt_signals.get("colors", [])}

    temp_dir = _unpack(docx_path)
    try:
        styles = StyleIndex(load_styles(temp_dir / "word"))
        paragraphs: list[ParagraphInfo] = []
        for xml_path in collect_content_parts(temp_dir / "word"):
            root = parse_xml(xml_path).getroot()
            for para in iter_paragraphs(root, skip_alternate_fallback=True):
                info = _describe_paragraph(para, styles, editorial_styles, editorial_colors)
                if info.text:
                    paragraphs.append(info)
        return paragraphs
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _describe_paragraph(
    para: etree._Element,
    styles: StyleIndex,
    editorial_styles: frozenset[str],
    editorial_colors: set[str],
) -> ParagraphInfo:
    """Read one paragraph's text and editorial formatting signals."""
    has_editorial_style = False
    has_editorial_color = False
    is_hidden = False
    runs_with_text = 0
    italic_runs = 0
    editorial_run_texts: list[str] = []

    para_style = None
    ppr = para.find(f"{W}pPr")
    if ppr is not None:
        pstyle = ppr.find(f"{W}pStyle")
        if pstyle is not None:
            para_style = pstyle.get(f"{W}val")
            has_editorial_style = styles.matches(para_style, editorial_styles)

    for run in iter_own_runs(para):
        text = run_text(run)
        if not text.strip():
            continue
        runs_with_text += 1

        rpr = run.find(f"{W}rPr")
        run_italic = toggle_on(rpr, f"{W}i")
        run_colored = False
        run_styled = False
        run_style = None

        if rpr is not None:
            color_elem = rpr.find(f"{W}color")
            if color_elem is not None:
                run_colored = (color_elem.get(f"{W}val") or "").upper() in editorial_colors
            rstyle = rpr.find(f"{W}rStyle")
            if rstyle is not None:
                run_style = rstyle.get(f"{W}val")
                run_styled = styles.matches(run_style, editorial_styles)

        # Hidden resolves as Word resolves it: an explicit w:vanish on the run
        # wins (w:val="0" un-hides), otherwise the styles decide.
        vanish = rpr.find(f"{W}vanish") if rpr is not None else None
        if vanish is not None:
            run_hidden = is_on(vanish)
        else:
            run_hidden = styles.is_hidden(run_style) or styles.is_hidden(para_style)

        italic_runs += bool(run_italic)
        has_editorial_color |= run_colored
        is_hidden |= run_hidden
        has_editorial_style |= run_styled

        if run_hidden or run_styled or (run_italic and run_colored):
            editorial_run_texts.append(text)

    return ParagraphInfo(
        text=paragraph_text(para).strip(),
        has_editorial_style=has_editorial_style,
        is_italic=runs_with_text > 0 and italic_runs == runs_with_text,
        has_editorial_color=has_editorial_color,
        is_hidden=is_hidden,
        editorial_run_texts=editorial_run_texts,
    )


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

def _match_patterns(text: str, patterns) -> tuple[str | None, str | None]:
    """Return the first ``(category, pattern)`` that matches, else (None, None)."""
    for category, pattern in patterns:
        if pattern.search(text):
            return category, pattern.pattern
    return None, None


def _matches_preserve(text: str, preserve_patterns) -> str | None:
    for pattern in preserve_patterns:
        if pattern.search(text):
            return pattern.pattern
    return None


def _removed_fragments(before: str, after: str) -> list[str] | None:
    """Fragments cut from ``before`` to make ``after``.

    Returns None if the change was not a pure deletion — anything inserted or
    substituted means the text was mutated, which is never something the
    cleaner is supposed to do.
    """
    fragments: list[str] = []
    matcher = difflib.SequenceMatcher(None, before, after, autojunk=False)
    for tag, i1, i2, _j1, _j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag != "delete":
            return None
        fragments.append(before[i1:i2])
    return fragments


# =============================================================================
# Main verification logic
# =============================================================================

def verify_clean(
    input_path: Path,
    output_path: Path,
    config_path: Path | None = None,
    engine: DetectionEngine | None = None,
) -> VerificationResult:
    """Compare input and output DOCX files, classifying every difference.

    Args:
        input_path:  Original DOCX before cleaning.
        output_path: Cleaned DOCX after processing.
        config_path: Path to patterns.yaml (auto-detected if None).
        engine:      The engine that did the cleaning.  Passing it keeps
                     verification and detection on one set of patterns; if it
                     is omitted an equivalent engine is built from the config.

    Returns:
        VerificationResult with every removal, modification, and structural
        violation classified.
    """
    if config_path is None:
        config_path = Path(__file__).parent / "patterns.yaml"
    config = engine.config if engine is not None else load_config(config_path)
    if engine is None:
        engine = DetectionEngine(config)

    removal_patterns = engine.removal_patterns()
    preserve_patterns = engine.preserve_patterns()
    trust_formatting_only = config.get("specifier_notes", {}).get(
        "formatting_only_removal", True
    )

    input_paras = extract_paragraphs(input_path, config)
    output_paras = extract_paragraphs(output_path, config)
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
            match = _pair_with_survivor(info.text, output_texts, j1, j2, paired)

            if match is None:
                result.removed.append(
                    _classify_removal(
                        info, removal_patterns, preserve_patterns, trust_formatting_only
                    )
                )
                continue

            jdx, fragments = match
            paired.add(jdx)
            if fragments:
                result.modified.append(
                    _classify_modification(
                        info, output_texts[jdx], fragments,
                        removal_patterns, preserve_patterns,
                    )
                )

        # Anything in the replacement block that no input paragraph explains
        # is text the clean invented.
        result.added.extend(
            output_texts[jdx] for jdx in range(j1, j2) if jdx not in paired
        )

    return result


def _pair_with_survivor(
    text: str,
    output_texts: list[str],
    j1: int,
    j2: int,
    paired: set[int],
) -> tuple[int, list[str]] | None:
    """Find the output paragraph this input paragraph turned into, if any."""
    for jdx in range(j1, j2):
        if jdx in paired:
            continue
        after = output_texts[jdx]
        if after == text:
            return jdx, []

        fragments = _removed_fragments(text, after)
        if fragments is None:
            continue
        similarity = difflib.SequenceMatcher(None, text, after, autojunk=False).ratio()
        if similarity < MIN_PAIR_SIMILARITY:
            continue
        return jdx, [f for f in fragments if f.strip()]

    return None


def _classify_removal(
    info: ParagraphInfo,
    removal_patterns,
    preserve_patterns,
    trust_formatting_only: bool,
) -> RemovedParagraph:
    """Decide whether a vanished paragraph was meant to vanish."""
    preserve_match = _matches_preserve(info.text, preserve_patterns)
    if preserve_match is not None:
        return RemovedParagraph(info.text, PRESERVE_VIOLATION, preserve_match)

    category, pattern = _match_patterns(info.text, removal_patterns)
    if category is None:
        signals = info.formatting_signals(trust_formatting_only)
        if signals:
            category, pattern = FORMATTING_BASED, "; ".join(signals)

    return RemovedParagraph(info.text, category, pattern)


def _classify_modification(
    info: ParagraphInfo,
    after: str,
    fragments: list[str],
    removal_patterns,
    preserve_patterns,
) -> ModifiedParagraph:
    """Decide whether the text a surviving paragraph lost was meant to go.

    Every fragment has to be accounted for; the worst verdict wins.
    """
    verdicts: list[tuple[str | None, str | None]] = []

    for fragment in fragments:
        preserve_match = _matches_preserve(fragment, preserve_patterns)
        if preserve_match is not None:
            return ModifiedParagraph(
                info.text, after, fragments, PRESERVE_VIOLATION, preserve_match
            )

        category, pattern = _match_patterns(fragment, removal_patterns)
        if category is None and any(
            fragment in run for run in info.editorial_run_texts
        ):
            category, pattern = FORMATTING_BASED, "editorial run formatting"
        verdicts.append((category, pattern))

    if any(category is None for category, _ in verdicts):
        return ModifiedParagraph(info.text, after, fragments)

    return ModifiedParagraph(
        before=info.text,
        after=after,
        fragments=fragments,
        category=verdicts[0][0],
        pattern_matched="; ".join(
            dict.fromkeys(pattern for _, pattern in verdicts if pattern)
        ) or None,
    )
