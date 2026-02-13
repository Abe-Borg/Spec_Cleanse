"""
SpecCleanse Verification Module

Compares input and output DOCX files to verify that only expected
content (bloat, editorial notes, etc.) was removed and no real
specification content was lost.

Classification is two-pronged:
  1. Text-pattern matching against patterns.yaml
  2. Formatting-based detection (italic + editorial color, specifier
     paragraph/character style, hidden text) from the source DOCX
"""

import difflib
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import yaml
from lxml import etree

# Word namespace
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


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

    @property
    def has_editorial_formatting(self) -> bool:
        """True if formatting signals suggest editorial/specifier content."""
        return (
            self.has_editorial_style
            or self.is_hidden
            or (self.is_italic and self.has_editorial_color)
        )


@dataclass
class RemovedParagraph:
    """A paragraph present in input but absent in output."""
    text: str
    category: str | None = None       # e.g. "specifier_note", or None if unexpected
    pattern_matched: str | None = None # the regex or signal that matched


PRESERVE_VIOLATION = "preserve_violation"


@dataclass
class VerificationResult:
    """Structured result of input-vs-output comparison."""
    input_path: Path
    output_path: Path
    input_paragraph_count: int = 0
    output_paragraph_count: int = 0
    removed: list[RemovedParagraph] = field(default_factory=list)

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
    def preserve_violations(self) -> list[RemovedParagraph]:
        """Removals that matched a preserve pattern — these should never be removed."""
        return [r for r in self.removed if r.category == PRESERVE_VIOLATION]

    @property
    def passed(self) -> bool:
        return (
            len(self.unexpected_removals) == 0
            and len(self.preserve_violations) == 0
        )


# =============================================================================
# Text extraction
# =============================================================================

def _collect_xml_files(word_dir: Path) -> list[Path]:
    """Return document.xml, headers, footers in a stable order."""
    xml_files: list[Path] = []
    doc_xml = word_dir / "document.xml"
    if doc_xml.exists():
        xml_files.append(doc_xml)
    xml_files.extend(sorted(word_dir.glob("header*.xml")))
    xml_files.extend(sorted(word_dir.glob("footer*.xml")))
    return xml_files


def extract_text(docx_path: Path) -> list[str]:
    """Extract paragraph-level text from a DOCX file.

    Returns an ordered list of non-empty paragraph strings
    (document.xml first, then headers, then footers).
    """
    temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_verify_"))
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            zf.extractall(temp_dir)

        paragraphs: list[str] = []
        for xml_path in _collect_xml_files(temp_dir / "word"):
            parser = etree.XMLParser(remove_blank_text=False)
            tree = etree.parse(str(xml_path), parser)
            root = tree.getroot()

            for para in root.iter(f"{W}p"):
                texts: list[str] = []
                for t in para.iter(f"{W}t"):
                    if t.text:
                        texts.append(t.text)
                text = "".join(texts).strip()
                if text:
                    paragraphs.append(text)

        return paragraphs

    finally:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)


def _extract_paragraphs_with_formatting(
    docx_path: Path, config: dict
) -> list[ParagraphInfo]:
    """Extract paragraphs from a DOCX with formatting metadata.

    Uses the style and colour lists from the config to determine
    whether each paragraph carries editorial formatting signals.
    """
    # Gather editorial style names and colours from config
    style_section = config.get("style_based_detection", {})
    editorial_styles: set[str] = set(
        style_section.get("paragraph_styles", [])
        + style_section.get("character_styles", [])
    )

    spec_section = config.get("specifier_notes", {})
    fmt_signals = spec_section.get("formatting_signals", {})
    editorial_colors: set[str] = {c.upper() for c in fmt_signals.get("colors", [])}

    temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_verify_"))
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            zf.extractall(temp_dir)

        paragraphs: list[ParagraphInfo] = []

        for xml_path in _collect_xml_files(temp_dir / "word"):
            parser = etree.XMLParser(remove_blank_text=False)
            tree = etree.parse(str(xml_path), parser)
            root = tree.getroot()

            for para in root.iter(f"{W}p"):
                texts: list[str] = []
                has_editorial_style = False
                has_editorial_color = False
                is_hidden = False
                runs_with_text = 0
                italic_runs = 0

                # Check paragraph style
                ppr = para.find(f"{W}pPr")
                if ppr is not None:
                    pstyle = ppr.find(f"{W}pStyle")
                    if pstyle is not None:
                        val = pstyle.get(f"{W}val")
                        if val in editorial_styles:
                            has_editorial_style = True

                # Walk runs
                for run in para.iter(f"{W}r"):
                    run_texts: list[str] = []
                    for t in run.iter(f"{W}t"):
                        if t.text:
                            run_texts.append(t.text)
                    run_text = "".join(run_texts)
                    texts.append(run_text)

                    if not run_text.strip():
                        continue
                    runs_with_text += 1

                    rpr = run.find(f"{W}rPr")
                    if rpr is not None:
                        if rpr.find(f"{W}i") is not None:
                            italic_runs += 1
                        color_elem = rpr.find(f"{W}color")
                        if color_elem is not None:
                            color_val = (color_elem.get(f"{W}val") or "").upper()
                            if color_val in editorial_colors:
                                has_editorial_color = True
                        if rpr.find(f"{W}vanish") is not None:
                            is_hidden = True
                        rstyle = rpr.find(f"{W}rStyle")
                        if rstyle is not None:
                            val = rstyle.get(f"{W}val")
                            if val in editorial_styles:
                                has_editorial_style = True

                is_italic = runs_with_text > 0 and italic_runs == runs_with_text

                text = "".join(texts).strip()
                if text:
                    paragraphs.append(ParagraphInfo(
                        text=text,
                        has_editorial_style=has_editorial_style,
                        is_italic=is_italic,
                        has_editorial_color=has_editorial_color,
                        is_hidden=is_hidden,
                    ))

        return paragraphs

    finally:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)


# =============================================================================
# Pattern classification
# =============================================================================

def _build_removal_patterns(config: dict) -> list[tuple[str, re.Pattern]]:
    """Compile all text-based removal patterns from the config.

    Returns a list of (category_name, compiled_regex) tuples.
    """
    category_map = {
        "specifier_notes": "specifier_note",
        "copyright_notices": "copyright",
        "specagent_references": "specagent",
        "editorial_artifacts": "editorial_artifact",
    }

    patterns: list[tuple[str, re.Pattern]] = []

    for section_key, category_name in category_map.items():
        section = config.get(section_key, {})
        if not section.get("enabled", True):
            continue
        for pat_str in section.get("text_patterns", []):
            try:
                compiled = re.compile(pat_str, re.IGNORECASE | re.DOTALL)
                patterns.append((category_name, compiled))
            except re.error:
                continue

    return patterns


def _build_preserve_patterns(config: dict) -> list[re.Pattern]:
    """Compile preserve patterns from the config.

    These patterns identify content that should NEVER be removed.
    If removed content matches one of these, it's a verification error.
    """
    preserve_section = config.get("preserve_patterns", {})
    if not preserve_section.get("enabled", True):
        return []

    patterns: list[re.Pattern] = []
    for pat_str in preserve_section.get("text_patterns", []):
        try:
            patterns.append(re.compile(pat_str, re.IGNORECASE | re.DOTALL))
        except re.error:
            continue
    return patterns


def _classify_removal(
    text: str, patterns: list[tuple[str, re.Pattern]]
) -> tuple[str | None, str | None]:
    """Check if removed text matches a known removal pattern.

    Returns (category, pattern_string) or (None, None) if unexpected.
    """
    for category, pattern in patterns:
        if pattern.search(text):
            return category, pattern.pattern
    return None, None


# =============================================================================
# Main verification logic
# =============================================================================

def verify_clean(
    input_path: Path,
    output_path: Path,
    config_path: Path | None = None,
) -> VerificationResult:
    """Compare input and output DOCX files, classify every removal.

    Classification uses two layers:
      1. Text-pattern matching from patterns.yaml
      2. Formatting-based detection (italic + editorial colour,
         specifier paragraph/character style, hidden text)

    Args:
        input_path:  Original DOCX before cleaning.
        output_path: Cleaned DOCX after processing.
        config_path: Path to patterns.yaml (auto-detected if None).

    Returns:
        VerificationResult with all removals classified.
    """
    if config_path is None:
        config_path = Path(__file__).parent / "patterns.yaml"
    with open(config_path, "r") as f:
        config = yaml.safe_load(f)

    # Rich extraction for input (text + formatting metadata)
    input_paras = _extract_paragraphs_with_formatting(input_path, config)
    input_texts = [p.text for p in input_paras]

    # Plain text extraction for output
    output_texts = extract_text(output_path)

    removal_patterns = _build_removal_patterns(config)
    preserve_patterns = _build_preserve_patterns(config)

    # Sequence diff — find paragraphs that disappeared
    matcher = difflib.SequenceMatcher(None, input_texts, output_texts)

    removed: list[RemovedParagraph] = []
    for tag, i1, i2, _j1, _j2 in matcher.get_opcodes():
        if tag in ("delete", "replace"):
            for idx in range(i1, i2):
                para_info = input_paras[idx]

                # Layer 0: check if this content should have been preserved
                preserve_match = None
                for ppat in preserve_patterns:
                    if ppat.search(para_info.text):
                        preserve_match = ppat.pattern
                        break

                if preserve_match is not None:
                    removed.append(RemovedParagraph(
                        text=para_info.text,
                        category=PRESERVE_VIOLATION,
                        pattern_matched=preserve_match,
                    ))
                    continue

                # Layer 1: text-pattern match
                category, pattern_str = _classify_removal(
                    para_info.text, removal_patterns
                )

                # Layer 2: formatting-based fallback
                if category is None and para_info.has_editorial_formatting:
                    category = "formatting_based"
                    signals: list[str] = []
                    if para_info.has_editorial_style:
                        signals.append("editorial style")
                    if para_info.is_hidden:
                        signals.append("hidden text")
                    if para_info.is_italic and para_info.has_editorial_color:
                        signals.append("italic + editorial color")
                    pattern_str = "; ".join(signals)

                removed.append(RemovedParagraph(
                    text=para_info.text,
                    category=category,
                    pattern_matched=pattern_str,
                ))

    return VerificationResult(
        input_path=input_path,
        output_path=output_path,
        input_paragraph_count=len(input_texts),
        output_paragraph_count=len(output_texts),
        removed=removed,
    )

