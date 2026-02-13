#!/usr/bin/env python3
"""
SpecCleanse Verification Module

Compares input and output DOCX files to verify that only expected
content (bloat, editorial notes, etc.) was removed and no real
specification content was lost.

Standalone usage:
    python verify.py input.docx output.docx [-v]
"""

import difflib
import re
import shutil
import sys
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
class RemovedParagraph:
    """A paragraph present in input but absent in output."""
    text: str
    category: str | None = None       # e.g. "specifier_note", or None if unexpected
    pattern_matched: str | None = None # the regex that matched


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
        """Removals that matched a known bloat pattern."""
        return [r for r in self.removed if r.category is not None]

    @property
    def unexpected_removals(self) -> list[RemovedParagraph]:
        """Removals that did NOT match any known pattern — red flags."""
        return [r for r in self.removed if r.category is None]

    @property
    def passed(self) -> bool:
        return len(self.unexpected_removals) == 0


# =============================================================================
# Text extraction
# =============================================================================

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
        word_dir = temp_dir / "word"

        # Collect XML parts in a stable order
        xml_files: list[Path] = []
        doc_xml = word_dir / "document.xml"
        if doc_xml.exists():
            xml_files.append(doc_xml)
        xml_files.extend(sorted(word_dir.glob("header*.xml")))
        xml_files.extend(sorted(word_dir.glob("footer*.xml")))

        for xml_path in xml_files:
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

    input_paragraphs = extract_text(input_path)
    output_paragraphs = extract_text(output_path)

    patterns = _build_removal_patterns(config)

    # Sequence diff — find paragraphs that disappeared
    matcher = difflib.SequenceMatcher(None, input_paragraphs, output_paragraphs)

    removed_texts: list[str] = []
    for tag, i1, i2, _j1, _j2 in matcher.get_opcodes():
        if tag in ("delete", "replace"):
            removed_texts.extend(input_paragraphs[i1:i2])

    # Classify each removal
    removed: list[RemovedParagraph] = []
    for text in removed_texts:
        category, pattern_str = _classify_removal(text, patterns)
        removed.append(RemovedParagraph(
            text=text,
            category=category,
            pattern_matched=pattern_str,
        ))

    return VerificationResult(
        input_path=input_path,
        output_path=output_path,
        input_paragraph_count=len(input_paragraphs),
        output_paragraph_count=len(output_paragraphs),
        removed=removed,
    )


# =============================================================================
# Reporting
# =============================================================================

def print_verification(result: VerificationResult, verbose: bool = False):
    """Print a human-readable verification report."""
    print()
    print("=" * 60)
    print("SpecCleanse Verification Report")
    print("=" * 60)
    print()
    print(f"Input:  {result.input_path} ({result.input_paragraph_count} paragraphs)")
    print(f"Output: {result.output_path} ({result.output_paragraph_count} paragraphs)")
    print(f"Paragraphs removed: {len(result.removed)}")
    print()

    expected = result.expected_removals
    unexpected = result.unexpected_removals

    # Category breakdown for expected removals
    if expected:
        category_counts: dict[str, int] = {}
        for r in expected:
            category_counts[r.category] = category_counts.get(r.category, 0) + 1

        print(f"Expected removals ({len(expected)}):")
        for cat, count in sorted(category_counts.items()):
            print(f"  {cat}: {count}")
        print()

    if verbose and expected:
        print("-" * 60)
        print("EXPECTED REMOVALS (matched known patterns):")
        print("-" * 60)
        for i, r in enumerate(expected, 1):
            preview = r.text[:100] + "..." if len(r.text) > 100 else r.text
            preview = preview.replace("\n", " ")
            print(f"  {i}. [{r.category}] \"{preview}\"")
        print()

    if unexpected:
        print("-" * 60)
        print(f"UNEXPECTED REMOVALS ({len(unexpected)}) — review these:")
        print("-" * 60)
        for i, r in enumerate(unexpected, 1):
            preview = r.text[:120] + "..." if len(r.text) > 120 else r.text
            preview = preview.replace("\n", " ")
            print(f"  {i}. \"{preview}\"")
        print()

    if result.passed:
        print("PASS — all removals match known bloat patterns")
    else:
        print(f"WARN — {len(unexpected)} removal(s) did not match any known pattern")
        print("  These may be legitimate spec content. Please review above.")


# =============================================================================
# Standalone CLI
# =============================================================================

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python verify.py <input.docx> <output.docx> [-v]")
        sys.exit(1)

    inp = Path(sys.argv[1])
    out = Path(sys.argv[2])
    verbose = "-v" in sys.argv

    if not inp.exists():
        print(f"Error: {inp} not found", file=sys.stderr)
        sys.exit(1)
    if not out.exists():
        print(f"Error: {out} not found", file=sys.stderr)
        sys.exit(1)

    result = verify_clean(inp, out)
    print_verification(result, verbose=verbose)
    sys.exit(0 if result.passed else 1)
