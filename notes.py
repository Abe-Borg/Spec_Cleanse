"""
SpecCleanse Specifier-Note Inspection Module

Extracts specifier notes from a DOCX with stable location information so
the GUI can present them to the user one at a time, and removes only the
notes the user chooses to delete.

The DOCX file is manipulated directly as a ZIP of XML files — Word does
NOT need to be open, and indeed must not be (Word holds a write lock on
open files on Windows).
"""

from __future__ import annotations

import re
import shutil
import tempfile
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

from detection import (
    ContentType,
    Detection,
    DetectionEngine,
    SpecifierNoteDetector,
)
from processor import repack_docx

# Word namespace
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

# Paragraph styles that look like a heading or section/part/article marker.
# Used purely to provide friendly "nearby heading" context to the user.
_HEADING_STYLE_RE = re.compile(r"^(heading|title|section|part|article)", re.IGNORECASE)
# Heading-text detection is intentionally case-sensitive: real spec section
# headers use uppercase ("SECTION 238126", "PART 1 - GENERAL"), while ordinary
# prose like "Section includes:" is mixed case and must NOT be classified as
# a heading.
_HEADING_TEXT_RE = re.compile(r"^\s*(SECTION|PART|ARTICLE)\s+[\dA-Z]")


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------


@dataclass
class SpecifierNote:
    """A single specifier note with enough info to relocate and remove it."""

    note_id: str               # stable id (UUID) used by the GUI
    text: str                  # the note's text content
    xml_file: str              # e.g. "document.xml", "header1.xml"
    paragraph_index: int       # 0-based index of <w:p> within that XML file
    run_index: int | None      # 0-based <w:r> index, or None for whole paragraph
    is_whole_paragraph: bool   # True when the entire paragraph is the note
    nearby_heading: str        # nearest preceding heading text (best effort)
    confidence: float          # detection confidence (0.0–1.0)
    reason: str                # human-readable detection reason

    @property
    def location_label(self) -> str:
        """Human-readable location, e.g. 'Main body, ¶47'."""
        part = _friendly_part_name(self.xml_file)
        label = f"{part}, ¶{self.paragraph_index + 1}"
        if self.run_index is not None and not self.is_whole_paragraph:
            label += f" (run {self.run_index + 1})"
        if self.nearby_heading:
            heading = self.nearby_heading.strip()
            if len(heading) > 60:
                heading = heading[:60] + "…"
            label += f" — under {heading}"
        return label


@dataclass
class RemovalReport:
    """Result of remove_specifier_notes_at_locations."""

    output_path: Path
    removed_count: int = 0
    skipped: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def success(self) -> bool:
        return not self.errors


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def extract_specifier_notes(
    input_path: Path,
    config: dict,
) -> list[SpecifierNote]:
    """Find every specifier note in the DOCX with stable location info.

    Only ContentType.SPECIFIER_NOTE detections are returned — copyright
    boilerplate, hidden-text runs, SpecAgent watermarks, and editorial
    artifacts are intentionally excluded so the user-facing review list
    stays focused on what they asked for.

    The DOCX is unpacked into a temp directory, scanned, and the temp
    directory is removed before returning.
    """
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    engine = DetectionEngine(config)
    note_detector = _find_specifier_note_detector(engine)

    temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_notes_"))
    try:
        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(temp_dir)

        word_dir = temp_dir / "word"
        notes: list[SpecifierNote] = []

        for xml_path in _collect_xml_files(word_dir):
            notes.extend(
                _scan_xml_for_notes(xml_path, engine, note_detector, word_dir)
            )

        return notes

    finally:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)


def remove_specifier_notes_at_locations(
    input_path: Path,
    output_path: Path,
    notes: list[SpecifierNote],
) -> RemovalReport:
    """Remove the given specifier notes from input_path, writing to output_path.

    Notes are located by (xml_file, paragraph_index, run_index) — the same
    triple captured during extract_specifier_notes. The original input file
    is never modified; the result is written as a fresh DOCX at output_path.

    Word does not need to be running, and the input file does not need to
    be closed in any application other than processes that hold an
    exclusive write lock on it (Windows-only behaviour).
    """
    report = RemovalReport(output_path=output_path)

    if not input_path.exists():
        report.errors.append(f"Input file not found: {input_path}")
        return report

    if not notes:
        report.errors.append("No notes were selected for removal.")
        return report

    temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_remove_"))
    try:
        unpacked_dir = temp_dir / "unpacked"
        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(unpacked_dir)

        word_dir = unpacked_dir / "word"

        # Group notes by xml_file so we open each file once.
        notes_by_file: dict[str, list[SpecifierNote]] = {}
        for note in notes:
            notes_by_file.setdefault(note.xml_file, []).append(note)

        for xml_file, file_notes in notes_by_file.items():
            xml_path = word_dir / xml_file
            if not xml_path.exists():
                report.errors.append(
                    f"Expected XML part missing from DOCX: {xml_file}"
                )
                continue

            removed, skipped = _remove_notes_in_file(xml_path, file_notes)
            report.removed_count += removed
            report.skipped.extend(skipped)

        if report.errors:
            return report

        output_path.parent.mkdir(parents=True, exist_ok=True)
        repack_docx(unpacked_dir, output_path)

    except Exception as exc:
        report.errors.append(f"Failed to remove notes: {exc}")

    finally:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)

    return report


# ---------------------------------------------------------------------------
# Internals
# ---------------------------------------------------------------------------


def _find_specifier_note_detector(engine: DetectionEngine) -> SpecifierNoteDetector:
    """Pull the SpecifierNoteDetector instance out of the engine."""
    for detector in engine.detectors:
        if isinstance(detector, SpecifierNoteDetector):
            return detector
    raise RuntimeError("DetectionEngine has no SpecifierNoteDetector configured")


def _collect_xml_files(word_dir: Path) -> list[Path]:
    """document.xml + headers + footers + footnotes + endnotes (stable order).

    Kept in sync with the file list walked by processor.py and verify.py.
    """
    xml_files: list[Path] = []
    doc_xml = word_dir / "document.xml"
    if doc_xml.exists():
        xml_files.append(doc_xml)
    xml_files.extend(sorted(word_dir.glob("header*.xml")))
    xml_files.extend(sorted(word_dir.glob("footer*.xml")))
    for extra in ("footnotes.xml", "endnotes.xml"):
        p = word_dir / extra
        if p.exists():
            xml_files.append(p)
    return xml_files


def _friendly_part_name(xml_file: str) -> str:
    """Convert an XML part filename into a label users can reason about."""
    if xml_file == "document.xml":
        return "Main body"
    m = re.match(r"header(\d+)\.xml$", xml_file)
    if m:
        return f"Header {m.group(1)}"
    m = re.match(r"footer(\d+)\.xml$", xml_file)
    if m:
        return f"Footer {m.group(1)}"
    if xml_file == "footnotes.xml":
        return "Footnotes"
    if xml_file == "endnotes.xml":
        return "Endnotes"
    return xml_file


def _get_paragraph_text(para: etree._Element) -> str:
    parts: list[str] = []
    for t in para.iter(f"{W}t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _get_run_text(run: etree._Element) -> str:
    parts: list[str] = []
    for t in run.iter(f"{W}t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _paragraph_style(para: etree._Element) -> str | None:
    ppr = para.find(f"{W}pPr")
    if ppr is None:
        return None
    pstyle = ppr.find(f"{W}pStyle")
    if pstyle is None:
        return None
    return pstyle.get(f"{W}val")


def _is_heading_paragraph(para: etree._Element, text: str) -> bool:
    """Best-effort: does this paragraph look like a heading?"""
    style = _paragraph_style(para)
    if style and _HEADING_STYLE_RE.search(style):
        return True
    return bool(text and _HEADING_TEXT_RE.match(text))


def _scan_xml_for_notes(
    xml_path: Path,
    engine: DetectionEngine,
    note_detector: SpecifierNoteDetector,
    word_dir: Path,
) -> list[SpecifierNote]:
    """Walk one XML part and return every specifier note found."""
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(str(xml_path), parser)
    root = tree.getroot()

    rel_name = xml_path.relative_to(word_dir).as_posix()

    notes: list[SpecifierNote] = []
    nearby_heading = ""

    for p_idx, para in enumerate(root.iter(f"{W}p")):
        para_text = _get_paragraph_text(para)

        # Track headings *before* deciding whether this paragraph itself
        # is a note — a note cannot also serve as the heading context for
        # later paragraphs.
        if _is_heading_paragraph(para, para_text):
            nearby_heading = para_text.strip()

        # Preserve patterns short-circuit: if the paragraph is whitelisted,
        # skip both whole-paragraph and per-run note detection.
        if engine.preserve_detector.detect(para, para_text) is not None:
            continue

        # Run-level scan first so we can tell whether the paragraph is
        # entirely note content or a mix of real text and inline notes.
        runs_with_text: list[tuple[int, etree._Element]] = []
        run_detections: list[tuple[int, etree._Element, Detection]] = []
        for r_idx, run in enumerate(para.iter(f"{W}r")):
            run_text = _get_run_text(run)
            if not run_text.strip():
                continue
            runs_with_text.append((r_idx, run))
            if engine.preserve_detector.detect(run, run_text) is not None:
                continue
            det = note_detector.detect(run, run_text)
            if det is not None:
                run_detections.append((r_idx, run, det))

        whole_para_detection = note_detector.detect(para, para_text)

        if whole_para_detection is None and not run_detections:
            continue

        all_runs_are_notes = (
            bool(runs_with_text)
            and len(run_detections) == len(runs_with_text)
        )

        # Whole-paragraph removal when the entire paragraph is editorial:
        # either every text-bearing run matched, or only paragraph-level
        # signals (style, multi-run text patterns) fired.
        if all_runs_are_notes or (whole_para_detection is not None and not run_detections):
            representative = whole_para_detection or run_detections[0][2]
            notes.append(_build_note(
                detection=representative,
                xml_file=rel_name,
                paragraph_index=p_idx,
                run_index=None,
                is_whole_paragraph=True,
                nearby_heading=nearby_heading,
            ))
            continue

        # Otherwise some runs are real spec content — remove only the
        # offending runs so the surrounding text survives.
        for r_idx, _, det in run_detections:
            notes.append(_build_note(
                detection=det,
                xml_file=rel_name,
                paragraph_index=p_idx,
                run_index=r_idx,
                is_whole_paragraph=False,
                nearby_heading=nearby_heading,
            ))

    return notes


def _build_note(
    *,
    detection: Detection,
    xml_file: str,
    paragraph_index: int,
    run_index: int | None,
    is_whole_paragraph: bool,
    nearby_heading: str,
) -> SpecifierNote:
    if detection.content_type != ContentType.SPECIFIER_NOTE:
        # Defensive: extract should only call this for note detections.
        raise ValueError(
            f"Refusing to build SpecifierNote from {detection.content_type}"
        )
    return SpecifierNote(
        note_id=uuid.uuid4().hex,
        text=detection.text,
        xml_file=xml_file,
        paragraph_index=paragraph_index,
        run_index=run_index,
        is_whole_paragraph=is_whole_paragraph,
        nearby_heading=nearby_heading,
        confidence=detection.confidence,
        reason=detection.reason,
    )


def _remove_notes_in_file(
    xml_path: Path,
    notes: list[SpecifierNote],
) -> tuple[int, list[str]]:
    """Remove all the given notes from one XML part. Returns (removed, skipped)."""
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(str(xml_path), parser)
    root = tree.getroot()

    paragraphs = list(root.iter(f"{W}p"))

    # Resolve each note to a target element first; remove only after we've
    # finished resolving so the index walk isn't disturbed by removals.
    targets: list[etree._Element] = []
    skipped: list[str] = []

    for note in notes:
        if note.paragraph_index >= len(paragraphs):
            skipped.append(
                f"Paragraph {note.paragraph_index + 1} no longer exists in "
                f"{note.xml_file}"
            )
            continue

        para = paragraphs[note.paragraph_index]

        if note.is_whole_paragraph or note.run_index is None:
            targets.append(para)
            continue

        runs = list(para.iter(f"{W}r"))
        if note.run_index >= len(runs):
            skipped.append(
                f"Run {note.run_index + 1} no longer exists in "
                f"paragraph {note.paragraph_index + 1} of {note.xml_file}"
            )
            continue

        targets.append(runs[note.run_index])

    # Deduplicate: removing the same element twice would crash. This can
    # happen if the user selected both a whole-paragraph note and a run
    # note inside the same paragraph (the paragraph removal subsumes the
    # run removal).
    seen: set[int] = set()
    unique_targets: list[etree._Element] = []
    for elem in targets:
        elem_id = id(elem)
        if elem_id in seen:
            continue
        # Skip a run if its enclosing paragraph is also being removed.
        if elem.tag == f"{W}r":
            ancestor = elem.getparent()
            while ancestor is not None and ancestor.tag != f"{W}p":
                ancestor = ancestor.getparent()
            if ancestor is not None and id(ancestor) in seen:
                continue
        seen.add(elem_id)
        unique_targets.append(elem)

    removed = 0
    for elem in unique_targets:
        if _remove_with_tail(elem):
            removed += 1
        else:
            skipped.append("Element had no parent and was skipped")

    tree.write(
        str(xml_path),
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    return removed, skipped


def _remove_with_tail(elem: etree._Element) -> bool:
    """Remove an element while preserving its tail text. Returns False if no parent."""
    parent = elem.getparent()
    if parent is None:
        return False
    if elem.tail:
        prev = elem.getprevious()
        if prev is not None:
            prev.tail = (prev.tail or "") + elem.tail
        else:
            parent.text = (parent.text or "") + elem.tail
    parent.remove(elem)
    return True
