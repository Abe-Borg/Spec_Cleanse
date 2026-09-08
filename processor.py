"""
SpecCleanse Document Processor Module

Handles DOCX file manipulation: unpacking, content removal, repacking.
Uses lxml for direct XML manipulation to preserve all formatting.
"""

import os
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from lxml import etree

from detection import Detection, DetectionEngine, ContentType
from docx_xml import (
    KEEP_ON_STRIP,
    R_TAG,
    T_TAG,
    accept_revisions,
    can_delete_paragraph,
    collect_content_parts,
    cut_spans,
    field_chars_balanced,
    has_embedded_content,
    has_section_properties,
    is_layout_break,
    iter_own_runs,
    iter_paragraphs,
    iter_text_nodes,
    load_styles,
    merge_spans,
    orphaned_range_markers,
    paragraph_text,
    parse_xml,
    remove_comment_parts,
    run_text,
    set_text,
    spans_cover,
    strip_comment_markers,
    strip_text_leaves,
    tidy_spans,
    write_xml,
)


def repack_docx(unpacked_dir: Path, output_path: Path):
    """Repack an unpacked directory into a DOCX (ZIP) file.

    The archive is built beside the destination and moved into place only once
    it is complete, so an interrupted run (the window closed mid-clean, a disk
    filling up) can never leave a truncated file where a document should be.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_name(output_path.name + ".tmp")

    try:
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            content_types_path = unpacked_dir / "[Content_Types].xml"
            if content_types_path.exists():
                zf.write(content_types_path, "[Content_Types].xml")

            for root, dirs, files in os.walk(unpacked_dir):
                for file in files:
                    file_path = Path(root) / file
                    if file_path == content_types_path:
                        continue
                    arcname = file_path.relative_to(unpacked_dir)
                    zf.write(file_path, arcname)

        os.replace(temp_path, output_path)

    finally:
        if temp_path.exists():
            temp_path.unlink()


@dataclass
class ProcessingResult:
    """Results from processing a document."""
    input_path: Path
    output_path: Path
    detections: list[Detection] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    #: Things worth telling the user that did not stop the run — a structure
    #: kept because removing it would have gone beyond removing content.
    #: Deliberately separate from ``errors``, which decide ``success``.
    warnings: list[str] = field(default_factory=list)

    @property
    def success(self) -> bool:
        return len(self.errors) == 0


class DocxProcessor:
    """
    Processes DOCX files to remove unwanted content.
    
    Strategy:
    1. Unpack DOCX (it's a ZIP file)
    2. Parse document.xml and other content XMLs
    3. Walk the document tree, detect removable content
    4. Remove detected elements while preserving structure
    5. Repack into new DOCX
    """
    
    def __init__(
        self,
        engine: DetectionEngine,
        verbose: bool = False,
        dry_run: bool = False,
        strip_revisions: bool = False,
    ):
        self.engine = engine
        self.verbose = verbose
        self.dry_run = dry_run
        # Accept tracked changes and drop comments before detection.  Deleted
        # text lives in w:delText, which no text extractor sees, so without
        # this it survives an ordinary clean along with its revision markup.
        self.strip_revisions = strip_revisions
        self._temp_dir: Optional[Path] = None
        self._warnings: list[str] = []
    
    def process(self, input_path: Path, output_path: Path) -> ProcessingResult:
        """
        Process a DOCX file, removing detected content.

        Args:
            input_path: Path to input DOCX
            output_path: Path for output DOCX

        Returns:
            ProcessingResult with details of what was done
        """
        result = ProcessingResult(input_path=input_path, output_path=output_path)
        self._warnings = []

        try:
            # Validate input
            if not input_path.exists():
                result.errors.append(f"Input file not found: {input_path}")
                return result

            if not self._is_valid_docx(input_path):
                result.errors.append(f"Invalid DOCX file: {input_path}")
                return result

            # Create temp directory for unpacking
            self._temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_"))

            try:
                # Unpack
                unpacked_dir = self._temp_dir / "unpacked"
                self._unpack_docx(input_path, unpacked_dir)

                # Resolve editorial styles against this document's own style
                # table, so display names and w:basedOn chains are matched too.
                self.engine.bind_styles(load_styles(unpacked_dir / "word"))

                if self.strip_revisions and not self.dry_run:
                    remove_comment_parts(unpacked_dir)

                # Process every content-bearing part: document.xml, headers,
                # footers, footnotes, endnotes, glossary (kept in sync with
                # verify.py through docx_xml.collect_content_parts)
                for xml_path in collect_content_parts(unpacked_dir / "word"):
                    result.detections.extend(self._process_xml_file(xml_path))

                # Repack
                if not self.dry_run:
                    repack_docx(unpacked_dir, output_path)

            finally:
                # Cleanup temp directory
                if self._temp_dir and self._temp_dir.exists():
                    shutil.rmtree(self._temp_dir)

        except PermissionError:
            result.errors.append(
                f"Cannot write {output_path.name} — it is open in Word or "
                "read-only.  Close the file and try again."
            )

        except Exception as e:
            result.errors.append(f"Processing error: {str(e)}")

        result.warnings.extend(self._warnings)
        return result
    
    def _is_valid_docx(self, path: Path) -> bool:
        """Check if file is a valid DOCX."""
        try:
            with zipfile.ZipFile(path, 'r') as zf:
                # Must contain word/document.xml
                return "word/document.xml" in zf.namelist()
        except zipfile.BadZipFile:
            return False
    
    def _unpack_docx(self, docx_path: Path, output_dir: Path):
        """Unpack DOCX to directory."""
        with zipfile.ZipFile(docx_path, 'r') as zf:
            zf.extractall(output_dir)
    
    def _process_xml_file(self, xml_path: Path) -> list[Detection]:
        """Process an XML file, returning detections and modifying in place."""
        detections = []

        tree = parse_xml(xml_path)
        root = tree.getroot()

        revisions = 0
        if self.strip_revisions:
            revisions = accept_revisions(root) + strip_comment_markers(root)

        # Targets are collected during the walk and applied afterwards:
        # mutating the tree while iterating it skips elements.
        paragraphs_to_remove: list[etree._Element] = []
        runs_to_remove: list[etree._Element] = []
        redactions: list[tuple[etree._Element, list[tuple[int, int]]]] = []

        for para in iter_paragraphs(root):
            para_detections = self._process_paragraph(para)
            detections.extend(para_detections)

            if self._should_remove_paragraph(para, para_detections):
                paragraphs_to_remove.append(para)
                continue

            spans = self._redaction_spans(para, para_detections)
            if spans is not None:
                if spans:
                    redactions.append((para, spans))
                else:
                    # Nothing but placeholders — the paragraph itself goes.
                    paragraphs_to_remove.append(para)
                    continue

            # Paragraph survives — remove individual detected runs
            # (e.g. hidden text runs in a mixed-content paragraph)
            for run, run_detections in self._group_run_detections(para, para_detections):
                if self.engine.should_remove(run_detections):
                    runs_to_remove.append(run)

        if self.dry_run:
            return detections

        # Redactions first: their spans are offsets into the paragraph text as
        # it stands now.  Run removal then skips anything redaction already
        # detached, and paragraph removal comes last.
        for para, spans in redactions:
            self._redact_spans(para, spans)
        for run in runs_to_remove:
            self._remove_run(run)
        for para in paragraphs_to_remove:
            self._remove_paragraph(para)

        if paragraphs_to_remove or runs_to_remove or redactions or revisions:
            write_xml(tree, xml_path)

        return detections

    def _redaction_spans(
        self, para: etree._Element, detections: list[Detection]
    ) -> list[tuple[int, int]] | None:
        """Character ranges to cut out of a surviving paragraph.

        Returns None when there is nothing to redact, and an empty list when
        the paragraph is nothing but placeholders — in which case the caller
        removes the whole paragraph after all.
        """
        spans = [
            span
            for d in detections
            if d.content_type == ContentType.INLINE_PLACEHOLDER and d.element == para
            for span in d.spans
        ]
        if not spans:
            return None

        text = paragraph_text(para)
        spans = tidy_spans(text, merge_spans(spans))
        if not cut_spans(text, spans).strip():
            return []
        return spans

    def _group_run_detections(
        self, para: etree._Element, detections: list[Detection]
    ) -> list[tuple[etree._Element, list[Detection]]]:
        """Group run-level detections by the run they belong to."""
        own_runs = set(iter_own_runs(para))
        grouped: dict[etree._Element, list[Detection]] = {}
        for d in detections:
            if d.element in own_runs:
                grouped.setdefault(d.element, []).append(d)
        return list(grouped.items())

    def _process_paragraph(self, para: etree._Element) -> list[Detection]:
        """Process a paragraph, detecting removable content."""
        detections = []
        
        # Get full paragraph text for context
        para_text = paragraph_text(para)
        
        # Check paragraph-level detection (for style-based and full-paragraph patterns)
        para_detections = self.engine.detect_in_element(para, para_text)
        detections.extend(para_detections)
        
        # Only check individual runs if paragraph wasn't already fully detected
        # This avoids duplicate detections
        para_should_remove = self.engine.should_remove(para_detections)
        para_preserve = any(d.content_type == ContentType.PRESERVE for d in para_detections)
        
        if not para_should_remove and not para_preserve:
            # Check each run for run-specific detection (hidden text, formatting).
            # Runs inside a nested text box belong to their own paragraph, which
            # is visited separately — walking into them here would detect, count,
            # and remove the same run twice.
            for run in iter_own_runs(para):
                text = run_text(run)
                if not text.strip():
                    continue

                detections.extend(self.engine.detect_in_element(run, text))
        
        return detections
    
    def _should_remove_paragraph(self, para: etree._Element, detections: list[Detection]) -> bool:
        """
        Determine if entire paragraph should be removed.
        
        Rules:
        - If PRESERVE detection exists, don't remove
        - If paragraph-level detection meets threshold, remove
        - If all runs are detected for removal, remove paragraph
        """
        # Check for preserve
        if any(d.content_type == ContentType.PRESERVE for d in detections):
            return False
        
        # Check for paragraph-level detection that meets removal threshold.
        # Inline placeholders are excluded by the engine: they are cut out of
        # the paragraph, not carried off with it.
        para_detections = [d for d in detections if d.element == para]
        if self.engine.should_remove(para_detections):
            return True
        
        # Check if all runs are detected
        runs = list(iter_own_runs(para))
        if not runs:
            return False
        
        run_detections = [d for d in detections if d.element in runs]
        detected_runs = set(d.element for d in run_detections if d.confidence >= 0.5)
        
        # Only remove paragraph if ALL runs with text are detected
        runs_with_text = [r for r in runs if run_text(r).strip()]
        if runs_with_text and all(r in detected_runs for r in runs_with_text):
            return True
        
        return False

    def _remove_paragraph(self, para: etree._Element):
        """Remove a paragraph, deleting the element only when that is safe.

        Deleting a ``w:p`` outright is wrong in four situations, and in each of
        them the paragraph's content is stripped in place instead:

        * a field starts inside it and ends later — unbalanced ``w:fldChar``
          structure is the kind of damage Word refuses to open;
        * it carries a ``w:sectPr`` — deleting it merges that section into the
          next and takes its headers, footers and page setup with it;
        * it holds a picture, embedded object, field, or note reference;
        * its parent (a table cell, header, footer, footnote, text box) would
          be left with no block-level content, which Word reports as
          "unreadable content".
        """
        keep_reason = None
        if not field_chars_balanced(para):
            keep_reason = "unbalanced field"
        elif has_section_properties(para):
            keep_reason = "section break"
        elif has_embedded_content(para):
            keep_reason = "embedded content"
        elif not can_delete_paragraph(para):
            keep_reason = "container needs block content"

        if keep_reason is not None:
            self._strip_paragraph_content(para)
            if self.verbose:
                print(f"  Emptied paragraph ({keep_reason})")
            return

        self._relocate_orphaned_markers(para)

        parent = para.getparent()
        if parent is not None:
            parent.remove(para)
            if self.verbose:
                print("  Removed paragraph")

    def _remove_run(self, run: etree._Element):
        """Remove a run, keeping structure the document depends on.

        A run that anchors a field, a picture, or a footnote reference is
        blanked rather than deleted: its text goes, its structure stays.
        """
        if not field_chars_balanced(run) or has_embedded_content(run):
            strip_text_leaves(run)
            if self.verbose:
                print("  Emptied run (embedded content)")
            return

        parent = run.getparent()
        if parent is not None:
            parent.remove(run)
            if self.verbose:
                print("  Removed run")

    def _redact_spans(self, para: etree._Element, spans: list[tuple[int, int]]):
        """Cut character ranges out of a paragraph's text, in place.

        The paragraph's text nodes are walked in document order with a running
        offset, so each span is mapped back onto the node it came from.  A
        ``w:t`` is edited; a separator that renders as a character — a tab, a
        break, a non-breaking hyphen — is removed whole when the redaction
        covers it, because it has no partial state to edit.

        Advancing the offset past a separator without removing it is what left
        ``Provide -units.`` behind: the hyphen inside ``[Verify quantity-with
        Owner]`` was counted for the offsets and then never touched.

        Runs left holding nothing are removed; runs that still carry a picture
        or a field are kept.
        """
        touched_runs: list[etree._Element] = []
        covered_separators: list[etree._Element] = []
        offset = 0
        # Snapshot before the first edit: w:t nodes are rewritten as the walk
        # goes, so a preview taken later would quote half-redacted text.
        preview = self._preview(para)

        for node, text in list(iter_text_nodes(para)):
            start = offset
            offset += len(text)

            if node.tag != T_TAG:
                if not spans_cover(spans, start, offset):
                    continue
                if is_layout_break(node):
                    # A page or column break renders as "\n" and so can fall
                    # inside a match, but it is page setup rather than content.
                    # Cleaning removes content; reflowing the document from
                    # here is a bigger claim than any pattern makes.
                    self._warn(
                        "Kept a page or column break that fell inside removed "
                        f"text: \"{preview}\""
                    )
                    continue
                covered_separators.append(node)
                run = self._owning_run(node, para)
                if run is not None:
                    touched_runs.append(run)
                continue

            local = [
                (max(span_start - start, 0), min(span_end - start, len(text)))
                for span_start, span_end in spans
                if span_end > start and span_start < offset
            ]
            local = [(s, e) for s, e in local if e > s]
            if not local:
                continue

            run = self._owning_run(node, para)
            if run is not None:
                touched_runs.append(run)

            remaining = cut_spans(text, local)
            if remaining:
                set_text(node, remaining)
            else:
                node.getparent().remove(node)

        # Before the empty-run sweep below, so a run left holding nothing but
        # a redacted separator is seen as empty.
        for node in covered_separators:
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)

        for run in touched_runs:
            parent = run.getparent()
            if parent is None or has_embedded_content(run) or run_text(run):
                continue
            parent.remove(run)

    def _warn(self, message: str) -> None:
        """Record something the user should know that did not stop the run."""
        if message not in self._warnings:
            self._warnings.append(message)

    @staticmethod
    def _preview(para: etree._Element, width: int = 60) -> str:
        """A one-line excerpt of a paragraph, for a warning message."""
        flat = " ".join(paragraph_text(para).split())
        return flat if len(flat) <= width else flat[:width] + "..."

    def _owning_run(
        self, node: etree._Element, para: etree._Element
    ) -> Optional[etree._Element]:
        """Return the ``w:r`` a text node belongs to, if any."""
        parent = node.getparent()
        while parent is not None and parent is not para:
            if parent.tag == R_TAG:
                return parent
            parent = parent.getparent()
        return None

    def _strip_paragraph_content(self, para: etree._Element):
        """Empty a paragraph in place, keeping properties and anchors.

        Paragraph properties, range markers, and any child holding embedded
        content are kept; everything else goes.  Kept children lose their text
        so no editorial content survives the strip.
        """
        for child in list(para):
            if child.tag in KEEP_ON_STRIP:
                continue
            if has_embedded_content(child):
                strip_text_leaves(child)
                continue
            para.remove(child)

    def _relocate_orphaned_markers(self, para: etree._Element):
        """Move half-open bookmark/comment ranges out of a doomed paragraph.

        A ``w:bookmarkStart`` whose ``w:bookmarkEnd`` lives further down the
        document is legal beside the paragraph as well as inside it, so the
        range survives the removal intact instead of being left half-open.
        """
        orphans = orphaned_range_markers(para)
        if not orphans:
            return

        parent = para.getparent()
        if parent is None:
            return

        index = parent.index(para)
        for marker in orphans:
            marker_parent = marker.getparent()
            if marker_parent is not None:
                marker_parent.remove(marker)
            parent.insert(index, marker)
            index += 1
