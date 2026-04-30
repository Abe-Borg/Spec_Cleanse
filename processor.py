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

# Namespaces
NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}

W_NS = NAMESPACES["w"]
W = f"{{{W_NS}}}"


def repack_docx(unpacked_dir: Path, output_path: Path):
    """Repack an unpacked directory into a DOCX (ZIP) file."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
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


@dataclass
class ProcessingResult:
    """Results from processing a document."""
    input_path: Path
    output_path: Path
    detections: list[Detection] = field(default_factory=list)
    removed_count: int = 0
    preserved_count: int = 0
    errors: list[str] = field(default_factory=list)
    
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
    
    def __init__(self, engine: DetectionEngine, verbose: bool = False, dry_run: bool = False):
        self.engine = engine
        self.verbose = verbose
        self.dry_run = dry_run
        self._temp_dir: Optional[Path] = None
    
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

                # Process main document
                doc_path = unpacked_dir / "word" / "document.xml"
                if doc_path.exists():
                    doc_detections = self._process_xml_file(doc_path)
                    result.detections.extend(doc_detections)

                # Process headers
                for header_path in (unpacked_dir / "word").glob("header*.xml"):
                    header_detections = self._process_xml_file(header_path)
                    result.detections.extend(header_detections)

                # Process footers
                for footer_path in (unpacked_dir / "word").glob("footer*.xml"):
                    footer_detections = self._process_xml_file(footer_path)
                    result.detections.extend(footer_detections)

                # Process footnotes and endnotes (kept in sync with verify.py)
                for extra in ("footnotes.xml", "endnotes.xml"):
                    extra_path = unpacked_dir / "word" / extra
                    if extra_path.exists():
                        extra_detections = self._process_xml_file(extra_path)
                        result.detections.extend(extra_detections)

                # Count results
                for d in result.detections:
                    if d.content_type == ContentType.PRESERVE:
                        result.preserved_count += 1
                    else:
                        result.removed_count += 1

                # Repack
                if not self.dry_run:
                    repack_docx(unpacked_dir, output_path)

            finally:
                # Cleanup temp directory
                if self._temp_dir and self._temp_dir.exists():
                    shutil.rmtree(self._temp_dir)

        except Exception as e:
            result.errors.append(f"Processing error: {str(e)}")

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

        # Parse XML preserving whitespace
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(xml_path), parser)
        root = tree.getroot()

        # Track elements to remove (can't modify during iteration)
        elements_to_remove = []

        # Process paragraphs
        for para in root.iter(f"{W}p"):
            para_detections = self._process_paragraph(para)
            detections.extend(para_detections)

            # Check if entire paragraph should be removed
            if self._should_remove_paragraph(para, para_detections):
                elements_to_remove.append(("paragraph", para))
            else:
                # Paragraph survives — remove individual detected runs
                # (e.g. hidden text runs in a mixed-content paragraph)
                runs = set(para.iter(f"{W}r"))
                for d in para_detections:
                    if d.element in runs and d.content_type != ContentType.PRESERVE:
                        elements_to_remove.append(("run", d.element))

        # Process runs not in paragraphs (rare but possible in headers/footers)
        for run in root.iter(f"{W}r"):
            if run.getparent().tag != f"{W}p":
                run_text = self._get_run_text(run)
                run_detections = self.engine.detect_in_element(run, run_text)
                detections.extend(run_detections)

                if self.engine.should_remove(run_detections):
                    elements_to_remove.append(("run", run))

        if not self.dry_run:
            # Remove detected elements
            for elem_type, elem in elements_to_remove:
                self._remove_element(elem, elem_type)

            # Write back
            tree.write(str(xml_path), xml_declaration=True, encoding="UTF-8", standalone=True)

        return detections
    
    def _process_paragraph(self, para: etree._Element) -> list[Detection]:
        """Process a paragraph, detecting removable content."""
        detections = []
        
        # Get full paragraph text for context
        para_text = self._get_paragraph_text(para)
        
        # Check paragraph-level detection (for style-based and full-paragraph patterns)
        para_detections = self.engine.detect_in_element(para, para_text)
        for d in para_detections:
            d.parent_paragraph = para
        detections.extend(para_detections)
        
        # Only check individual runs if paragraph wasn't already fully detected
        # This avoids duplicate detections
        para_should_remove = self.engine.should_remove(para_detections)
        para_preserve = any(d.content_type == ContentType.PRESERVE for d in para_detections)
        
        if not para_should_remove and not para_preserve:
            # Check each run for run-specific detection (hidden text, formatting)
            for run in para.iter(f"{W}r"):
                run_text = self._get_run_text(run)
                if not run_text.strip():
                    continue
                    
                run_detections = self.engine.detect_in_element(run, run_text)
                for d in run_detections:
                    d.parent_paragraph = para
                detections.extend(run_detections)
        
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
        
        # Check for paragraph-level detection that meets removal threshold
        para_detections = [d for d in detections if d.element == para]
        if any(d.confidence >= 0.5 for d in para_detections):
            return True
        
        # Check if all runs are detected
        runs = list(para.iter(f"{W}r"))
        if not runs:
            return False
        
        run_detections = [d for d in detections if d.element in runs]
        detected_runs = set(d.element for d in run_detections if d.confidence >= 0.5)
        
        # Only remove paragraph if ALL runs with text are detected
        runs_with_text = [r for r in runs if self._get_run_text(r).strip()]
        if runs_with_text and all(r in detected_runs for r in runs_with_text):
            return True
        
        return False
    
    def _remove_element(self, elem: etree._Element, elem_type: str):
        """Remove an element from the document."""
        parent = elem.getparent()
        if parent is not None:
            # Handle tail text (text after element but before next element)
            if elem.tail:
                prev = elem.getprevious()
                if prev is not None:
                    prev.tail = (prev.tail or "") + elem.tail
                else:
                    parent.text = (parent.text or "") + elem.tail
            
            parent.remove(elem)
            
            if self.verbose:
                print(f"  Removed {elem_type}")
    
    def _get_paragraph_text(self, para: etree._Element) -> str:
        """Extract all text from a paragraph."""
        texts = []
        for t in para.iter(f"{W}t"):
            if t.text:
                texts.append(t.text)
        return "".join(texts)
    
    def _get_run_text(self, run: etree._Element) -> str:
        """Extract text from a run."""
        texts = []
        for t in run.iter(f"{W}t"):
            if t.text:
                texts.append(t.text)
        return "".join(texts)

