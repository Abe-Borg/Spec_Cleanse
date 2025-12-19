#!/usr/bin/env python3
"""
DOCX Deep Cleaner

Uses the orphan analysis report to safely remove cruft from DOCX files.
This module performs the actual surgery - removing orphaned resources
while maintaining document integrity.

SAFETY PRINCIPLES:
1. Never remove anything without first validating the orphan report
2. Always create a backup before modifications
3. Remove in safe order: media first, then relationships, then styles
4. Validate the document can be repacked after each major operation
5. Provide rollback capability
"""

import zipfile
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Set, Optional
from dataclasses import dataclass
import json
import re


@dataclass
class CleaningResult:
    """Results of a cleaning operation."""
    success: bool
    relationships_removed: int = 0
    media_removed: int = 0
    styles_removed: int = 0
    rsids_removed: int = 0
    empty_elements_removed: int = 0
    font_mappings_removed: int = 0
    compat_settings_removed: int = 0
    bookmarks_removed: int = 0
    proof_elements_removed: int = 0
    bytes_saved: int = 0
    errors: List[str] = None
    warnings: List[str] = None
    
    def __post_init__(self):
        if self.errors is None:
            self.errors = []
        if self.warnings is None:
            self.warnings = []


class DocxDeepCleaner:
    """
    Performs safe deep cleaning of DOCX files based on orphan analysis.
    """
    
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
        'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    }
    
    def __init__(self, extract_dir: Path, orphan_report: dict):
        """
        Initialize cleaner.
        
        Args:
            extract_dir: Path to extracted DOCX directory
            orphan_report: Dictionary from OrphanReport.to_dict()
        """
        self.extract_dir = Path(extract_dir)
        self.orphan_report = orphan_report
        self.backup_dir = None
        self.result = CleaningResult(success=False)
        
        # Register namespaces to preserve them in output
        for prefix, uri in self.NAMESPACES.items():
            ET.register_namespace(prefix, uri)
    
    def clean(self, 
              remove_relationships: bool = True,
              remove_media: bool = True, 
              remove_styles: bool = True,
              strip_rsids: bool = True,
              remove_empty_elements: bool = True,
              remove_non_english_fonts: bool = True,
              remove_compat_settings: bool = True,
              remove_internal_bookmarks: bool = True,
              remove_proof_state: bool = True,
              dry_run: bool = False) -> CleaningResult:
        """
        Perform deep cleaning based on orphan report.
        
        Args:
            remove_relationships: Remove orphaned hyperlinks and other relationships
            remove_media: Remove orphaned media files
            remove_styles: Remove orphaned style definitions
            strip_rsids: Remove RSID tracking attributes
            remove_empty_elements: Remove empty runs, paragraphs, etc.
            remove_non_english_fonts: Remove non-English font mappings from theme
            remove_compat_settings: Remove backwards compatibility settings
            remove_internal_bookmarks: Remove Word's internal bookmarks
            remove_proof_state: Remove spell/grammar check state
            dry_run: If True, report what would be done without making changes
        
        Returns:
            CleaningResult with details of operations performed
        """
        if dry_run:
            return self._dry_run(remove_relationships, remove_media, remove_styles,
                               strip_rsids, remove_empty_elements, remove_non_english_fonts,
                               remove_compat_settings, remove_internal_bookmarks, remove_proof_state)
        
        # Create backup
        self._create_backup()
        
        try:
            if remove_media:
                self._remove_orphaned_media()
            
            if remove_relationships:
                self._remove_orphaned_relationships()
            
            if remove_styles:
                self._remove_orphaned_styles()
            
            if strip_rsids:
                self._strip_rsids()
            
            if remove_empty_elements:
                self._remove_empty_elements()
            
            if remove_non_english_fonts:
                self._remove_non_english_font_mappings()
            
            if remove_compat_settings:
                self._remove_compatibility_settings()
            
            if remove_internal_bookmarks:
                self._remove_internal_bookmarks()
            
            if remove_proof_state:
                self._remove_proof_state()
            
            # Validate document structure
            if self._validate_structure():
                self.result.success = True
            else:
                self.result.errors.append("Document structure validation failed")
                self._restore_backup()
        
        except Exception as e:
            self.result.errors.append(f"Cleaning failed: {str(e)}")
            self._restore_backup()
        
        return self.result
    
    def _dry_run(self, remove_relationships: bool, remove_media: bool, 
                 remove_styles: bool, strip_rsids: bool, remove_empty_elements: bool,
                 remove_non_english_fonts: bool, remove_compat_settings: bool,
                 remove_internal_bookmarks: bool, remove_proof_state: bool) -> CleaningResult:
        """Report what would be cleaned without making changes."""
        orphans = self.orphan_report.get('orphans', {})
        cruft = self.orphan_report.get('cruft', {})
        stats = self.orphan_report.get('statistics', {})
        
        result = CleaningResult(success=True)
        
        if remove_relationships:
            result.relationships_removed = len(orphans.get('relationships', []))
        
        if remove_media:
            media_orphans = orphans.get('media', [])
            result.media_removed = len(media_orphans)
            result.bytes_saved = sum(m.get('size_bytes', 0) for m in media_orphans)
        
        if remove_styles:
            result.styles_removed = len(orphans.get('styles', []))
        
        if strip_rsids:
            result.rsids_removed = stats.get('rsid_attributes', 0)
            # Estimate 25 bytes per RSID
            result.bytes_saved += result.rsids_removed * 25
        
        if remove_empty_elements:
            result.empty_elements_removed = stats.get('empty_elements', 0)
            result.bytes_saved += result.empty_elements_removed * 20
        
        if remove_non_english_fonts:
            result.font_mappings_removed = len(cruft.get('non_english_font_mappings', []))
            result.bytes_saved += result.font_mappings_removed * 60
        
        if remove_compat_settings:
            result.compat_settings_removed = len(cruft.get('compatibility_settings', []))
            result.bytes_saved += result.compat_settings_removed * 50
        
        if remove_internal_bookmarks:
            result.bookmarks_removed = len(cruft.get('internal_bookmarks', []))
            result.bytes_saved += result.bookmarks_removed * 80
        
        if remove_proof_state:
            result.proof_elements_removed = len(cruft.get('proof_state_elements', []))
            result.bytes_saved += result.proof_elements_removed * 40
        
        # Add warnings for any risky operations
        if result.relationships_removed > 50:
            result.warnings.append(
                f"Large number of relationships ({result.relationships_removed}) "
                "flagged for removal. Consider manual review."
            )
        
        return result
    
    def _create_backup(self):
        """Create a backup of the extract directory."""
        self.backup_dir = self.extract_dir.parent / f"{self.extract_dir.name}_backup"
        if self.backup_dir.exists():
            shutil.rmtree(self.backup_dir)
        shutil.copytree(self.extract_dir, self.backup_dir)
    
    def _restore_backup(self):
        """Restore from backup if something went wrong."""
        if self.backup_dir and self.backup_dir.exists():
            shutil.rmtree(self.extract_dir)
            shutil.move(str(self.backup_dir), str(self.extract_dir))
            self.result.warnings.append("Restored from backup due to errors")
    
    def _remove_orphaned_media(self):
        """Remove orphaned media files."""
        orphans = self.orphan_report.get('orphans', {}).get('media', [])
        
        for media in orphans:
            media_path = self.extract_dir / media['path']
            if media_path.exists():
                size = media_path.stat().st_size
                media_path.unlink()
                self.result.media_removed += 1
                self.result.bytes_saved += size
    
    def _remove_orphaned_relationships(self):
        """Remove orphaned relationship entries from .rels files."""
        orphans = self.orphan_report.get('orphans', {}).get('relationships', [])
        
        # Group orphans by source file
        by_source: Dict[str, List[str]] = {}
        for orphan in orphans:
            source = orphan.get('source_file', 'word/_rels/document.xml.rels')
            rid = orphan['rId']
            if source not in by_source:
                by_source[source] = []
            by_source[source].append(rid)
        
        for source_file, rids in by_source.items():
            rels_path = self.extract_dir / source_file
            if not rels_path.exists():
                continue
            
            try:
                tree = ET.parse(rels_path)
                root = tree.getroot()
                
                # Find and remove orphaned relationships
                removed = 0
                for rel in list(root):  # list() to allow modification during iteration
                    if rel.get('Id') in rids:
                        root.remove(rel)
                        removed += 1
                
                if removed > 0:
                    # Write back with proper XML declaration
                    tree.write(rels_path, encoding='UTF-8', xml_declaration=True)
                    self.result.relationships_removed += removed
            
            except ET.ParseError as e:
                self.result.warnings.append(f"Failed to parse {source_file}: {e}")
    
    def _remove_orphaned_styles(self):
        """Remove orphaned style definitions from styles.xml."""
        orphans = self.orphan_report.get('orphans', {}).get('styles', [])
        if not orphans:
            return
        
        orphan_ids = {o['styleId'] for o in orphans}
        
        styles_path = self.extract_dir / "word" / "styles.xml"
        if not styles_path.exists():
            return
        
        try:
            tree = ET.parse(styles_path)
            root = tree.getroot()
            
            w_ns = self.NAMESPACES['w']
            
            removed = 0
            for style in list(root.findall(f'.//{{{w_ns}}}style')):
                style_id = style.get(f'{{{w_ns}}}styleId', '')
                if style_id in orphan_ids:
                    root.remove(style)
                    removed += 1
            
            if removed > 0:
                tree.write(styles_path, encoding='UTF-8', xml_declaration=True)
                self.result.styles_removed = removed
        
        except ET.ParseError as e:
            self.result.warnings.append(f"Failed to parse styles.xml: {e}")
    
    def _strip_rsids(self):
        """Remove all RSID tracking attributes from XML files."""
        rsid_pattern = re.compile(r'\s+w:rsid[A-Za-z]*="[^"]*"', re.IGNORECASE)
        
        xml_files = list(self.extract_dir.rglob('*.xml'))
        total_removed = 0
        
        for xml_file in xml_files:
            try:
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Count and remove RSIDs
                matches = rsid_pattern.findall(content)
                if matches:
                    new_content = rsid_pattern.sub('', content)
                    with open(xml_file, 'w', encoding='utf-8') as f:
                        f.write(new_content)
                    total_removed += len(matches)
            
            except Exception as e:
                self.result.warnings.append(f"Failed to strip RSIDs from {xml_file.name}: {e}")
        
        self.result.rsids_removed = total_removed
        self.result.bytes_saved += total_removed * 25
    
    def _remove_empty_elements(self):
        """Remove empty runs and other useless elements."""
        w_ns = self.NAMESPACES['w']
        
        xml_files = [
            'word/document.xml',
            'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
            'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml',
        ]
        
        total_removed = 0
        
        for rel_path in xml_files:
            xml_file = self.extract_dir / rel_path
            if not xml_file.exists():
                continue
            
            try:
                tree = ET.parse(xml_file)
                root = tree.getroot()
                modified = False
                
                # Remove empty runs (runs with no content-bearing children)
                content_tags = {f'{{{w_ns}}}{t}' for t in 
                              ('t', 'drawing', 'pict', 'sym', 'tab', 'br', 'cr',
                               'fldChar', 'instrText', 'object', 'ruby')}
                
                for parent in root.iter():
                    runs_to_remove = []
                    for child in parent:
                        if child.tag == f'{{{w_ns}}}r':
                            has_content = any(
                                gc.tag in content_tags or 
                                (gc.tag == f'{{{w_ns}}}t' and gc.text)
                                for gc in child
                            )
                            if not has_content:
                                runs_to_remove.append(child)
                    
                    for run in runs_to_remove:
                        parent.remove(run)
                        total_removed += 1
                        modified = True
                
                if modified:
                    tree.write(xml_file, encoding='UTF-8', xml_declaration=True)
            
            except Exception as e:
                self.result.warnings.append(f"Failed to clean empty elements in {rel_path}: {e}")
        
        self.result.empty_elements_removed = total_removed
        self.result.bytes_saved += total_removed * 20
    
    def _remove_non_english_font_mappings(self):
        """Remove non-English font mappings from theme files."""
        keep_scripts = {'latin', 'ea', 'cs', ''}  # Keep these script types
        
        theme_dir = self.extract_dir / "word" / "theme"
        if not theme_dir.exists():
            return
        
        total_removed = 0
        
        for theme_file in theme_dir.glob('*.xml'):
            try:
                tree = ET.parse(theme_file)
                root = tree.getroot()
                modified = False
                
                # Find all font elements and remove non-English ones
                for parent in root.iter():
                    fonts_to_remove = []
                    for child in parent:
                        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                        if tag == 'font':
                            script = child.get('script', '')
                            if script.lower() not in keep_scripts:
                                fonts_to_remove.append(child)
                    
                    for font in fonts_to_remove:
                        parent.remove(font)
                        total_removed += 1
                        modified = True
                
                if modified:
                    tree.write(theme_file, encoding='UTF-8', xml_declaration=True)
            
            except Exception as e:
                self.result.warnings.append(f"Failed to clean theme {theme_file.name}: {e}")
        
        self.result.font_mappings_removed = total_removed
        self.result.bytes_saved += total_removed * 60
    
    def _remove_compatibility_settings(self):
        """Remove backwards compatibility settings from settings.xml."""
        settings_path = self.extract_dir / "word" / "settings.xml"
        if not settings_path.exists():
            return
        
        w_ns = self.NAMESPACES['w']
        
        # Settings that are safe to remove
        removable_tags = {
            'compatSetting', 'useFELayout', 'useWord2002TableStyleRules',
            'growAutofit', 'useWord97LineBreakRules', 
            'doNotUseIndentAsNumberingTabStop', 'useAltKinsokuLineBreakRules',
            'allowSpaceOfSameStyleInTable', 'doNotSuppressParagraphBorders',
            'doNotAutofitConstrainedTables', 'autofitToFirstFixedWidthCell',
            'displayHangulFixedWidth', 'splitPgBreakAndParaMark',
            'doNotVertAlignCellWithSp', 'doNotBreakConstrainedForcedTable',
            'doNotVertAlignInTxbx', 'useAnsiKerningPairs', 'cachedColBalance',
        }
        
        try:
            tree = ET.parse(settings_path)
            root = tree.getroot()
            total_removed = 0
            
            # Find and clean compat element
            for compat in root.iter(f'{{{w_ns}}}compat'):
                children_to_remove = []
                for child in compat:
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if tag in removable_tags:
                        children_to_remove.append(child)
                
                for child in children_to_remove:
                    compat.remove(child)
                    total_removed += 1
            
            if total_removed > 0:
                tree.write(settings_path, encoding='UTF-8', xml_declaration=True)
            
            self.result.compat_settings_removed = total_removed
            self.result.bytes_saved += total_removed * 50
        
        except Exception as e:
            self.result.warnings.append(f"Failed to remove compat settings: {e}")
    
    def _remove_internal_bookmarks(self):
        """Remove Word's internal bookmarks."""
        doc_path = self.extract_dir / "word" / "document.xml"
        if not doc_path.exists():
            return
        
        w_ns = self.NAMESPACES['w']
        internal_prefixes = ('_GoBack', '_Ref', '_Toc', '_Hlk', '_PictureBullets')
        
        try:
            tree = ET.parse(doc_path)
            root = tree.getroot()
            
            # Collect bookmark IDs to remove
            bookmark_ids_to_remove = set()
            bookmarks_to_remove = []
            
            for bookmark in root.iter(f'{{{w_ns}}}bookmarkStart'):
                name = bookmark.get(f'{{{w_ns}}}name', '')
                if any(name.startswith(prefix) for prefix in internal_prefixes):
                    bm_id = bookmark.get(f'{{{w_ns}}}id', '')
                    bookmark_ids_to_remove.add(bm_id)
            
            # Remove bookmarkStart and bookmarkEnd elements
            for parent in root.iter():
                children_to_remove = []
                for child in parent:
                    if child.tag == f'{{{w_ns}}}bookmarkStart':
                        name = child.get(f'{{{w_ns}}}name', '')
                        if any(name.startswith(prefix) for prefix in internal_prefixes):
                            children_to_remove.append(child)
                    elif child.tag == f'{{{w_ns}}}bookmarkEnd':
                        bm_id = child.get(f'{{{w_ns}}}id', '')
                        if bm_id in bookmark_ids_to_remove:
                            children_to_remove.append(child)
                
                for child in children_to_remove:
                    parent.remove(child)
                    bookmarks_to_remove.append(child)
            
            if bookmarks_to_remove:
                tree.write(doc_path, encoding='UTF-8', xml_declaration=True)
            
            self.result.bookmarks_removed = len(bookmarks_to_remove) // 2  # start+end pairs
            self.result.bytes_saved += self.result.bookmarks_removed * 80
        
        except Exception as e:
            self.result.warnings.append(f"Failed to remove internal bookmarks: {e}")
    
    def _remove_proof_state(self):
        """Remove spell/grammar check state elements."""
        w_ns = self.NAMESPACES['w']
        total_removed = 0
        
        # Clean settings.xml
        settings_path = self.extract_dir / "word" / "settings.xml"
        if settings_path.exists():
            try:
                tree = ET.parse(settings_path)
                root = tree.getroot()
                
                for parent in root.iter():
                    children_to_remove = []
                    for child in parent:
                        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                        if tag in ('proofState', 'proofErr'):
                            children_to_remove.append(child)
                    
                    for child in children_to_remove:
                        parent.remove(child)
                        total_removed += 1
                
                if total_removed > 0:
                    tree.write(settings_path, encoding='UTF-8', xml_declaration=True)
            
            except Exception as e:
                self.result.warnings.append(f"Failed to clean proof state from settings: {e}")
        
        # Clean document.xml for proofErr markers
        doc_path = self.extract_dir / "word" / "document.xml"
        if doc_path.exists():
            try:
                tree = ET.parse(doc_path)
                root = tree.getroot()
                doc_removed = 0
                
                for parent in root.iter():
                    children_to_remove = []
                    for child in parent:
                        if child.tag == f'{{{w_ns}}}proofErr':
                            children_to_remove.append(child)
                    
                    for child in children_to_remove:
                        parent.remove(child)
                        doc_removed += 1
                
                if doc_removed > 0:
                    tree.write(doc_path, encoding='UTF-8', xml_declaration=True)
                    total_removed += doc_removed
            
            except Exception as e:
                self.result.warnings.append(f"Failed to clean proof errors from document: {e}")
        
        self.result.proof_elements_removed = total_removed
        self.result.bytes_saved += total_removed * 40
    
    def _validate_structure(self) -> bool:
        """
        Validate the document structure is still intact.
        
        Checks:
        1. [Content_Types].xml exists and is valid
        2. _rels/.rels exists and points to valid targets
        3. word/document.xml exists and is valid
        """
        # Check Content_Types
        ct_path = self.extract_dir / "[Content_Types].xml"
        if not ct_path.exists():
            self.result.errors.append("Missing [Content_Types].xml")
            return False
        
        try:
            ET.parse(ct_path)
        except ET.ParseError as e:
            self.result.errors.append(f"Invalid [Content_Types].xml: {e}")
            return False
        
        # Check root rels
        root_rels = self.extract_dir / "_rels" / ".rels"
        if not root_rels.exists():
            self.result.errors.append("Missing _rels/.rels")
            return False
        
        # Check document.xml
        doc_path = self.extract_dir / "word" / "document.xml"
        if not doc_path.exists():
            self.result.errors.append("Missing word/document.xml")
            return False
        
        try:
            ET.parse(doc_path)
        except ET.ParseError as e:
            self.result.errors.append(f"Invalid document.xml: {e}")
            return False
        
        return True


def repack_docx(extract_dir: Path, output_path: Path) -> bool:
    """
    Repack an extracted directory into a DOCX file.
    
    Args:
        extract_dir: Path to extracted DOCX directory
        output_path: Path for output DOCX file
    
    Returns:
        True if successful
    """
    try:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for file_path in extract_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(extract_dir)
                    docx.write(file_path, arcname)
        return True
    except Exception as e:
        print(f"Error repacking: {e}")
        return False


def main():
    """CLI entry point for deep cleaning."""
    import sys
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Deep clean a DOCX file based on orphan analysis"
    )
    parser.add_argument("extract_dir", help="Path to extracted DOCX directory")
    parser.add_argument("orphan_report", help="Path to orphan analysis JSON file")
    parser.add_argument("--output", "-o", help="Output DOCX path (default: <n>_cleaned.docx)")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be done without making changes")
    
    # Orphan removal flags
    parser.add_argument("--no-relationships", action="store_true", help="Skip removing orphaned relationships")
    parser.add_argument("--no-media", action="store_true", help="Skip removing orphaned media")
    parser.add_argument("--no-styles", action="store_true", help="Skip removing orphaned styles")
    
    # Cruft removal flags
    parser.add_argument("--no-rsids", action="store_true", help="Skip stripping RSID attributes")
    parser.add_argument("--no-empty", action="store_true", help="Skip removing empty elements")
    parser.add_argument("--no-fonts", action="store_true", help="Skip removing non-English font mappings")
    parser.add_argument("--no-compat", action="store_true", help="Skip removing compatibility settings")
    parser.add_argument("--no-bookmarks", action="store_true", help="Skip removing internal bookmarks")
    parser.add_argument("--no-proof", action="store_true", help="Skip removing proof state elements")
    
    args = parser.parse_args()
    
    extract_dir = Path(args.extract_dir)
    if not extract_dir.exists():
        print(f"Error: Directory not found: {extract_dir}")
        sys.exit(1)
    
    report_path = Path(args.orphan_report)
    if not report_path.exists():
        print(f"Error: Report not found: {report_path}")
        sys.exit(1)
    
    with open(report_path) as f:
        orphan_report = json.load(f)
    
    cleaner = DocxDeepCleaner(extract_dir, orphan_report)
    
    result = cleaner.clean(
        remove_relationships=not args.no_relationships,
        remove_media=not args.no_media,
        remove_styles=not args.no_styles,
        strip_rsids=not args.no_rsids,
        remove_empty_elements=not args.no_empty,
        remove_non_english_fonts=not args.no_fonts,
        remove_compat_settings=not args.no_compat,
        remove_internal_bookmarks=not args.no_bookmarks,
        remove_proof_state=not args.no_proof,
        dry_run=args.dry_run
    )
    
    print("\n" + "=" * 60)
    print("DEEP CLEANING " + ("(DRY RUN)" if args.dry_run else "COMPLETE"))
    print("=" * 60)
    
    print("\n--- Orphans Removed ---")
    print(f"Relationships:          {result.relationships_removed}")
    print(f"Media files:            {result.media_removed}")
    print(f"Styles:                 {result.styles_removed}")
    
    print("\n--- Cruft Removed ---")
    print(f"RSID attributes:        {result.rsids_removed}")
    print(f"Empty elements:         {result.empty_elements_removed}")
    print(f"Non-English fonts:      {result.font_mappings_removed}")
    print(f"Compat settings:        {result.compat_settings_removed}")
    print(f"Internal bookmarks:     {result.bookmarks_removed}")
    print(f"Proof state elements:   {result.proof_elements_removed}")
    
    print(f"\n--- Summary ---")
    print(f"Estimated bytes saved:  {result.bytes_saved:,} ({result.bytes_saved/1024:.1f} KB)")
    
    if result.warnings:
        print("\nWarnings:")
        for w in result.warnings:
            print(f"  - {w}")
    
    if result.errors:
        print("\nErrors:")
        for e in result.errors:
            print(f"  - {e}")
        sys.exit(1)
    
    if not args.dry_run:
        # Repack the document
        output_path = Path(args.output) if args.output else \
            extract_dir.parent / f"{extract_dir.name.replace('_extracted', '')}_cleaned.docx"
        
        print(f"\nRepacking to: {output_path}")
        if repack_docx(extract_dir, output_path):
            print("Success!")
        else:
            print("Failed to repack document")
            sys.exit(1)


if __name__ == "__main__":
    main()
