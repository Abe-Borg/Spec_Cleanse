#!/usr/bin/env python3
"""
DOCX Orphan Analyzer

Extends docx_obliterate.py to detect orphaned resources in DOCX files:
- Relationships (rIds) defined but never referenced
- Media files not linked by any relationship
- Styles defined but never used
- Fonts declared but not referenced
- Numbering definitions without usage
- WebSettings div artifacts from copy/paste

This module computes: ORPHANS = DEFINED - USED
"""

import xml.etree.ElementTree as ET
from pathlib import Path
from dataclasses import dataclass, field
from typing import Set, Dict, List, Optional
import re
import json


@dataclass
class OrphanReport:
    """Structured report of orphaned resources and cleanable cruft."""
    # Orphaned resources
    orphaned_relationships: List[Dict] = field(default_factory=list)
    orphaned_media: List[Dict] = field(default_factory=list)
    orphaned_styles: List[Dict] = field(default_factory=list)
    orphaned_fonts: List[Dict] = field(default_factory=list)
    orphaned_numbering: List[Dict] = field(default_factory=list)
    orphaned_web_divs: List[Dict] = field(default_factory=list)
    
    # Cruft to strip
    rsid_attributes: Dict = field(default_factory=dict)  # {file: count}
    empty_elements: List[Dict] = field(default_factory=list)
    non_english_font_mappings: List[Dict] = field(default_factory=list)
    compatibility_settings: List[Dict] = field(default_factory=list)
    internal_bookmarks: List[Dict] = field(default_factory=list)
    proof_state_elements: List[Dict] = field(default_factory=list)
    
    # Statistics
    total_relationships_defined: int = 0
    total_relationships_used: int = 0
    total_styles_defined: int = 0
    total_styles_used: int = 0
    total_media_files: int = 0
    total_media_referenced: int = 0
    total_rsid_attributes: int = 0
    total_empty_elements: int = 0
    
    estimated_savings_bytes: int = 0
    
    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            'orphans': {
                'relationships': self.orphaned_relationships,
                'media': self.orphaned_media,
                'styles': self.orphaned_styles,
                'fonts': self.orphaned_fonts,
                'numbering': self.orphaned_numbering,
                'web_divs': self.orphaned_web_divs,
            },
            'cruft': {
                'rsid_attributes': self.rsid_attributes,
                'empty_elements': self.empty_elements,
                'non_english_font_mappings': self.non_english_font_mappings,
                'compatibility_settings': self.compatibility_settings,
                'internal_bookmarks': self.internal_bookmarks,
                'proof_state_elements': self.proof_state_elements,
            },
            'statistics': {
                'relationships': {
                    'defined': self.total_relationships_defined,
                    'used': self.total_relationships_used,
                    'orphaned': len(self.orphaned_relationships)
                },
                'styles': {
                    'defined': self.total_styles_defined,
                    'used': self.total_styles_used,
                    'orphaned': len(self.orphaned_styles)
                },
                'media': {
                    'total_files': self.total_media_files,
                    'referenced': self.total_media_referenced,
                    'orphaned': len(self.orphaned_media)
                },
                'rsid_attributes': self.total_rsid_attributes,
                'empty_elements': self.total_empty_elements,
            },
            'estimated_savings_bytes': self.estimated_savings_bytes
        }
    
    def to_json(self, indent: int = 2) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict(), indent=indent)
    
    def to_yaml_manifest(self) -> str:
        """Generate YAML-style cleanup manifest."""
        lines = ["# DOCX Cleanup Manifest", "# Safe to remove the following orphaned resources and cruft", ""]
        
        if self.orphaned_relationships:
            lines.append("orphaned_relationships:")
            for rel in self.orphaned_relationships:
                lines.append(f"  - rId: {rel['rId']}")
                lines.append(f"    target: {rel['target']}")
                lines.append(f"    type: {rel['type']}")
                lines.append(f"    reason: \"{rel['reason']}\"")
                lines.append("")
        
        if self.orphaned_media:
            lines.append("orphaned_media:")
            for media in self.orphaned_media:
                lines.append(f"  - path: {media['path']}")
                lines.append(f"    size_bytes: {media['size_bytes']}")
                lines.append(f"    reason: \"{media['reason']}\"")
                lines.append("")
        
        if self.orphaned_styles:
            lines.append("orphaned_styles:")
            for style in self.orphaned_styles:
                lines.append(f"  - styleId: {style['styleId']}")
                lines.append(f"    name: {style.get('name', 'N/A')}")
                lines.append(f"    type: {style['type']}")
                lines.append(f"    reason: \"{style['reason']}\"")
                lines.append("")
        
        if self.orphaned_web_divs:
            lines.append("orphaned_web_divs:")
            for div in self.orphaned_web_divs:
                lines.append(f"  - divId: {div['divId']}")
                lines.append(f"    reason: \"{div['reason']}\"")
                lines.append("")
        
        # New cruft sections
        if self.rsid_attributes:
            lines.append("rsid_attributes:")
            lines.append(f"  total_count: {self.total_rsid_attributes}")
            lines.append("  by_file:")
            for file, count in self.rsid_attributes.items():
                lines.append(f"    - file: {file}")
                lines.append(f"      count: {count}")
            lines.append("  reason: \"Revision tracking IDs - no value in final documents\"")
            lines.append("")
        
        if self.empty_elements:
            lines.append("empty_elements:")
            lines.append(f"  total_count: {self.total_empty_elements}")
            for elem in self.empty_elements[:10]:  # Limit to first 10 examples
                lines.append(f"  - file: {elem['file']}")
                lines.append(f"    element: {elem['element']}")
                lines.append(f"    count: {elem['count']}")
            if len(self.empty_elements) > 10:
                lines.append(f"  # ... and {len(self.empty_elements) - 10} more")
            lines.append("")
        
        if self.non_english_font_mappings:
            lines.append("non_english_font_mappings:")
            lines.append(f"  total_count: {len(self.non_english_font_mappings)}")
            lines.append("  scripts_to_remove:")
            for mapping in self.non_english_font_mappings[:20]:
                lines.append(f"    - {mapping['script']}")
            if len(self.non_english_font_mappings) > 20:
                lines.append(f"    # ... and {len(self.non_english_font_mappings) - 20} more")
            lines.append("  reason: \"Font mappings for non-English scripts not needed\"")
            lines.append("")
        
        if self.compatibility_settings:
            lines.append("compatibility_settings:")
            for compat in self.compatibility_settings:
                lines.append(f"  - setting: {compat['setting']}")
                lines.append(f"    reason: \"{compat['reason']}\"")
            lines.append("")
        
        if self.internal_bookmarks:
            lines.append("internal_bookmarks:")
            for bm in self.internal_bookmarks:
                lines.append(f"  - name: {bm['name']}")
                lines.append(f"    reason: \"{bm['reason']}\"")
            lines.append("")
        
        if self.proof_state_elements:
            lines.append("proof_state_elements:")
            lines.append(f"  count: {len(self.proof_state_elements)}")
            lines.append("  reason: \"Spell/grammar check state - no value in final documents\"")
            lines.append("")
        
        lines.append(f"estimated_savings_bytes: {self.estimated_savings_bytes}")
        
        return "\n".join(lines)


class DocxOrphanAnalyzer:
    """Analyzes DOCX files to find orphaned/unused resources."""
    
    # Common Word namespaces
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
        'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    }
    
    def __init__(self, extract_dir: Path):
        """
        Initialize analyzer with path to extracted DOCX directory.
        
        Args:
            extract_dir: Path to directory containing extracted DOCX contents
        """
        self.extract_dir = Path(extract_dir)
        self.report = OrphanReport()
        
        # Tracking sets
        self._defined_rids: Dict[str, Dict] = {}  # rId -> {target, type, source_file}
        self._used_rids: Set[str] = set()
        
        self._defined_styles: Dict[str, Dict] = {}  # styleId -> {name, type, basedOn}
        self._used_styles: Set[str] = set()
        self._style_dependencies: Dict[str, Set[str]] = {}  # styleId -> set of dependent styleIds
        
        self._defined_media: Dict[str, int] = {}  # relative path -> size in bytes
        self._referenced_media: Set[str] = set()
        
        self._defined_fonts: Set[str] = set()
        self._used_fonts: Set[str] = set()
        
        self._defined_numIds: Set[str] = set()
        self._used_numIds: Set[str] = set()
        
        self._web_div_ids: Set[str] = set()
        
        # RSID tracking
        self._rsid_counts: Dict[str, int] = {}  # file -> count of RSID attributes
        
        # Empty element tracking
        self._empty_elements: Dict[str, Dict[str, int]] = {}  # file -> {element_type: count}
        
        # Non-English scripts in theme
        self._non_english_scripts: List[Dict] = []
        
        # Compatibility settings
        self._compat_settings: List[Dict] = []
        
        # Internal bookmarks
        self._internal_bookmarks: List[Dict] = []
        
        # Proof state elements
        self._proof_elements: List[Dict] = []
    
    def analyze(self) -> OrphanReport:
        """
        Run complete orphan analysis.
        
        Returns:
            OrphanReport with all findings
        """
        print(f"Analyzing {self.extract_dir} for orphaned resources and cruft...")
        
        # Phase 1: Collect all DEFINED resources
        self._collect_defined_relationships()
        self._collect_defined_styles()
        self._collect_defined_media()
        self._collect_defined_fonts()
        self._collect_defined_numbering()
        self._collect_web_settings()
        
        # Phase 2: Scan content to find USED resources
        self._scan_document_for_usage()
        self._scan_headers_footers_for_usage()
        self._scan_footnotes_endnotes_for_usage()
        
        # Phase 3: Resolve style dependencies (basedOn, next, link chains)
        self._resolve_style_dependencies()
        
        # Phase 4: Compute orphans (DEFINED - USED)
        self._compute_orphaned_relationships()
        self._compute_orphaned_styles()
        self._compute_orphaned_media()
        self._compute_orphaned_fonts()
        self._compute_orphaned_numbering()
        
        # Phase 5: Scan for cruft (RSIDs, empty elements, etc.)
        self._scan_for_rsids()
        self._scan_for_empty_elements()
        self._scan_theme_for_non_english()
        self._scan_for_compatibility_settings()
        self._scan_for_internal_bookmarks()
        self._scan_for_proof_state()
        
        # Phase 6: Calculate estimated savings
        self._calculate_savings()
        
        return self.report
    
    def _parse_xml(self, file_path: Path) -> Optional[ET.Element]:
        """Safely parse an XML file and return root element."""
        if not file_path.exists():
            return None
        try:
            tree = ET.parse(file_path)
            return tree.getroot()
        except ET.ParseError as e:
            print(f"  Warning: Failed to parse {file_path}: {e}")
            return None
    
    # =========================================================================
    # Phase 1: Collect DEFINED resources
    # =========================================================================
    
    def _collect_defined_relationships(self):
        """Collect all relationships from .rels files."""
        print("  Collecting defined relationships...")
        
        for rels_file in self.extract_dir.rglob('*.rels'):
            root = self._parse_xml(rels_file)
            if root is None:
                continue
            
            # Determine the source file this rels file belongs to
            source_file = str(rels_file.relative_to(self.extract_dir))
            
            for rel in root.iter():
                if rel.tag.endswith('Relationship'):
                    rid = rel.get('Id', '')
                    target = rel.get('Target', '')
                    rel_type = rel.get('Type', '')
                    target_mode = rel.get('TargetMode', 'Internal')
                    
                    # Create composite key for rels from different sources
                    key = f"{source_file}:{rid}"
                    
                    self._defined_rids[key] = {
                        'rId': rid,
                        'target': target,
                        'type': rel_type,
                        'target_mode': target_mode,
                        'source_file': source_file
                    }
        
        self.report.total_relationships_defined = len(self._defined_rids)
        print(f"    Found {len(self._defined_rids)} defined relationships")
    
    def _collect_defined_styles(self):
        """Collect all styles from styles.xml."""
        print("  Collecting defined styles...")
        
        styles_path = self.extract_dir / "word" / "styles.xml"
        root = self._parse_xml(styles_path)
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        
        for style in root.iter(f'{{{w_ns}}}style'):
            style_id = style.get(f'{{{w_ns}}}styleId', '')
            style_type = style.get(f'{{{w_ns}}}type', '')
            
            name_elem = style.find(f'{{{w_ns}}}name')
            name = name_elem.get(f'{{{w_ns}}}val', '') if name_elem is not None else ''
            
            based_on_elem = style.find(f'{{{w_ns}}}basedOn')
            based_on = based_on_elem.get(f'{{{w_ns}}}val', '') if based_on_elem is not None else None
            
            next_elem = style.find(f'{{{w_ns}}}next')
            next_style = next_elem.get(f'{{{w_ns}}}val', '') if next_elem is not None else None
            
            link_elem = style.find(f'{{{w_ns}}}link')
            link_style = link_elem.get(f'{{{w_ns}}}val', '') if link_elem is not None else None
            
            self._defined_styles[style_id] = {
                'styleId': style_id,
                'name': name,
                'type': style_type,
                'basedOn': based_on,
                'next': next_style,
                'link': link_style
            }
            
            # Track dependencies
            deps = set()
            if based_on:
                deps.add(based_on)
            if next_style:
                deps.add(next_style)
            if link_style:
                deps.add(link_style)
            self._style_dependencies[style_id] = deps
        
        self.report.total_styles_defined = len(self._defined_styles)
        print(f"    Found {len(self._defined_styles)} defined styles")
    
    def _collect_defined_media(self):
        """Collect all media files in word/media/."""
        print("  Collecting defined media files...")
        
        media_dir = self.extract_dir / "word" / "media"
        if not media_dir.exists():
            print("    No media directory found")
            return
        
        for media_file in media_dir.iterdir():
            if media_file.is_file():
                rel_path = str(media_file.relative_to(self.extract_dir))
                self._defined_media[rel_path] = media_file.stat().st_size
        
        self.report.total_media_files = len(self._defined_media)
        print(f"    Found {len(self._defined_media)} media files")
    
    def _collect_defined_fonts(self):
        """Collect all fonts from fontTable.xml."""
        print("  Collecting defined fonts...")
        
        font_path = self.extract_dir / "word" / "fontTable.xml"
        root = self._parse_xml(font_path)
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        
        for font in root.iter(f'{{{w_ns}}}font'):
            font_name = font.get(f'{{{w_ns}}}name', '')
            if font_name:
                self._defined_fonts.add(font_name)
        
        print(f"    Found {len(self._defined_fonts)} defined fonts")
    
    def _collect_defined_numbering(self):
        """Collect all numbering definitions from numbering.xml."""
        print("  Collecting defined numbering...")
        
        numbering_path = self.extract_dir / "word" / "numbering.xml"
        root = self._parse_xml(numbering_path)
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        
        for num in root.iter(f'{{{w_ns}}}num'):
            num_id = num.get(f'{{{w_ns}}}numId', '')
            if num_id:
                self._defined_numIds.add(num_id)
        
        print(f"    Found {len(self._defined_numIds)} numbering definitions")
    
    def _collect_web_settings(self):
        """Collect div IDs from webSettings.xml (paste artifacts)."""
        print("  Collecting webSettings divs...")
        
        web_path = self.extract_dir / "word" / "webSettings.xml"
        root = self._parse_xml(web_path)
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        
        for div in root.iter(f'{{{w_ns}}}div'):
            div_id = div.get(f'{{{w_ns}}}id', '')
            if div_id:
                self._web_div_ids.add(div_id)
        
        print(f"    Found {len(self._web_div_ids)} webSettings divs")
    
    # =========================================================================
    # Phase 2: Scan content for USED resources
    # =========================================================================
    
    def _scan_xml_for_references(self, root: ET.Element, source_rels_file: str):
        """
        Scan an XML element tree for resource references.
        
        Looks for:
        - r:id attributes (relationship references)
        - w:pStyle, w:rStyle values (style references)
        - w:rFonts values (font references)
        - w:numId values (numbering references)
        """
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        r_ns = self.NAMESPACES['r']
        
        for elem in root.iter():
            # Check for relationship references (r:id, r:embed, etc.)
            for attr_name, attr_val in elem.attrib.items():
                if 'id' in attr_name.lower() or 'embed' in attr_name.lower():
                    if attr_val.startswith('rId'):
                        key = f"{source_rels_file}:{attr_val}"
                        self._used_rids.add(key)
                        # Also add without source for cross-file refs
                        self._used_rids.add(attr_val)
            
            # Style references
            tag_local = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            
            if tag_local in ('pStyle', 'rStyle', 'tblStyle'):
                style_val = elem.get(f'{{{w_ns}}}val', '')
                if style_val:
                    self._used_styles.add(style_val)
            
            # Font references
            if tag_local == 'rFonts':
                for font_attr in ['ascii', 'hAnsi', 'cs', 'eastAsia']:
                    font_val = elem.get(f'{{{w_ns}}}{font_attr}', '')
                    if font_val:
                        self._used_fonts.add(font_val)
            
            # Numbering references
            if tag_local == 'numId':
                num_val = elem.get(f'{{{w_ns}}}val', '')
                if num_val and num_val != '0':
                    self._used_numIds.add(num_val)
    
    def _scan_document_for_usage(self):
        """Scan document.xml for resource usage."""
        print("  Scanning document.xml for usage...")
        
        doc_path = self.extract_dir / "word" / "document.xml"
        root = self._parse_xml(doc_path)
        
        # The rels file for document.xml
        rels_file = "word/_rels/document.xml.rels"
        
        self._scan_xml_for_references(root, rels_file)
    
    def _scan_headers_footers_for_usage(self):
        """Scan header and footer XML files for resource usage."""
        print("  Scanning headers and footers...")
        
        word_dir = self.extract_dir / "word"
        
        for hf_file in word_dir.glob('header*.xml'):
            root = self._parse_xml(hf_file)
            rels_file = f"word/_rels/{hf_file.name}.rels"
            self._scan_xml_for_references(root, rels_file)
        
        for hf_file in word_dir.glob('footer*.xml'):
            root = self._parse_xml(hf_file)
            rels_file = f"word/_rels/{hf_file.name}.rels"
            self._scan_xml_for_references(root, rels_file)
    
    def _scan_footnotes_endnotes_for_usage(self):
        """Scan footnotes and endnotes for resource usage."""
        print("  Scanning footnotes and endnotes...")
        
        for fn_file in ['footnotes.xml', 'endnotes.xml']:
            fn_path = self.extract_dir / "word" / fn_file
            root = self._parse_xml(fn_path)
            rels_file = f"word/_rels/{fn_file}.rels"
            self._scan_xml_for_references(root, rels_file)
    
    # =========================================================================
    # Phase 3: Resolve dependencies
    # =========================================================================
    
    def _resolve_style_dependencies(self):
        """
        Expand used styles to include their dependencies.
        
        If style A is used and A is basedOn B, then B is also "used".
        """
        print("  Resolving style dependencies...")
        
        # Keep expanding until no new styles are added
        expanded = True
        iterations = 0
        max_iterations = 100  # Safety limit
        
        while expanded and iterations < max_iterations:
            expanded = False
            iterations += 1
            
            new_styles = set()
            for style_id in self._used_styles:
                if style_id in self._style_dependencies:
                    for dep in self._style_dependencies[style_id]:
                        if dep and dep not in self._used_styles:
                            new_styles.add(dep)
                            expanded = True
            
            self._used_styles.update(new_styles)
        
        # Also mark built-in styles that Word requires
        builtin_required = {'Normal', 'DefaultParagraphFont', 'TableNormal', 'NoList'}
        self._used_styles.update(builtin_required)
        
        self.report.total_styles_used = len(self._used_styles)
        print(f"    {len(self._used_styles)} styles are in use (after dependency resolution)")
    
    # =========================================================================
    # Phase 4: Compute orphans
    # =========================================================================
    
    def _compute_orphaned_relationships(self):
        """Find relationships that are never referenced."""
        print("  Computing orphaned relationships...")
        
        for key, rel_info in self._defined_rids.items():
            rid = rel_info['rId']
            
            # Check if this rId is used (either with full key or just rId)
            is_used = key in self._used_rids or rid in self._used_rids
            
            # Skip certain essential relationship types
            essential_types = [
                'styles', 'settings', 'fontTable', 'numbering', 'webSettings',
                'theme', 'footnotes', 'endnotes'
            ]
            is_essential = any(t in rel_info['type'].lower() for t in essential_types)
            
            if not is_used and not is_essential:
                self.report.orphaned_relationships.append({
                    'rId': rid,
                    'target': rel_info['target'],
                    'type': rel_info['type'].split('/')[-1],  # Just the type name
                    'target_mode': rel_info['target_mode'],
                    'source_file': rel_info['source_file'],
                    'reason': f"No r:id reference found in content files"
                })
        
        self.report.total_relationships_used = len(self._used_rids)
        print(f"    Found {len(self.report.orphaned_relationships)} orphaned relationships")
    
    def _compute_orphaned_styles(self):
        """Find styles that are never used."""
        print("  Computing orphaned styles...")
        
        for style_id, style_info in self._defined_styles.items():
            if style_id not in self._used_styles:
                self.report.orphaned_styles.append({
                    'styleId': style_id,
                    'name': style_info['name'],
                    'type': style_info['type'],
                    'reason': "Never referenced in document content or by other styles"
                })
        
        print(f"    Found {len(self.report.orphaned_styles)} orphaned styles")
    
    def _compute_orphaned_media(self):
        """Find media files not referenced by any relationship."""
        print("  Computing orphaned media...")
        
        # Build set of media files referenced by relationships
        referenced_media = set()
        for rel_info in self._defined_rids.values():
            target = rel_info['target']
            if target.startswith('media/'):
                referenced_media.add(f"word/{target}")
            elif 'media/' in target:
                referenced_media.add(target)
        
        for media_path, size in self._defined_media.items():
            # Normalize path for comparison
            normalized = media_path.replace('\\', '/')
            
            is_referenced = any(
                normalized.endswith(ref.split('/')[-1]) 
                for ref in referenced_media
            )
            
            if not is_referenced:
                self.report.orphaned_media.append({
                    'path': media_path,
                    'size_bytes': size,
                    'reason': "No relationship references this media file"
                })
        
        self.report.total_media_referenced = len(referenced_media)
        print(f"    Found {len(self.report.orphaned_media)} orphaned media files")
    
    def _compute_orphaned_fonts(self):
        """Find fonts declared but not used."""
        # Font table cleanup is generally safe but minimal impact
        # For now, we just report for awareness
        orphaned = self._defined_fonts - self._used_fonts
        
        for font in orphaned:
            self.report.orphaned_fonts.append({
                'font_name': font,
                'reason': "Declared in fontTable.xml but not referenced in content"
            })
        
        print(f"    Found {len(self.report.orphaned_fonts)} orphaned font declarations")
    
    def _compute_orphaned_numbering(self):
        """Find numbering definitions not used."""
        orphaned = self._defined_numIds - self._used_numIds - {'0'}  # 0 means "no numbering"
        
        for num_id in orphaned:
            self.report.orphaned_numbering.append({
                'numId': num_id,
                'reason': "Numbering definition not referenced by any paragraph"
            })
        
        print(f"    Found {len(self.report.orphaned_numbering)} orphaned numbering definitions")
    
    # =========================================================================
    # Phase 5: Scan for cruft (RSIDs, empty elements, etc.)
    # =========================================================================
    
    def _scan_for_rsids(self):
        """
        Scan all XML files for RSID attributes.
        
        RSIDs (Revision Save IDs) track editing sessions but serve no purpose
        in final documents. They appear as:
        - w:rsidR, w:rsidRPr, w:rsidRDefault, w:rsidP on paragraphs
        - w:rsidR, w:rsidRPr on runs
        - w:rsid on various elements
        """
        print("  Scanning for RSID attributes...")
        
        rsid_pattern = re.compile(r'\bw:rsid[A-Za-z]*\s*=\s*"[^"]*"', re.IGNORECASE)
        rsid_attr_pattern = re.compile(r'rsid', re.IGNORECASE)
        
        total_rsids = 0
        
        for xml_file in self.extract_dir.rglob('*.xml'):
            rel_path = str(xml_file.relative_to(self.extract_dir))
            
            try:
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Count RSID attributes in file
                matches = rsid_pattern.findall(content)
                if matches:
                    self._rsid_counts[rel_path] = len(matches)
                    total_rsids += len(matches)
            
            except Exception as e:
                pass  # Skip files that can't be read
        
        self.report.rsid_attributes = self._rsid_counts
        self.report.total_rsid_attributes = total_rsids
        print(f"    Found {total_rsids} RSID attributes across {len(self._rsid_counts)} files")
    
    def _scan_for_empty_elements(self):
        """
        Scan for empty runs, paragraphs, and other elements that serve no purpose.
        
        Empty elements include:
        - <w:r></w:r> or <w:r/> (empty runs)
        - <w:p><w:pPr/></w:p> (paragraphs with only empty properties)
        - <w:rPr></w:rPr> or <w:rPr/> (empty run properties)
        """
        print("  Scanning for empty elements...")
        
        xml_files_to_scan = [
            'word/document.xml',
            'word/styles.xml',
            'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
            'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml',
            'word/footnotes.xml', 'word/endnotes.xml',
        ]
        
        total_empty = 0
        
        for rel_path in xml_files_to_scan:
            xml_file = self.extract_dir / rel_path
            if not xml_file.exists():
                continue
            
            root = self._parse_xml(xml_file)
            if root is None:
                continue
            
            w_ns = self.NAMESPACES['w']
            file_empties = {}
            
            # Check for empty runs (no text children)
            for run in root.iter(f'{{{w_ns}}}r'):
                has_content = False
                for child in run:
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    # Content elements: t (text), drawing, pict, sym, tab, br, etc.
                    if tag in ('t', 'drawing', 'pict', 'sym', 'tab', 'br', 'cr', 
                              'fldChar', 'instrText', 'object', 'ruby'):
                        has_content = True
                        break
                    # Check if text element has actual content
                    if tag == 't' and child.text:
                        has_content = True
                        break
                
                if not has_content:
                    file_empties['empty_runs'] = file_empties.get('empty_runs', 0) + 1
            
            # Check for empty rPr (run properties with no children)
            for rpr in root.iter(f'{{{w_ns}}}rPr'):
                if len(rpr) == 0:
                    file_empties['empty_rPr'] = file_empties.get('empty_rPr', 0) + 1
            
            # Check for empty pPr (paragraph properties with no meaningful children)
            for ppr in root.iter(f'{{{w_ns}}}pPr'):
                has_meaningful = False
                for child in ppr:
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    # Meaningful paragraph properties
                    if tag in ('pStyle', 'numPr', 'spacing', 'ind', 'jc', 'tabs',
                              'keepNext', 'keepLines', 'pageBreakBefore', 'outlineLvl'):
                        has_meaningful = True
                        break
                    # rPr inside pPr is meaningful if it has children
                    if tag == 'rPr' and len(child) > 0:
                        has_meaningful = True
                        break
                
                if not has_meaningful and len(ppr) == 0:
                    file_empties['empty_pPr'] = file_empties.get('empty_pPr', 0) + 1
            
            if file_empties:
                self._empty_elements[rel_path] = file_empties
                for elem_type, count in file_empties.items():
                    total_empty += count
                    self.report.empty_elements.append({
                        'file': rel_path,
                        'element': elem_type,
                        'count': count
                    })
        
        self.report.total_empty_elements = total_empty
        print(f"    Found {total_empty} empty elements")
    
    def _scan_theme_for_non_english(self):
        """
        Scan theme files for non-English font mappings.
        
        Theme files contain font mappings for many scripts (Jpan, Hans, Arab, etc.)
        that are unnecessary for English-only documents.
        """
        print("  Scanning theme for non-English font mappings...")
        
        # Scripts to keep (Latin/Western)
        keep_scripts = {'latin', 'ea', 'cs'}  # ea=East Asian placeholder, cs=Complex Script placeholder
        
        theme_dir = self.extract_dir / "word" / "theme"
        if not theme_dir.exists():
            return
        
        for theme_file in theme_dir.glob('*.xml'):
            root = self._parse_xml(theme_file)
            if root is None:
                continue
            
            # Look for font elements with script attribute
            for elem in root.iter():
                tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                
                if tag == 'font':
                    script = elem.get('script', '')
                    typeface = elem.get('typeface', '')
                    
                    if script and script.lower() not in keep_scripts:
                        self._non_english_scripts.append({
                            'file': str(theme_file.relative_to(self.extract_dir)),
                            'script': script,
                            'typeface': typeface
                        })
        
        self.report.non_english_font_mappings = self._non_english_scripts
        print(f"    Found {len(self._non_english_scripts)} non-English font mappings")
    
    def _scan_for_compatibility_settings(self):
        """
        Scan settings.xml for backwards compatibility cruft.
        
        Word adds many <w:compat> child elements for compatibility with
        older Word versions. These are unnecessary for modern documents.
        """
        print("  Scanning for compatibility settings...")
        
        settings_path = self.extract_dir / "word" / "settings.xml"
        root = self._parse_xml(settings_path)
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        
        # Known compatibility settings that can be removed
        removable_compat = {
            'compatSetting': 'Compatibility mode settings',
            'useFELayout': 'Far East layout compatibility',
            'useWord2002TableStyleRules': 'Word 2002 table style compatibility',
            'growAutofit': 'Legacy autofit behavior',
            'useWord97LineBreakRules': 'Word 97 line break compatibility',
            'doNotUseIndentAsNumberingTabStop': 'Legacy numbering compatibility',
            'useAltKinsokuLineBreakRules': 'Japanese line break compatibility',
            'allowSpaceOfSameStyleInTable': 'Legacy table spacing',
            'doNotSuppressParagraphBorders': 'Legacy paragraph border behavior',
            'doNotAutofitConstrainedTables': 'Legacy table autofit',
            'autofitToFirstFixedWidthCell': 'Legacy autofit behavior',
            'displayHangulFixedWidth': 'Korean text compatibility',
            'splitPgBreakAndParaMark': 'Legacy page break behavior',
            'doNotVertAlignCellWithSp': 'Legacy cell alignment',
            'doNotBreakConstrainedForcedTable': 'Legacy table behavior',
            'doNotVertAlignInTxbx': 'Legacy textbox alignment',
            'useAnsiKerningPairs': 'Legacy kerning',
            'cachedColBalance': 'Legacy column balancing',
        }
        
        # Find compat element
        for compat in root.iter(f'{{{w_ns}}}compat'):
            for child in compat:
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                
                # Check for compatSetting elements
                if tag == 'compatSetting':
                    name = child.get(f'{{{w_ns}}}name', '')
                    uri = child.get(f'{{{w_ns}}}uri', '')
                    # Skip essential compatibility settings
                    if 'overrideTableStyleFontSizeAndJustification' not in name:
                        self._compat_settings.append({
                            'setting': f"compatSetting: {name}",
                            'reason': f"Compatibility setting for legacy behavior ({uri})"
                        })
                elif tag in removable_compat:
                    self._compat_settings.append({
                        'setting': tag,
                        'reason': removable_compat[tag]
                    })
        
        # Also check for documentCompatibility at root level
        for elem in root:
            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            if 'compat' in tag.lower() and tag != 'compat':
                self._compat_settings.append({
                    'setting': tag,
                    'reason': 'Legacy compatibility element'
                })
        
        self.report.compatibility_settings = self._compat_settings
        print(f"    Found {len(self._compat_settings)} compatibility settings")
    
    def _scan_for_internal_bookmarks(self):
        """
        Scan for Word's internal bookmarks that serve no user purpose.
        
        Internal bookmarks include:
        - _GoBack (last cursor position)
        - _Ref* (cross-reference anchors if references are removed)
        - _Toc* (TOC anchors if TOC is removed)
        - _Hlk* (hyperlink anchors)
        """
        print("  Scanning for internal bookmarks...")
        
        internal_prefixes = ('_GoBack', '_Ref', '_Toc', '_Hlk', '_PictureBullets')
        
        doc_path = self.extract_dir / "word" / "document.xml"
        root = self._parse_xml(doc_path)
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        
        for bookmark in root.iter(f'{{{w_ns}}}bookmarkStart'):
            name = bookmark.get(f'{{{w_ns}}}name', '')
            
            if any(name.startswith(prefix) for prefix in internal_prefixes):
                reason = "Internal Word bookmark"
                if name == '_GoBack':
                    reason = "Cursor position bookmark - no user value"
                elif name.startswith('_Hlk'):
                    reason = "Hyperlink anchor bookmark - can be removed if hyperlink is gone"
                elif name.startswith('_Toc'):
                    reason = "Table of Contents anchor"
                elif name.startswith('_Ref'):
                    reason = "Cross-reference anchor"
                
                self._internal_bookmarks.append({
                    'name': name,
                    'reason': reason
                })
        
        self.report.internal_bookmarks = self._internal_bookmarks
        print(f"    Found {len(self._internal_bookmarks)} internal bookmarks")
    
    def _scan_for_proof_state(self):
        """
        Scan for spell/grammar check state elements.
        
        These track whether the document has been proofed but serve
        no purpose in final documents.
        """
        print("  Scanning for proof state elements...")
        
        settings_path = self.extract_dir / "word" / "settings.xml"
        root = self._parse_xml(settings_path)
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        
        for elem in root.iter():
            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            
            if tag in ('proofState', 'proofErr', 'noProof'):
                self._proof_elements.append({
                    'element': tag,
                    'file': 'word/settings.xml'
                })
        
        # Also check document.xml for proofErr markers
        doc_path = self.extract_dir / "word" / "document.xml"
        root = self._parse_xml(doc_path)
        if root:
            for elem in root.iter(f'{{{w_ns}}}proofErr'):
                self._proof_elements.append({
                    'element': 'proofErr',
                    'file': 'word/document.xml'
                })
        
        self.report.proof_state_elements = self._proof_elements
        print(f"    Found {len(self._proof_elements)} proof state elements")
    
    # =========================================================================
    # Phase 5: Calculate savings
    # =========================================================================
    
    def _calculate_savings(self):
        """Estimate bytes that could be saved by removing orphans and cruft."""
        print("  Calculating estimated savings...")
        
        savings = 0
        
        # Media files are the big savings
        for media in self.report.orphaned_media:
            savings += media['size_bytes']
        
        # Estimate ~200 bytes per orphaned relationship (XML text)
        savings += len(self.report.orphaned_relationships) * 200
        
        # Estimate ~500 bytes per orphaned style definition
        savings += len(self.report.orphaned_styles) * 500
        
        # Web divs are usually ~200 bytes each
        savings += len(self.report.orphaned_web_divs) * 200
        
        # RSID attributes: ~25 bytes each (w:rsidR="00A77B3E")
        savings += self.report.total_rsid_attributes * 25
        
        # Empty elements: ~20 bytes each on average
        savings += self.report.total_empty_elements * 20
        
        # Non-English font mappings: ~60 bytes each
        savings += len(self.report.non_english_font_mappings) * 60
        
        # Compatibility settings: ~50 bytes each
        savings += len(self.report.compatibility_settings) * 50
        
        # Internal bookmarks: ~80 bytes each (start + end tags)
        savings += len(self.report.internal_bookmarks) * 80
        
        # Proof state elements: ~40 bytes each
        savings += len(self.report.proof_state_elements) * 40
        
        self.report.estimated_savings_bytes = savings
        print(f"    Estimated savings: {savings:,} bytes ({savings/1024:.1f} KB)")


def main():
    """CLI entry point."""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python docx_orphan_analyzer.py <extracted_docx_directory>")
        print("\nThis tool analyzes an extracted DOCX directory for orphaned resources.")
        print("Use docx_obliterate.py to extract a DOCX file first.")
        sys.exit(1)
    
    extract_dir = Path(sys.argv[1])
    
    if not extract_dir.exists():
        print(f"Error: Directory not found: {extract_dir}")
        sys.exit(1)
    
    # Check for expected DOCX structure
    if not (extract_dir / "[Content_Types].xml").exists():
        print(f"Error: {extract_dir} does not appear to be an extracted DOCX directory")
        print("Expected to find [Content_Types].xml")
        sys.exit(1)
    
    analyzer = DocxOrphanAnalyzer(extract_dir)
    report = analyzer.analyze()
    
    # Output results
    print("\n" + "=" * 60)
    print("ORPHAN & CRUFT ANALYSIS COMPLETE")
    print("=" * 60)
    
    print("\n--- Orphaned Resources ---")
    print(f"Orphaned Relationships: {len(report.orphaned_relationships)}")
    print(f"Orphaned Media Files:   {len(report.orphaned_media)}")
    print(f"Orphaned Styles:        {len(report.orphaned_styles)}")
    print(f"Orphaned Fonts:         {len(report.orphaned_fonts)}")
    print(f"Orphaned Numbering:     {len(report.orphaned_numbering)}")
    
    print("\n--- Cleanable Cruft ---")
    print(f"RSID Attributes:        {report.total_rsid_attributes}")
    print(f"Empty Elements:         {report.total_empty_elements}")
    print(f"Non-English Fonts:      {len(report.non_english_font_mappings)}")
    print(f"Compat Settings:        {len(report.compatibility_settings)}")
    print(f"Internal Bookmarks:     {len(report.internal_bookmarks)}")
    print(f"Proof State Elements:   {len(report.proof_state_elements)}")
    
    print(f"\n--- Summary ---")
    print(f"Estimated Savings:      {report.estimated_savings_bytes:,} bytes ({report.estimated_savings_bytes/1024:.1f} KB)")
    
    # Save manifest
    manifest_path = extract_dir.parent / f"{extract_dir.name}_orphans.yaml"
    with open(manifest_path, 'w') as f:
        f.write(report.to_yaml_manifest())
    print(f"\nManifest saved to: {manifest_path}")
    
    # Save JSON report
    json_path = extract_dir.parent / f"{extract_dir.name}_orphans.json"
    with open(json_path, 'w') as f:
        f.write(report.to_json())
    print(f"JSON report saved to: {json_path}")


if __name__ == "__main__":
    main()
