"""
SpecCleanse Deep Cleaner Module

Performs deep cleaning of DOCX files at the ZIP/XML level:
- Removes orphaned resources (relationships, media, styles, fonts, numbering)
- Strips cruft (RSIDs, empty elements, non-English fonts, compat settings, bookmarks, proof state)

This module consolidates the functionality from:
- docx_orphan_analyzer.py (detection)
- docx_deep_cleaner.py (removal)
"""

import re
from lxml import etree
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Set, Optional


# =============================================================================
# Data Classes
# =============================================================================

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
        """Convert to dictionary for internal use."""
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


@dataclass
class DeepCleanResult:
    """Results from deep cleaning operation."""
    success: bool = False
    
    # Orphans removed
    relationships_removed: int = 0
    media_removed: int = 0
    styles_removed: int = 0
    
    # Cruft removed
    rsids_removed: int = 0
    empty_elements_removed: int = 0
    font_mappings_removed: int = 0
    compat_settings_removed: int = 0
    bookmarks_removed: int = 0
    proof_elements_removed: int = 0
    
    bytes_saved: int = 0
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# =============================================================================
# Orphan Analyzer
# =============================================================================

class OrphanAnalyzer:
    """Analyzes DOCX files to find orphaned/unused resources."""
    
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
        'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    }
    
    def __init__(self, extract_dir: Path, verbose: bool = False):
        self.extract_dir = Path(extract_dir)
        self.verbose = verbose
        self.report = OrphanReport()
        
        # Tracking sets
        self._defined_rids: Dict[str, Dict] = {}
        self._used_rids: Set[str] = set()
        
        self._defined_styles: Dict[str, Dict] = {}
        self._used_styles: Set[str] = set()
        self._style_dependencies: Dict[str, Set[str]] = {}
        
        self._defined_media: Dict[str, int] = {}
        self._referenced_media: Set[str] = set()
        
        self._defined_fonts: Set[str] = set()
        self._used_fonts: Set[str] = set()
        
        self._defined_numIds: Set[str] = set()
        self._used_numIds: Set[str] = set()
        
        self._rsid_counts: Dict[str, int] = {}
        self._empty_elements: Dict[str, Dict[str, int]] = {}
        self._non_english_scripts: List[Dict] = []
        self._compat_settings: List[Dict] = []
        self._internal_bookmarks: List[Dict] = []
        self._proof_elements: List[Dict] = []
    
    def analyze(self) -> OrphanReport:
        """Run complete orphan analysis."""
        if self.verbose:
            print("  Analyzing for orphaned resources...")
        
        # Phase 1: Collect DEFINED resources
        self._collect_defined_relationships()
        self._collect_defined_styles()
        self._collect_defined_media()
        self._collect_defined_fonts()
        self._collect_defined_numbering()
        
        # Phase 2: Scan content for USED resources
        self._scan_document_for_usage()
        self._scan_headers_footers_for_usage()
        self._scan_footnotes_endnotes_for_usage()
        
        # Phase 3: Resolve style dependencies
        self._resolve_style_dependencies()
        
        # Phase 4: Compute orphans
        self._compute_orphaned_relationships()
        self._compute_orphaned_styles()
        self._compute_orphaned_media()
        
        # Phase 5: Scan for cruft
        self._scan_for_rsids()
        self._scan_for_empty_elements()
        self._scan_theme_for_non_english()
        self._scan_for_compatibility_settings()
        self._scan_for_internal_bookmarks()
        self._scan_for_proof_state()
        
        # Phase 6: Calculate savings
        self._calculate_savings()
        
        return self.report
    
    def _parse_xml(self, file_path: Path) -> Optional[etree._Element]:
        """Safely parse an XML file."""
        if not file_path.exists():
            return None
        try:
            parser = etree.XMLParser(remove_blank_text=False)
            tree = etree.parse(str(file_path), parser)
            return tree.getroot()
        except etree.XMLSyntaxError:
            return None
    
    def _collect_defined_relationships(self):
        """Collect all relationships from .rels files."""
        for rels_file in self.extract_dir.rglob('*.rels'):
            root = self._parse_xml(rels_file)
            if root is None:
                continue
            
            source_file = str(rels_file.relative_to(self.extract_dir))
            
            for rel in root.iter():
                if rel.tag.endswith('Relationship'):
                    rid = rel.get('Id', '')
                    target = rel.get('Target', '')
                    rel_type = rel.get('Type', '')
                    target_mode = rel.get('TargetMode', 'Internal')
                    
                    key = f"{source_file}:{rid}"
                    self._defined_rids[key] = {
                        'rId': rid,
                        'target': target,
                        'type': rel_type,
                        'target_mode': target_mode,
                        'source_file': source_file
                    }
        
        self.report.total_relationships_defined = len(self._defined_rids)
    
    def _collect_defined_styles(self):
        """Collect all styles from styles.xml."""
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
            
            deps = set()
            if based_on:
                deps.add(based_on)
            if next_style:
                deps.add(next_style)
            if link_style:
                deps.add(link_style)
            self._style_dependencies[style_id] = deps
        
        self.report.total_styles_defined = len(self._defined_styles)
    
    def _collect_defined_media(self):
        """Collect all media files in word/media/."""
        media_dir = self.extract_dir / "word" / "media"
        if not media_dir.exists():
            return
        
        for media_file in media_dir.iterdir():
            if media_file.is_file():
                rel_path = str(media_file.relative_to(self.extract_dir))
                self._defined_media[rel_path] = media_file.stat().st_size
        
        self.report.total_media_files = len(self._defined_media)
    
    def _collect_defined_fonts(self):
        """Collect all fonts from fontTable.xml."""
        font_path = self.extract_dir / "word" / "fontTable.xml"
        root = self._parse_xml(font_path)
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        for font in root.iter(f'{{{w_ns}}}font'):
            font_name = font.get(f'{{{w_ns}}}name', '')
            if font_name:
                self._defined_fonts.add(font_name)
    
    def _collect_defined_numbering(self):
        """Collect all numbering definitions from numbering.xml."""
        numbering_path = self.extract_dir / "word" / "numbering.xml"
        root = self._parse_xml(numbering_path)
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        for num in root.iter(f'{{{w_ns}}}num'):
            num_id = num.get(f'{{{w_ns}}}numId', '')
            if num_id:
                self._defined_numIds.add(num_id)
    
    def _scan_xml_for_references(self, root: etree._Element, source_rels_file: str):
        """Scan an XML element tree for resource references."""
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        
        for elem in root.iter():
            # Check for relationship references
            for attr_name, attr_val in elem.attrib.items():
                if 'id' in attr_name.lower() or 'embed' in attr_name.lower():
                    if attr_val.startswith('rId'):
                        key = f"{source_rels_file}:{attr_val}"
                        self._used_rids.add(key)
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
        doc_path = self.extract_dir / "word" / "document.xml"
        root = self._parse_xml(doc_path)
        rels_file = "word/_rels/document.xml.rels"
        self._scan_xml_for_references(root, rels_file)
    
    def _scan_headers_footers_for_usage(self):
        """Scan header and footer XML files for resource usage."""
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
        for fn_file in ['footnotes.xml', 'endnotes.xml']:
            fn_path = self.extract_dir / "word" / fn_file
            root = self._parse_xml(fn_path)
            rels_file = f"word/_rels/{fn_file}.rels"
            self._scan_xml_for_references(root, rels_file)
    
    def _resolve_style_dependencies(self):
        """Expand used styles to include their dependencies."""
        expanded = True
        iterations = 0
        max_iterations = 100
        
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
        
        # Mark built-in required styles
        builtin_required = {'Normal', 'DefaultParagraphFont', 'TableNormal', 'NoList'}
        self._used_styles.update(builtin_required)
        
        self.report.total_styles_used = len(self._used_styles)
    
    def _compute_orphaned_relationships(self):
        """Find relationships that are never referenced."""
        for key, rel_info in self._defined_rids.items():
            rid = rel_info['rId']
            is_used = key in self._used_rids or rid in self._used_rids
            
            essential_types = [
                'styles', 'settings', 'fontTable', 'numbering', 'webSettings',
                'theme', 'footnotes', 'endnotes'
            ]
            is_essential = any(t in rel_info['type'].lower() for t in essential_types)
            
            if not is_used and not is_essential:
                self.report.orphaned_relationships.append({
                    'rId': rid,
                    'target': rel_info['target'],
                    'type': rel_info['type'].split('/')[-1],
                    'target_mode': rel_info['target_mode'],
                    'source_file': rel_info['source_file'],
                    'reason': "No r:id reference found in content files"
                })
        
        self.report.total_relationships_used = len(self._used_rids)
    
    def _compute_orphaned_styles(self):
        """Find styles that are never used."""
        for style_id, style_info in self._defined_styles.items():
            if style_id not in self._used_styles:
                self.report.orphaned_styles.append({
                    'styleId': style_id,
                    'name': style_info['name'],
                    'type': style_info['type'],
                    'reason': "Never referenced in document content or by other styles"
                })
    
    def _compute_orphaned_media(self):
        """Find media files not referenced by any relationship."""
        referenced_media = set()
        for rel_info in self._defined_rids.values():
            target = rel_info['target']
            if target.startswith('media/'):
                referenced_media.add(f"word/{target}")
            elif 'media/' in target:
                referenced_media.add(target)
        
        for media_path, size in self._defined_media.items():
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
    
    def _scan_for_rsids(self):
        """Scan all XML files for RSID attributes."""
        rsid_pattern = re.compile(r'\bw:rsid[A-Za-z]*\s*=\s*"[^"]*"', re.IGNORECASE)
        total_rsids = 0
        
        for xml_file in self.extract_dir.rglob('*.xml'):
            rel_path = str(xml_file.relative_to(self.extract_dir))
            
            try:
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                matches = rsid_pattern.findall(content)
                if matches:
                    self._rsid_counts[rel_path] = len(matches)
                    total_rsids += len(matches)
            except Exception:
                pass
        
        self.report.rsid_attributes = self._rsid_counts
        self.report.total_rsid_attributes = total_rsids
    
    def _scan_for_empty_elements(self):
        """Scan for empty runs and other useless elements."""
        xml_files_to_scan = [
            'word/document.xml',
            'word/styles.xml',
            'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
            'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml',
            'word/footnotes.xml', 'word/endnotes.xml',
        ]
        
        total_empty = 0
        w_ns = self.NAMESPACES['w']
        
        for rel_path in xml_files_to_scan:
            xml_file = self.extract_dir / rel_path
            if not xml_file.exists():
                continue
            
            root = self._parse_xml(xml_file)
            if root is None:
                continue
            
            file_empties = {}
            
            # Check for empty runs
            for run in root.iter(f'{{{w_ns}}}r'):
                has_content = False
                for child in run:
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if tag in ('t', 'drawing', 'pict', 'sym', 'tab', 'br', 'cr', 
                              'fldChar', 'instrText', 'object', 'ruby'):
                        has_content = True
                        break
                
                if not has_content:
                    file_empties['empty_runs'] = file_empties.get('empty_runs', 0) + 1
            
            # Check for empty rPr
            for rpr in root.iter(f'{{{w_ns}}}rPr'):
                if len(rpr) == 0:
                    file_empties['empty_rPr'] = file_empties.get('empty_rPr', 0) + 1
            
            if file_empties:
                for elem_type, count in file_empties.items():
                    total_empty += count
                    self.report.empty_elements.append({
                        'file': rel_path,
                        'element': elem_type,
                        'count': count
                    })
        
        self.report.total_empty_elements = total_empty
    
    def _scan_theme_for_non_english(self):
        """Scan theme files for non-English font mappings."""
        keep_scripts = {'latin', 'ea', 'cs'}
        
        theme_dir = self.extract_dir / "word" / "theme"
        if not theme_dir.exists():
            return
        
        for theme_file in theme_dir.glob('*.xml'):
            root = self._parse_xml(theme_file)
            if root is None:
                continue
            
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
    
    def _scan_for_compatibility_settings(self):
        """Scan settings.xml for backwards compatibility cruft."""
        settings_path = self.extract_dir / "word" / "settings.xml"
        root = self._parse_xml(settings_path)
        if root is None:
            return
        
        w_ns = self.NAMESPACES['w']
        
        removable_compat = {
            'compatSetting': 'Compatibility mode settings',
            'useFELayout': 'Far East layout compatibility',
            'useWord2002TableStyleRules': 'Word 2002 table style compatibility',
            'growAutofit': 'Legacy autofit behavior',
            'useWord97LineBreakRules': 'Word 97 line break compatibility',
        }
        
        for compat in root.iter(f'{{{w_ns}}}compat'):
            for child in compat:
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                
                if tag == 'compatSetting':
                    name = child.get(f'{{{w_ns}}}name', '')
                    if 'overrideTableStyleFontSizeAndJustification' not in name:
                        self._compat_settings.append({
                            'setting': f"compatSetting: {name}",
                            'reason': "Compatibility setting for legacy behavior"
                        })
                elif tag in removable_compat:
                    self._compat_settings.append({
                        'setting': tag,
                        'reason': removable_compat[tag]
                    })
        
        self.report.compatibility_settings = self._compat_settings
    
    def _scan_for_internal_bookmarks(self):
        """Scan for Word's internal bookmarks."""
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
                    reason = "Cursor position bookmark"
                elif name.startswith('_Hlk'):
                    reason = "Hyperlink anchor bookmark"
                
                self._internal_bookmarks.append({
                    'name': name,
                    'reason': reason
                })
        
        self.report.internal_bookmarks = self._internal_bookmarks
    
    def _scan_for_proof_state(self):
        """Scan for spell/grammar check state elements."""
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
        
        # Also check document.xml
        doc_path = self.extract_dir / "word" / "document.xml"
        root = self._parse_xml(doc_path)
        if root:
            for elem in root.iter(f'{{{w_ns}}}proofErr'):
                self._proof_elements.append({
                    'element': 'proofErr',
                    'file': 'word/document.xml'
                })
        
        self.report.proof_state_elements = self._proof_elements
    
    def _calculate_savings(self):
        """Estimate bytes that could be saved."""
        savings = 0
        
        for media in self.report.orphaned_media:
            savings += media['size_bytes']
        
        savings += len(self.report.orphaned_relationships) * 200
        savings += len(self.report.orphaned_styles) * 500
        savings += self.report.total_rsid_attributes * 25
        savings += self.report.total_empty_elements * 20
        savings += len(self.report.non_english_font_mappings) * 60
        savings += len(self.report.compatibility_settings) * 50
        savings += len(self.report.internal_bookmarks) * 80
        savings += len(self.report.proof_state_elements) * 40
        
        self.report.estimated_savings_bytes = savings


# =============================================================================
# Deep Cleaner
# =============================================================================

class DeepCleaner:
    """Performs safe deep cleaning of DOCX files based on orphan analysis."""
    
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
        'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    }
    
    def __init__(self, extract_dir: Path, orphan_report: OrphanReport, verbose: bool = False):
        self.extract_dir = Path(extract_dir)
        self.orphan_report = orphan_report
        self.verbose = verbose
        self.result = DeepCleanResult()
        
    def clean(self,
              remove_relationships: bool = True,
              remove_media: bool = True,
              remove_styles: bool = True,
              strip_rsids: bool = True,
              remove_empty_elements: bool = True,
              remove_non_english_fonts: bool = True,
              remove_compat_settings: bool = True,
              remove_internal_bookmarks: bool = True,
              remove_proof_state: bool = True) -> DeepCleanResult:
        """Perform deep cleaning based on orphan report."""
        
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
        
        except Exception as e:
            self.result.errors.append(f"Cleaning failed: {str(e)}")
        
        return self.result
    
    def _remove_orphaned_media(self):
        """Remove orphaned media files."""
        for media in self.orphan_report.orphaned_media:
            media_path = self.extract_dir / media['path']
            if media_path.exists():
                size = media_path.stat().st_size
                media_path.unlink()
                self.result.media_removed += 1
                self.result.bytes_saved += size
    
    def _remove_orphaned_relationships(self):
        """Remove orphaned relationship entries from .rels files."""
        by_source: Dict[str, List[str]] = {}
        for orphan in self.orphan_report.orphaned_relationships:
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
                parser = etree.XMLParser(remove_blank_text = False)
                tree = etree.parse(str(rels_path), parser)
                root = tree.getroot()

                removed = 0
                for rel in list(root):
                    if rel.get('Id') in rids:
                        root.remove(rel)
                        removed += 1
                
                if removed > 0:
                    tree.write(str(rels_path), xml_declaration=True, encoding='UTF-8', standalone=True)
                    self.result.relationships_removed += removed
            
            except etree.XMLSyntaxError as e:
                self.result.warnings.append(f"Failed to parse {source_file}: {e}")
    
    def _remove_orphaned_styles(self):
        """Remove orphaned style definitions from styles.xml."""
        if not self.orphan_report.orphaned_styles:
            return
        
        orphan_ids = {o['styleId'] for o in self.orphan_report.orphaned_styles}
        
        styles_path = self.extract_dir / "word" / "styles.xml"
        if not styles_path.exists():
            return
        
        try:
            parser = etree.XMLParser(remove_blank_text = False)
            tree = etree.parse(str(styles_path), parser)
            root = tree.getroot()
            
            w_ns = self.NAMESPACES['w']
            
            removed = 0
            for style in list(root.findall(f'.//{{{w_ns}}}style')):
                style_id = style.get(f'{{{w_ns}}}styleId', '')
                if style_id in orphan_ids:
                    root.remove(style)
                    removed += 1
            
            if removed > 0:
                tree.write(str(styles_path), xml_declaration=True, encoding='UTF-8', standalone=True)

                self.result.styles_removed = removed
        
        except etree.XMLSyntaxError as e:
            self.result.warnings.append(f"Failed to parse styles.xml: {e}")
    
    def _strip_rsids(self):
        """Remove all RSID tracking attributes from XML files."""
        rsid_pattern = re.compile(r'\s+w:rsid[A-Za-z]*="[^"]*"', re.IGNORECASE)
        
        total_removed = 0
        
        for xml_file in self.extract_dir.rglob('*.xml'):
            try:
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
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
                parser = etree.XMLParser(remove_blank_text = False)
                tree = etree.parse(str(xml_file), parser)
                root = tree.getroot()
                modified = False
                
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
                    tree.write(str(xml_file), xml_declaration=True, encoding='UTF-8', standalone=True)

            
            except Exception as e:
                self.result.warnings.append(f"Failed to clean empty elements in {rel_path}: {e}")
        
        self.result.empty_elements_removed = total_removed
        self.result.bytes_saved += total_removed * 20
    
    def _remove_non_english_font_mappings(self):
        """Remove non-English font mappings from theme files."""
        keep_scripts = {'latin', 'ea', 'cs', ''}
        
        theme_dir = self.extract_dir / "word" / "theme"
        if not theme_dir.exists():
            return
        
        total_removed = 0
        
        for theme_file in theme_dir.glob('*.xml'):
            try:
                parser = etree.XMLParser(remove_blank_text=False)
                tree = etree.parse(str(theme_file), parser)
                root = tree.getroot()
                modified = False
                
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
                    tree.write(str(theme_file), xml_declaration=True, encoding='UTF-8', standalone=True)

            
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
            parser = etree.XMLParser(remove_blank_text=False)
            tree = etree.parse(str(settings_path), parser)
            root = tree.getroot()
            total_removed = 0
            
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
                tree.write(str(settings_path), xml_declaration=True, encoding='UTF-8', standalone=True)
            
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
            parser = etree.XMLParser(remove_blank_text=False)
            tree = etree.parse(str(doc_path), parser)
            root = tree.getroot()
            
            bookmark_ids_to_remove = set()
            bookmarks_removed = 0
            
            for bookmark in root.iter(f'{{{w_ns}}}bookmarkStart'):
                name = bookmark.get(f'{{{w_ns}}}name', '')
                if any(name.startswith(prefix) for prefix in internal_prefixes):
                    bm_id = bookmark.get(f'{{{w_ns}}}id', '')
                    bookmark_ids_to_remove.add(bm_id)
            
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
                    bookmarks_removed += 1
            
            if bookmarks_removed:
                tree.write(str(doc_path), xml_declaration=True, encoding='UTF-8', standalone=True)

            self.result.bookmarks_removed = bookmarks_removed // 2
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
                parser = etree.XMLParser(remove_blank_text=False)
                tree = etree.parse(str(settings_path), parser)
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
                    tree.write(str(settings_path), xml_declaration=True, encoding='UTF-8', standalone=True)

            
            except Exception as e:
                self.result.warnings.append(f"Failed to clean proof state from settings: {e}")
        
        # Clean document.xml
        doc_path = self.extract_dir / "word" / "document.xml"
        if doc_path.exists():
            try:
                parser = etree.XMLParser(remove_blank_text=False)
                tree = etree.parse(str(doc_path), parser)
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
                    tree.write(str(doc_path), xml_declaration=True, encoding='UTF-8', standalone=True)
                    total_removed += doc_removed
            
            except Exception as e:
                self.result.warnings.append(f"Failed to clean proof errors from document: {e}")
        
        self.result.proof_elements_removed = total_removed
        self.result.bytes_saved += total_removed * 40
    
    def _validate_structure(self) -> bool:
        """Validate the document structure is still intact."""
        # Check Content_Types
        ct_path = self.extract_dir / "[Content_Types].xml"
        if not ct_path.exists():
            self.result.errors.append("Missing [Content_Types].xml")
            return False
        
        try:
            etree.parse(str(ct_path))
        except etree.XMLSyntaxError as e:
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
            etree.parse(str(doc_path))
        except etree.XMLSyntaxError as e:
            self.result.errors.append(f"Invalid document.xml: {e}")
            return False
        
        return True


# =============================================================================
# Public API
# =============================================================================

def analyze_and_clean(
    unpacked_dir: Path,
    remove_relationships: bool = True,
    remove_media: bool = True,
    remove_styles: bool = True,
    strip_rsids: bool = True,
    remove_empty_elements: bool = True,
    remove_non_english_fonts: bool = True,
    remove_compat_settings: bool = True,
    remove_internal_bookmarks: bool = True,
    remove_proof_state: bool = True,
    verbose: bool = False
) -> DeepCleanResult:
    """
    Analyze and deep clean an unpacked DOCX directory.
    
    This is the main entry point for deep cleaning.
    
    Args:
        unpacked_dir: Path to unpacked DOCX directory
        remove_relationships: Remove orphaned hyperlinks/relationships
        remove_media: Remove orphaned media files
        remove_styles: Remove orphaned style definitions
        strip_rsids: Remove RSID tracking attributes
        remove_empty_elements: Remove empty runs/elements
        remove_non_english_fonts: Remove non-English font mappings from theme
        remove_compat_settings: Remove backwards compatibility settings
        remove_internal_bookmarks: Remove Word's internal bookmarks
        remove_proof_state: Remove spell/grammar check state
        verbose: Print progress information
    
    Returns:
        DeepCleanResult with details of operations performed
    """
    # Phase 1: Analyze for orphans
    analyzer = OrphanAnalyzer(unpacked_dir, verbose=verbose)
    orphan_report = analyzer.analyze()
    
    if verbose:
        print(f"    Found {len(orphan_report.orphaned_relationships)} orphaned relationships")
        print(f"    Found {len(orphan_report.orphaned_styles)} orphaned styles")
        print(f"    Found {len(orphan_report.orphaned_media)} orphaned media files")
        print(f"    Found {orphan_report.total_rsid_attributes} RSID attributes")
    
    # Phase 2: Deep clean
    cleaner = DeepCleaner(unpacked_dir, orphan_report, verbose=verbose)
    result = cleaner.clean(
        remove_relationships=remove_relationships,
        remove_media=remove_media,
        remove_styles=remove_styles,
        strip_rsids=strip_rsids,
        remove_empty_elements=remove_empty_elements,
        remove_non_english_fonts=remove_non_english_fonts,
        remove_compat_settings=remove_compat_settings,
        remove_internal_bookmarks=remove_internal_bookmarks,
        remove_proof_state=remove_proof_state,
    )
    
    return result


def get_analysis_only(unpacked_dir: Path, verbose: bool = False) -> OrphanReport:
    """
    Analyze an unpacked DOCX directory without making changes.
    
    Useful for dry-run reporting.
    
    Args:
        unpacked_dir: Path to unpacked DOCX directory
        verbose: Print progress information
    
    Returns:
        OrphanReport with analysis results
    """
    analyzer = OrphanAnalyzer(unpacked_dir, verbose=verbose)
    return analyzer.analyze()