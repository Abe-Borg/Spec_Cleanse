#!/usr/bin/env python3
"""
Word Document Decomposer and Reconstructor

This tool extracts the internal components of a .docx file (which is a ZIP archive
containing XML and other files), documents the structure in markdown, and can
reconstruct the original document from the extracted components.
"""

import zipfile
import os
import shutil
from pathlib import Path
from datetime import datetime
import xml.etree.ElementTree as ET


class DocxDecomposer:
    def __init__(self, docx_path):
        """
        Initialize the decomposer with a path to a .docx file.
        
        Args:
            docx_path: Path to the input .docx file
        """
        self.docx_path = Path(docx_path)
        self.extract_dir = None
        self.markdown_report = []
        
    def extract(self, output_dir=None):
        """
        Extract the .docx file to a directory.
        
        Args:
            output_dir: Directory to extract to. If None, creates a directory
                       based on the docx filename.
        
        Returns:
            Path to the extraction directory
        """
        if output_dir is None:
            base_name = self.docx_path.stem
            output_dir = Path(f"{base_name}_extracted")
        else:
            output_dir = Path(output_dir)
        
        # Remove existing directory if it exists
        if output_dir.exists():
            shutil.rmtree(output_dir)
        
        # Extract the ZIP archive
        print(f"Extracting {self.docx_path} to {output_dir}...")
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)
        
        self.extract_dir = output_dir
        print(f"Extraction complete: {len(list(output_dir.rglob('*')))} items extracted")
        return output_dir
    
    def analyze_structure(self):
        """
        Analyze the extracted directory structure and generate a COMPLETE markdown report.
        This goes to the atomic level - every file, every XML element, every attribute.
        
        Returns:
            String containing the markdown report
        """
        if self.extract_dir is None:
            raise ValueError("Must call extract() before analyze_structure()")
        
        self.markdown_report = []
        
        # Header
        self._add_header()
        
        # Directory structure
        self._add_directory_tree()
        
        # Complete file inventory
        self._add_complete_file_inventory()
        
        # Content types - COMPLETE
        self._add_content_types_complete()
        
        # All relationships - COMPLETE
        self._add_all_relationships()
        
        # Document XML - COMPLETE breakdown
        self._add_document_xml_complete()
        
        # Styles XML - COMPLETE
        self._add_styles_xml_complete()
        
        # Settings XML - COMPLETE
        self._add_settings_xml_complete()
        
        # Font table - COMPLETE
        self._add_font_table_complete()
        
        # Numbering - COMPLETE
        self._add_numbering_complete()
        
        # Theme - COMPLETE
        self._add_theme_complete()
        
        # Document properties - COMPLETE
        self._add_doc_properties_complete()
        
        # Custom XML - COMPLETE
        self._add_custom_xml_complete()
        
        # Web settings - COMPLETE
        self._add_web_settings_complete()
        
        # Any other XML files - COMPLETE
        self._add_other_xml_files()
        
        # Binary files analysis
        self._add_binary_files()
        
        # Raw XML dumps for all files
        self._add_raw_xml_dumps()
        
        return "\n".join(self.markdown_report)
    
    def _add_header(self):
        """Add markdown header."""
        self.markdown_report.append(f"# Word Document Structure Analysis")
        self.markdown_report.append(f"\n**Source Document:** `{self.docx_path.name}`")
        self.markdown_report.append(f"**Analysis Date:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.markdown_report.append(f"**Extraction Directory:** `{self.extract_dir}`")
        self.markdown_report.append("\n---\n")
    
    def _add_directory_tree(self):
        """Add directory tree structure."""
        self.markdown_report.append("## Directory Structure\n")
        self.markdown_report.append("```")
        self._print_tree(self.extract_dir, prefix="")
        self.markdown_report.append("```\n")
    
    def _print_tree(self, directory, prefix="", is_last=True):
        """Recursively print directory tree."""
        items = sorted(directory.iterdir(), key=lambda x: (not x.is_dir(), x.name))
        
        for i, item in enumerate(items):
            is_last_item = (i == len(items) - 1)
            current_prefix = "└── " if is_last_item else "├── "
            self.markdown_report.append(f"{prefix}{current_prefix}{item.name}")
            
            if item.is_dir():
                extension = "    " if is_last_item else "│   "
                self._print_tree(item, prefix + extension, is_last_item)
    
    def _add_complete_file_inventory(self):
        """Complete inventory of every single file."""
        self.markdown_report.append("## Complete File Inventory\n")
        
        all_files = sorted(self.extract_dir.rglob('*'))
        
        for file_path in all_files:
            if file_path.is_file():
                rel_path = file_path.relative_to(self.extract_dir)
                size = file_path.stat().st_size
                
                # Determine file type
                if file_path.suffix == '.xml':
                    file_type = "XML Document"
                elif file_path.suffix == '.rels':
                    file_type = "Relationships"
                elif file_path.suffix in ['.jpeg', '.jpg', '.png', '.gif']:
                    file_type = "Image"
                else:
                    file_type = "Other"
                
                self.markdown_report.append(f"### `{rel_path}`")
                self.markdown_report.append(f"- **Type:** {file_type}")
                self.markdown_report.append(f"- **Size:** {size:,} bytes ({size/1024:.2f} KB)")
                self.markdown_report.append("")
    
    def _parse_xml_with_namespaces(self, file_path):
        """Parse XML and return tree with namespace mapping."""
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        # Extract all namespaces
        namespaces = {}
        for event, elem in ET.iterparse(file_path, events=['start-ns']):
            prefix, uri = elem
            if prefix:
                namespaces[prefix] = uri
            else:
                namespaces['default'] = uri
        
        return tree, root, namespaces
    
    def _element_to_dict(self, element, namespaces):
        """Convert XML element to detailed dict representation."""
        result = {
            'tag': element.tag,
            'attributes': dict(element.attrib),
            'text': element.text.strip() if element.text and element.text.strip() else None,
            'tail': element.tail.strip() if element.tail and element.tail.strip() else None,
            'children': []
        }
        
        for child in element:
            result['children'].append(self._element_to_dict(child, namespaces))
        
        return result
    
    def _add_content_types_complete(self):
        """COMPLETE analysis of content types."""
        content_types_path = self.extract_dir / "[Content_Types].xml"
        
        if not content_types_path.exists():
            return
        
        self.markdown_report.append("## [Content_Types].xml - COMPLETE ANALYSIS\n")
        
        try:
            tree, root, namespaces = self._parse_xml_with_namespaces(content_types_path)
            
            self.markdown_report.append("### File Metadata")
            self.markdown_report.append(f"- **Size:** {content_types_path.stat().st_size:,} bytes")
            self.markdown_report.append(f"- **Root Element:** `{root.tag}`")
            self.markdown_report.append(f"- **Namespaces:** {namespaces}")
            self.markdown_report.append("")
            
            # Parse without namespace for easier reading
            for elem in root.iter():
                if '}' in elem.tag:
                    elem.tag = elem.tag.split('}', 1)[1]
            
            defaults = root.findall('.//Default')
            overrides = root.findall('.//Override')
            
            self.markdown_report.append(f"### Default Content Types ({len(defaults)} entries)\n")
            for i, default in enumerate(defaults, 1):
                ext = default.get('Extension')
                content_type = default.get('ContentType')
                self.markdown_report.append(f"{i}. **Extension:** `.{ext}`")
                self.markdown_report.append(f"   - **Content-Type:** `{content_type}`")
                self.markdown_report.append("")
            
            self.markdown_report.append(f"### Override Content Types ({len(overrides)} entries)\n")
            for i, override in enumerate(overrides, 1):
                part_name = override.get('PartName')
                content_type = override.get('ContentType')
                self.markdown_report.append(f"{i}. **Part:** `{part_name}`")
                self.markdown_report.append(f"   - **Content-Type:** `{content_type}`")
                self.markdown_report.append("")
        
        except Exception as e:
            self.markdown_report.append(f"Error: {e}\n")
    
    def _add_all_relationships(self):
        """COMPLETE analysis of ALL relationship files."""
        self.markdown_report.append("## Relationships - COMPLETE ANALYSIS\n")
        
        # Find all .rels files
        rels_files = list(self.extract_dir.rglob('*.rels'))
        
        for rels_file in sorted(rels_files):
            rel_path = rels_file.relative_to(self.extract_dir)
            self.markdown_report.append(f"### `{rel_path}`\n")
            
            try:
                tree, root, namespaces = self._parse_xml_with_namespaces(rels_file)
                
                self.markdown_report.append(f"**File Size:** {rels_file.stat().st_size:,} bytes")
                self.markdown_report.append(f"**Namespaces:** {namespaces}")
                self.markdown_report.append("")
                
                # Remove namespace for easier parsing
                for elem in root.iter():
                    if '}' in elem.tag:
                        elem.tag = elem.tag.split('}', 1)[1]
                
                relationships = root.findall('.//Relationship')
                
                self.markdown_report.append(f"**Total Relationships:** {len(relationships)}\n")
                
                for i, rel in enumerate(relationships, 1):
                    rel_id = rel.get('Id')
                    rel_type = rel.get('Type')
                    target = rel.get('Target')
                    target_mode = rel.get('TargetMode', 'Internal')
                    
                    self.markdown_report.append(f"{i}. **Relationship ID:** `{rel_id}`")
                    self.markdown_report.append(f"   - **Type:** `{rel_type}`")
                    self.markdown_report.append(f"   - **Target:** `{target}`")
                    self.markdown_report.append(f"   - **Target Mode:** `{target_mode}`")
                    self.markdown_report.append("")
            
            except Exception as e:
                self.markdown_report.append(f"Error parsing: {e}\n")
    
    def _add_document_xml_complete(self):
        """COMPLETE atomic-level analysis of document.xml."""
        doc_path = self.extract_dir / "word" / "document.xml"
        
        if not doc_path.exists():
            return
        
        self.markdown_report.append("## word/document.xml - COMPLETE ATOMIC ANALYSIS\n")
        
        try:
            tree, root, namespaces = self._parse_xml_with_namespaces(doc_path)
            
            self.markdown_report.append("### File Metadata")
            self.markdown_report.append(f"- **Size:** {doc_path.stat().st_size:,} bytes")
            self.markdown_report.append(f"- **Root Element:** `{root.tag}`")
            self.markdown_report.append(f"- **Namespaces:**")
            for prefix, uri in namespaces.items():
                self.markdown_report.append(f"  - `{prefix}`: `{uri}`")
            self.markdown_report.append("")
            
            # Register all namespaces for xpath queries
            for prefix, uri in namespaces.items():
                if prefix != 'default':
                    ET.register_namespace(prefix, uri)
            
            # Use the actual namespace prefixes
            w_ns = namespaces.get('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            ns = {'w': w_ns}
            
            # Get all major elements
            body = root.find('.//w:body', ns)
            paragraphs = root.findall('.//w:p', ns)
            tables = root.findall('.//w:tbl', ns)
            sections = root.findall('.//w:sectPr', ns)
            
            self.markdown_report.append("### Document Structure Overview")
            self.markdown_report.append(f"- **Body Element Present:** {'Yes' if body is not None else 'No'}")
            self.markdown_report.append(f"- **Total Paragraphs:** {len(paragraphs)}")
            self.markdown_report.append(f"- **Total Tables:** {len(tables)}")
            self.markdown_report.append(f"- **Total Sections:** {len(sections)}")
            self.markdown_report.append("")
            
            # Detailed paragraph analysis
            self.markdown_report.append(f"### Detailed Paragraph Analysis ({len(paragraphs)} paragraphs)\n")
            
            for i, para in enumerate(paragraphs, 1):
                self.markdown_report.append(f"#### Paragraph {i}\n")
                
                # Paragraph properties
                pPr = para.find('w:pPr', ns)
                if pPr is not None:
                    self.markdown_report.append("**Paragraph Properties:**")
                    for prop in pPr:
                        tag_name = prop.tag.split('}')[-1] if '}' in prop.tag else prop.tag
                        attrs = ', '.join([f"{k}={v}" for k, v in prop.attrib.items()])
                        self.markdown_report.append(f"- `{tag_name}` {f'({attrs})' if attrs else ''}")
                    self.markdown_report.append("")
                
                # Runs analysis
                runs = para.findall('w:r', ns)
                self.markdown_report.append(f"**Runs:** {len(runs)}")
                
                for j, run in enumerate(runs, 1):
                    self.markdown_report.append(f"\n**Run {j}:**")
                    
                    # Run properties
                    rPr = run.find('w:rPr', ns)
                    if rPr is not None:
                        self.markdown_report.append("- Properties:")
                        for prop in rPr:
                            tag_name = prop.tag.split('}')[-1] if '}' in prop.tag else prop.tag
                            attrs = ', '.join([f"{k}={v}" for k, v in prop.attrib.items()])
                            self.markdown_report.append(f"  - `{tag_name}` {f'({attrs})' if attrs else ''}")
                    
                    # Text content
                    texts = run.findall('w:t', ns)
                    for t in texts:
                        if t.text:
                            space_attr = t.get('{http://www.w3.org/XML/1998/namespace}space', '')
                            self.markdown_report.append(f"- Text: `{t.text}`")
                            if space_attr:
                                self.markdown_report.append(f"  - xml:space: `{space_attr}`")
                
                self.markdown_report.append("")
            
            # Detailed table analysis
            if tables:
                self.markdown_report.append(f"### Detailed Table Analysis ({len(tables)} tables)\n")
                
                for i, table in enumerate(tables, 1):
                    self.markdown_report.append(f"#### Table {i}\n")
                    
                    # Table properties
                    tblPr = table.find('w:tblPr', ns)
                    if tblPr is not None:
                        self.markdown_report.append("**Table Properties:**")
                        for prop in tblPr:
                            tag_name = prop.tag.split('}')[-1] if '}' in prop.tag else prop.tag
                            attrs = ', '.join([f"{k}={v}" for k, v in prop.attrib.items()])
                            self.markdown_report.append(f"- `{tag_name}` {f'({attrs})' if attrs else ''}")
                        self.markdown_report.append("")
                    
                    # Table grid
                    tblGrid = table.find('w:tblGrid', ns)
                    if tblGrid is not None:
                        grid_cols = tblGrid.findall('w:gridCol', ns)
                        self.markdown_report.append(f"**Table Grid:** {len(grid_cols)} columns")
                        for k, col in enumerate(grid_cols, 1):
                            width = col.get(f'{{{w_ns}}}w', 'auto')
                            self.markdown_report.append(f"- Column {k}: width = `{width}`")
                        self.markdown_report.append("")
                    
                    # Rows
                    rows = table.findall('w:tr', ns)
                    self.markdown_report.append(f"**Rows:** {len(rows)}\n")
                    
                    for r_idx, row in enumerate(rows, 1):
                        cells = row.findall('w:tc', ns)
                        self.markdown_report.append(f"**Row {r_idx}:** {len(cells)} cells")
                        
                        for c_idx, cell in enumerate(cells, 1):
                            # Cell properties
                            tcPr = cell.find('w:tcPr', ns)
                            cell_props = []
                            if tcPr is not None:
                                for prop in tcPr:
                                    tag_name = prop.tag.split('}')[-1] if '}' in prop.tag else prop.tag
                                    cell_props.append(tag_name)
                            
                            # Cell text
                            cell_paras = cell.findall('w:p', ns)
                            cell_text = []
                            for cp in cell_paras:
                                texts = cp.findall('.//w:t', ns)
                                para_text = ''.join([t.text for t in texts if t.text])
                                if para_text:
                                    cell_text.append(para_text)
                            
                            self.markdown_report.append(f"  - Cell {c_idx}: {', '.join(cell_props) if cell_props else 'no special properties'}")
                            if cell_text:
                                self.markdown_report.append(f"    - Text: `{' '.join(cell_text)}`")
                        
                        self.markdown_report.append("")
            
            # Section properties
            if sections:
                self.markdown_report.append(f"### Section Properties ({len(sections)} sections)\n")
                
                for i, section in enumerate(sections, 1):
                    self.markdown_report.append(f"#### Section {i}\n")
                    
                    for prop in section:
                        tag_name = prop.tag.split('}')[-1] if '}' in prop.tag else prop.tag
                        attrs = dict(prop.attrib)
                        
                        self.markdown_report.append(f"**{tag_name}:**")
                        if attrs:
                            for k, v in attrs.items():
                                attr_name = k.split('}')[-1] if '}' in k else k
                                self.markdown_report.append(f"- {attr_name}: `{v}`")
                        
                        # Check for child elements
                        if len(prop) > 0:
                            self.markdown_report.append("- Child elements:")
                            for child in prop:
                                child_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                                child_attrs = ', '.join([f"{k.split('}')[-1]}={v}" for k, v in child.attrib.items()])
                                self.markdown_report.append(f"  - `{child_name}` {f'({child_attrs})' if child_attrs else ''}")
                        
                        self.markdown_report.append("")
        
        except Exception as e:
            self.markdown_report.append(f"Error: {e}\n")
            import traceback
            self.markdown_report.append(f"```\n{traceback.format_exc()}\n```\n")
    
    def _add_styles_xml_complete(self):
        """COMPLETE analysis of styles.xml."""
        styles_path = self.extract_dir / "word" / "styles.xml"
        
        if not styles_path.exists():
            return
        
        self.markdown_report.append("## word/styles.xml - COMPLETE ANALYSIS\n")
        
        try:
            tree, root, namespaces = self._parse_xml_with_namespaces(styles_path)
            
            self.markdown_report.append("### File Metadata")
            self.markdown_report.append(f"- **Size:** {styles_path.stat().st_size:,} bytes")
            self.markdown_report.append("")
            
            w_ns = namespaces.get('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            ns = {'w': w_ns}
            
            # Get all styles
            styles = root.findall('.//w:style', ns)
            
            self.markdown_report.append(f"### Total Styles: {len(styles)}\n")
            
            for i, style in enumerate(styles, 1):
                style_type = style.get(f'{{{w_ns}}}type', 'unknown')
                style_id = style.get(f'{{{w_ns}}}styleId', 'unknown')
                default = style.get(f'{{{w_ns}}}default', '0')
                custom_style = style.get(f'{{{w_ns}}}customStyle', '0')
                
                self.markdown_report.append(f"#### Style {i}: `{style_id}`\n")
                self.markdown_report.append(f"- **Type:** `{style_type}`")
                self.markdown_report.append(f"- **Default:** `{default}`")
                self.markdown_report.append(f"- **Custom:** `{custom_style}`")
                
                # Style name
                name_elem = style.find('w:name', ns)
                if name_elem is not None:
                    self.markdown_report.append(f"- **Name:** `{name_elem.get(f'{{{w_ns}}}val', 'N/A')}`")
                
                # Based on
                based_on = style.find('w:basedOn', ns)
                if based_on is not None:
                    self.markdown_report.append(f"- **Based On:** `{based_on.get(f'{{{w_ns}}}val', 'N/A')}`")
                
                # Next style
                next_style = style.find('w:next', ns)
                if next_style is not None:
                    self.markdown_report.append(f"- **Next:** `{next_style.get(f'{{{w_ns}}}val', 'N/A')}`")
                
                # UI Priority
                ui_priority = style.find('w:uiPriority', ns)
                if ui_priority is not None:
                    self.markdown_report.append(f"- **UI Priority:** `{ui_priority.get(f'{{{w_ns}}}val', 'N/A')}`")
                
                # Properties
                self.markdown_report.append("\n**Properties:**")
                for child in style:
                    tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if tag_name not in ['name', 'basedOn', 'next', 'uiPriority']:
                        attrs = ', '.join([f"{k.split('}')[-1]}={v}" for k, v in child.attrib.items()])
                        self.markdown_report.append(f"- `{tag_name}` {f'({attrs})' if attrs else ''}")
                
                self.markdown_report.append("")
        
        except Exception as e:
            self.markdown_report.append(f"Error: {e}\n")
    
    def _add_settings_xml_complete(self):
        """COMPLETE analysis of settings.xml."""
        settings_path = self.extract_dir / "word" / "settings.xml"
        
        if not settings_path.exists():
            return
        
        self.markdown_report.append("## word/settings.xml - COMPLETE ANALYSIS\n")
        
        try:
            tree, root, namespaces = self._parse_xml_with_namespaces(settings_path)
            
            self.markdown_report.append("### File Metadata")
            self.markdown_report.append(f"- **Size:** {settings_path.stat().st_size:,} bytes")
            self.markdown_report.append("")
            
            self.markdown_report.append("### All Settings\n")
            
            for child in root:
                tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                attrs = dict(child.attrib)
                
                self.markdown_report.append(f"**{tag_name}:**")
                
                if attrs:
                    for k, v in attrs.items():
                        attr_name = k.split('}')[-1] if '}' in k else k
                        self.markdown_report.append(f"- {attr_name}: `{v}`")
                
                if child.text and child.text.strip():
                    self.markdown_report.append(f"- Text: `{child.text.strip()}`")
                
                if len(child) > 0:
                    self.markdown_report.append("- Child elements:")
                    for subchild in child:
                        subchild_name = subchild.tag.split('}')[-1] if '}' in subchild.tag else subchild.tag
                        subchild_attrs = ', '.join([f"{k.split('}')[-1]}={v}" for k, v in subchild.attrib.items()])
                        self.markdown_report.append(f"  - `{subchild_name}` {f'({subchild_attrs})' if subchild_attrs else ''}")
                
                self.markdown_report.append("")
        
        except Exception as e:
            self.markdown_report.append(f"Error: {e}\n")
    
    def _add_font_table_complete(self):
        """COMPLETE analysis of fontTable.xml."""
        font_path = self.extract_dir / "word" / "fontTable.xml"
        
        if not font_path.exists():
            return
        
        self.markdown_report.append("## word/fontTable.xml - COMPLETE ANALYSIS\n")
        
        try:
            tree, root, namespaces = self._parse_xml_with_namespaces(font_path)
            
            self.markdown_report.append("### File Metadata")
            self.markdown_report.append(f"- **Size:** {font_path.stat().st_size:,} bytes")
            self.markdown_report.append("")
            
            w_ns = namespaces.get('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            ns = {'w': w_ns}
            
            fonts = root.findall('.//w:font', ns)
            
            self.markdown_report.append(f"### Total Fonts: {len(fonts)}\n")
            
            for i, font in enumerate(fonts, 1):
                font_name = font.get(f'{{{w_ns}}}name', 'unknown')
                
                self.markdown_report.append(f"#### Font {i}: `{font_name}`\n")
                
                for child in font:
                    tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    attrs = ', '.join([f"{k.split('}')[-1]}={v}" for k, v in child.attrib.items()])
                    self.markdown_report.append(f"- **{tag_name}:** {attrs if attrs else '(no attributes)'}")
                
                self.markdown_report.append("")
        
        except Exception as e:
            self.markdown_report.append(f"Error: {e}\n")
    
    def _add_numbering_complete(self):
        """COMPLETE analysis of numbering.xml."""
        numbering_path = self.extract_dir / "word" / "numbering.xml"
        
        if not numbering_path.exists():
            return
        
        self.markdown_report.append("## word/numbering.xml - COMPLETE ANALYSIS\n")
        
        try:
            tree, root, namespaces = self._parse_xml_with_namespaces(numbering_path)
            
            self.markdown_report.append("### File Metadata")
            self.markdown_report.append(f"- **Size:** {numbering_path.stat().st_size:,} bytes")
            self.markdown_report.append("")
            
            w_ns = namespaces.get('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            ns = {'w': w_ns}
            
            abstract_nums = root.findall('.//w:abstractNum', ns)
            num_defs = root.findall('.//w:num', ns)
            
            self.markdown_report.append(f"### Abstract Numbering Definitions: {len(abstract_nums)}\n")
            
            for i, abs_num in enumerate(abstract_nums, 1):
                abs_num_id = abs_num.get(f'{{{w_ns}}}abstractNumId', 'unknown')
                
                self.markdown_report.append(f"#### Abstract Num {i} (ID: {abs_num_id})\n")
                
                for child in abs_num:
                    tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    self.markdown_report.append(f"**{tag_name}:**")
                    
                    for k, v in child.attrib.items():
                        attr_name = k.split('}')[-1] if '}' in k else k
                        self.markdown_report.append(f"- {attr_name}: `{v}`")
                    
                    if len(child) > 0:
                        for subchild in child:
                            subchild_name = subchild.tag.split('}')[-1] if '}' in subchild.tag else subchild.tag
                            subchild_attrs = ', '.join([f"{k.split('}')[-1]}={v}" for k, v in subchild.attrib.items()])
                            self.markdown_report.append(f"  - `{subchild_name}` {f'({subchild_attrs})' if subchild_attrs else ''}")
                    
                    self.markdown_report.append("")
            
            self.markdown_report.append(f"### Numbering Instances: {len(num_defs)}\n")
            
            for i, num in enumerate(num_defs, 1):
                num_id = num.get(f'{{{w_ns}}}numId', 'unknown')
                
                self.markdown_report.append(f"#### Numbering {i} (ID: {num_id})\n")
                
                abstract_num_id = num.find('w:abstractNumId', ns)
                if abstract_num_id is not None:
                    self.markdown_report.append(f"- **References Abstract Num:** `{abstract_num_id.get(f'{{{w_ns}}}val', 'N/A')}`")
                
                self.markdown_report.append("")
        
        except Exception as e:
            self.markdown_report.append(f"Error: {e}\n")
    
    def _add_theme_complete(self):
        """COMPLETE analysis of theme files."""
        theme_dir = self.extract_dir / "word" / "theme"
        
        if not theme_dir.exists():
            return
        
        self.markdown_report.append("## word/theme/ - COMPLETE ANALYSIS\n")
        
        theme_files = list(theme_dir.glob('*.xml'))
        
        for theme_file in sorted(theme_files):
            rel_path = theme_file.relative_to(self.extract_dir)
            self.markdown_report.append(f"### `{rel_path}`\n")
            
            try:
                tree, root, namespaces = self._parse_xml_with_namespaces(theme_file)
                
                self.markdown_report.append(f"**Size:** {theme_file.stat().st_size:,} bytes")
                self.markdown_report.append(f"**Root Element:** `{root.tag}`")
                self.markdown_report.append("")
                
                # Recursively document all elements
                self._document_element_recursive(root, 0)
                
                self.markdown_report.append("")
            
            except Exception as e:
                self.markdown_report.append(f"Error: {e}\n")
    
    def _document_element_recursive(self, element, depth, max_depth=5):
        """Recursively document an XML element and its children."""
        if depth > max_depth:
            return
        
        indent = "  " * depth
        tag_name = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        attrs = ', '.join([f"{k.split('}')[-1]}={v}" for k, v in element.attrib.items()])
        
        self.markdown_report.append(f"{indent}- **{tag_name}** {f'({attrs})' if attrs else ''}")
        
        if element.text and element.text.strip():
            self.markdown_report.append(f"{indent}  - Text: `{element.text.strip()[:100]}`")
        
        for child in element:
            self._document_element_recursive(child, depth + 1, max_depth)
    
    def _add_doc_properties_complete(self):
        """COMPLETE analysis of document properties."""
        self.markdown_report.append("## Document Properties - COMPLETE ANALYSIS\n")
        
        # Core properties
        core_path = self.extract_dir / "docProps" / "core.xml"
        if core_path.exists():
            self.markdown_report.append("### docProps/core.xml\n")
            
            try:
                tree, root, namespaces = self._parse_xml_with_namespaces(core_path)
                
                self.markdown_report.append(f"**Size:** {core_path.stat().st_size:,} bytes")
                self.markdown_report.append("")
                
                for child in root:
                    tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    text = child.text.strip() if child.text else 'N/A'
                    attrs = ', '.join([f"{k.split('}')[-1]}={v}" for k, v in child.attrib.items()])
                    
                    self.markdown_report.append(f"**{tag_name}:** `{text}` {f'({attrs})' if attrs else ''}")
                
                self.markdown_report.append("")
            
            except Exception as e:
                self.markdown_report.append(f"Error: {e}\n")
        
        # App properties
        app_path = self.extract_dir / "docProps" / "app.xml"
        if app_path.exists():
            self.markdown_report.append("### docProps/app.xml\n")
            
            try:
                tree, root, namespaces = self._parse_xml_with_namespaces(app_path)
                
                self.markdown_report.append(f"**Size:** {app_path.stat().st_size:,} bytes")
                self.markdown_report.append("")
                
                for child in root:
                    tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    text = child.text.strip() if child.text else 'N/A'
                    
                    self.markdown_report.append(f"**{tag_name}:** `{text}`")
                
                self.markdown_report.append("")
            
            except Exception as e:
                self.markdown_report.append(f"Error: {e}\n")
    
    def _add_custom_xml_complete(self):
        """COMPLETE analysis of custom XML."""
        custom_dir = self.extract_dir / "customXml"
        
        if not custom_dir.exists():
            return
        
        self.markdown_report.append("## customXml/ - COMPLETE ANALYSIS\n")
        
        xml_files = list(custom_dir.glob('*.xml'))
        
        for xml_file in sorted(xml_files):
            rel_path = xml_file.relative_to(self.extract_dir)
            self.markdown_report.append(f"### `{rel_path}`\n")
            
            try:
                tree, root, namespaces = self._parse_xml_with_namespaces(xml_file)
                
                self.markdown_report.append(f"**Size:** {xml_file.stat().st_size:,} bytes")
                self.markdown_report.append(f"**Root Element:** `{root.tag}`")
                self.markdown_report.append("")
                
                self._document_element_recursive(root, 0, max_depth=10)
                
                self.markdown_report.append("")
            
            except Exception as e:
                self.markdown_report.append(f"Error: {e}\n")
    
    def _add_web_settings_complete(self):
        """COMPLETE analysis of webSettings.xml."""
        web_path = self.extract_dir / "word" / "webSettings.xml"
        
        if not web_path.exists():
            return
        
        self.markdown_report.append("## word/webSettings.xml - COMPLETE ANALYSIS\n")
        
        try:
            tree, root, namespaces = self._parse_xml_with_namespaces(web_path)
            
            self.markdown_report.append(f"**Size:** {web_path.stat().st_size:,} bytes")
            self.markdown_report.append("")
            
            for child in root:
                tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                attrs = ', '.join([f"{k.split('}')[-1]}={v}" for k, v in child.attrib.items()])
                
                self.markdown_report.append(f"**{tag_name}:** {attrs if attrs else '(no attributes)'}")
            
            self.markdown_report.append("")
        
        except Exception as e:
            self.markdown_report.append(f"Error: {e}\n")
    
    def _add_other_xml_files(self):
        """Analyze any other XML files not covered."""
        self.markdown_report.append("## Other XML Files - COMPLETE ANALYSIS\n")
        
        covered_files = {
            'document.xml', 'styles.xml', 'settings.xml', 'fontTable.xml',
            'numbering.xml', 'webSettings.xml', 'stylesWithEffects.xml',
            'core.xml', 'app.xml', '[Content_Types].xml'
        }
        
        all_xml = list(self.extract_dir.rglob('*.xml'))
        other_xml = [f for f in all_xml if f.name not in covered_files and 'theme' not in str(f) and 'customXml' not in str(f)]
        
        if not other_xml:
            self.markdown_report.append("No other XML files found.\n")
            return
        
        for xml_file in sorted(other_xml):
            rel_path = xml_file.relative_to(self.extract_dir)
            self.markdown_report.append(f"### `{rel_path}`\n")
            
            try:
                tree, root, namespaces = self._parse_xml_with_namespaces(xml_file)
                
                self.markdown_report.append(f"**Size:** {xml_file.stat().st_size:,} bytes")
                self.markdown_report.append(f"**Root Element:** `{root.tag}`")
                self.markdown_report.append(f"**Namespaces:** {namespaces}")
                self.markdown_report.append("")
                
                self._document_element_recursive(root, 0, max_depth=10)
                
                self.markdown_report.append("")
            
            except Exception as e:
                self.markdown_report.append(f"Error: {e}\n")
    
    def _add_binary_files(self):
        """Analyze binary files (images, etc.)."""
        self.markdown_report.append("## Binary Files Analysis\n")
        
        binary_extensions = {'.jpeg', '.jpg', '.png', '.gif', '.bmp', '.tiff', '.emf', '.wmf'}
        all_files = list(self.extract_dir.rglob('*'))
        binary_files = [f for f in all_files if f.is_file() and f.suffix.lower() in binary_extensions]
        
        if not binary_files:
            self.markdown_report.append("No binary files found.\n")
            return
        
        for bin_file in sorted(binary_files):
            rel_path = bin_file.relative_to(self.extract_dir)
            size = bin_file.stat().st_size
            
            self.markdown_report.append(f"### `{rel_path}`")
            self.markdown_report.append(f"- **Type:** {bin_file.suffix.upper()}")
            self.markdown_report.append(f"- **Size:** {size:,} bytes ({size/1024:.2f} KB)")
            
            # Read file signature (magic bytes)
            with open(bin_file, 'rb') as f:
                magic = f.read(16)
                hex_magic = ' '.join([f'{b:02x}' for b in magic])
                self.markdown_report.append(f"- **Magic Bytes:** `{hex_magic}`")
            
            self.markdown_report.append("")
    
    def _add_raw_xml_dumps(self):
        """Add complete raw XML dumps for all XML files."""
        self.markdown_report.append("## RAW XML DUMPS\n")
        self.markdown_report.append("Complete, unprocessed XML content for every XML file.\n")
        
        all_xml = sorted(self.extract_dir.rglob('*.xml'))
        
        for xml_file in all_xml:
            rel_path = xml_file.relative_to(self.extract_dir)
            self.markdown_report.append(f"### `{rel_path}` - RAW XML\n")
            
            try:
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                self.markdown_report.append("```xml")
                self.markdown_report.append(content)
                self.markdown_report.append("```\n")
            
            except Exception as e:
                self.markdown_report.append(f"Error reading file: {e}\n")
    
    def save_analysis(self, output_path=None):
        """
        Save the markdown analysis to a file.
        
        Args:
            output_path: Path to save the markdown file. If None, uses default name.
        
        Returns:
            Path to the saved markdown file
        """
        if not self.markdown_report:
            self.analyze_structure()
        
        if output_path is None:
            output_path = self.extract_dir.parent / f"{self.extract_dir.name}_analysis.md"
        else:
            output_path = Path(output_path)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(self.markdown_report))
        
        print(f"Analysis saved to: {output_path}")
        return output_path
    
    def reconstruct(self, output_path=None):
        """
        Reconstruct the .docx file from the extracted components.
        
        Args:
            output_path: Path for the reconstructed .docx file. If None, uses default name.
        
        Returns:
            Path to the reconstructed .docx file
        """
        if self.extract_dir is None:
            raise ValueError("Must call extract() before reconstruct()")
        
        if output_path is None:
            output_path = self.extract_dir.parent / f"{self.extract_dir.name}_reconstructed.docx"
        else:
            output_path = Path(output_path)
        
        print(f"Reconstructing document from {self.extract_dir}...")
        
        # Create a new ZIP file
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            # Walk through all files in the extracted directory
            for file_path in self.extract_dir.rglob('*'):
                if file_path.is_file():
                    # Get the relative path for the archive
                    arcname = file_path.relative_to(self.extract_dir)
                    docx.write(file_path, arcname)
        
        print(f"Reconstruction complete: {output_path}")
        return output_path


def main():
    """Main function demonstrating usage."""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python python docx_obliterate.py <path_to_docx>")
        print("\nExample: python python docx_obliterate.py MECH_SPEC.docx")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    
    if not os.path.exists(docx_path):
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)
    
    # Create decomposer
    decomposer = DocxDecomposer(docx_path)
    
    # Extract the document
    extract_dir = decomposer.extract()
    
    # Analyze and save report
    analysis_path = decomposer.save_analysis()
    

    
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(f"Original document:     {docx_path}")
    print(f"Extracted to:          {extract_dir}")
    print(f"Analysis report:       {analysis_path}")


if __name__ == "__main__":
    main()
