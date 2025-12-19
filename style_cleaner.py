"""
SpecCleanse Style Cleaner Module

Detects and removes unused styles from DOCX documents.
Styles are "unused" if they're not referenced anywhere in the document content
and not required by other styles (via basedOn, link, next).
"""

from pathlib import Path
from dataclasses import dataclass, field
from lxml import etree

# Namespaces
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

# Built-in styles that should never be removed
PROTECTED_STYLES = {
    "Normal",
    "DefaultParagraphFont", 
    "TableNormal",
    "NoList",
    "Heading1", "Heading2", "Heading3", "Heading4", "Heading5",
    "Heading6", "Heading7", "Heading8", "Heading9",
    "Title",
    "Subtitle",
    "ListParagraph",
    "TOCHeading",
    "TOC1", "TOC2", "TOC3", "TOC4", "TOC5",
    "Header",
    "Footer",
    "FootnoteText",
    "EndnoteText",
    "CommentText",
    "BalloonText",
    "Hyperlink",
}


@dataclass
class StyleInfo:
    """Information about a style definition."""
    style_id: str
    name: str
    style_type: str  # paragraph, character, table, numbering
    based_on: str | None = None
    linked: str | None = None
    next_style: str | None = None
    is_builtin: bool = False
    is_default: bool = False


@dataclass
class StyleCleanResult:
    """Results from style cleaning operation."""
    total_styles: int = 0
    used_styles: set = field(default_factory=set)
    unused_styles: set = field(default_factory=set)
    protected_styles: set = field(default_factory=set)
    removed_styles: list = field(default_factory=list)
    errors: list = field(default_factory=list)


class StyleCleaner:
    """
    Analyzes and removes unused styles from DOCX documents.
    
    Strategy:
    1. Parse styles.xml to get all defined styles
    2. Scan all content XML files for style references
    3. Build dependency graph (basedOn, link, next)
    4. Mark styles as used via transitive closure
    5. Remove unused styles from styles.xml
    """
    
    def __init__(self, verbose: bool = False):
        self.verbose = verbose
    
    def analyze(self, unpacked_dir: Path) -> StyleCleanResult:
        """
        Analyze styles without modifying.
        Returns which styles are used/unused.
        """
        result = StyleCleanResult()
        
        styles_path = unpacked_dir / "word" / "styles.xml"
        if not styles_path.exists():
            result.errors.append("No styles.xml found")
            return result
        
        # Parse styles
        parser = etree.XMLParser(remove_blank_text=False)
        styles_tree = etree.parse(str(styles_path), parser)
        styles_root = styles_tree.getroot()
        
        # Get all defined styles
        defined_styles = self._get_defined_styles(styles_root)
        result.total_styles = len(defined_styles)
        
        if self.verbose:
            print(f"  Found {len(defined_styles)} defined styles")
        
        # Find directly used styles in content
        directly_used = self._find_used_styles(unpacked_dir)
        
        if self.verbose:
            print(f"  Found {len(directly_used)} directly used styles")
        
        # Build dependency graph and find all required styles
        all_used = self._expand_dependencies(defined_styles, directly_used)
        result.used_styles = all_used
        
        # Determine unused styles
        all_style_ids = set(defined_styles.keys())
        result.unused_styles = all_style_ids - all_used
        
        # Filter out protected styles
        result.protected_styles = result.unused_styles & PROTECTED_STYLES
        result.unused_styles = result.unused_styles - PROTECTED_STYLES
        
        # Also protect default styles
        for style_id, info in defined_styles.items():
            if info.is_default and style_id in result.unused_styles:
                result.protected_styles.add(style_id)
                result.unused_styles.discard(style_id)
        
        if self.verbose:
            print(f"  {len(result.unused_styles)} styles can be removed")
            print(f"  {len(result.protected_styles)} unused but protected")
        
        return result
    
    def clean(self, unpacked_dir: Path, dry_run: bool = False) -> StyleCleanResult:
        """
        Remove unused styles from styles.xml.
        """
        result = self.analyze(unpacked_dir)
        
        if result.errors or dry_run:
            return result
        
        if not result.unused_styles:
            return result
        
        # Remove unused styles from styles.xml
        styles_path = unpacked_dir / "word" / "styles.xml"
        parser = etree.XMLParser(remove_blank_text=False)
        styles_tree = etree.parse(str(styles_path), parser)
        styles_root = styles_tree.getroot()
        
        for style_elem in styles_root.findall(f"{W}style"):
            style_id = style_elem.get(f"{W}styleId")
            if style_id in result.unused_styles:
                styles_root.remove(style_elem)
                result.removed_styles.append(style_id)
                if self.verbose:
                    print(f"  Removed style: {style_id}")
        
        # Write back
        styles_tree.write(
            str(styles_path), 
            xml_declaration=True, 
            encoding="UTF-8", 
            standalone=True
        )
        
        return result
    
    def _get_defined_styles(self, styles_root: etree._Element) -> dict[str, StyleInfo]:
        """Extract all style definitions from styles.xml."""
        styles = {}
        
        for style_elem in styles_root.findall(f"{W}style"):
            style_id = style_elem.get(f"{W}styleId")
            if not style_id:
                continue
            
            # Get style name
            name_elem = style_elem.find(f"{W}name")
            name = name_elem.get(f"{W}val") if name_elem is not None else style_id
            
            # Get style type
            style_type = style_elem.get(f"{W}type", "paragraph")
            
            # Get dependencies
            based_on = None
            based_on_elem = style_elem.find(f"{W}basedOn")
            if based_on_elem is not None:
                based_on = based_on_elem.get(f"{W}val")
            
            linked = None
            link_elem = style_elem.find(f"{W}link")
            if link_elem is not None:
                linked = link_elem.get(f"{W}val")
            
            next_style = None
            next_elem = style_elem.find(f"{W}next")
            if next_elem is not None:
                next_style = next_elem.get(f"{W}val")
            
            # Check if default
            is_default = style_elem.get(f"{W}default") == "1"
            
            styles[style_id] = StyleInfo(
                style_id=style_id,
                name=name,
                style_type=style_type,
                based_on=based_on,
                linked=linked,
                next_style=next_style,
                is_builtin=style_id in PROTECTED_STYLES,
                is_default=is_default,
            )
        
        return styles
    
    def _find_used_styles(self, unpacked_dir: Path) -> set[str]:
        """Find all styles directly referenced in document content."""
        used = set()
        word_dir = unpacked_dir / "word"
        
        # Files to check for style references
        content_files = [
            word_dir / "document.xml",
            word_dir / "footnotes.xml",
            word_dir / "endnotes.xml", 
            word_dir / "comments.xml",
        ]
        
        # Add headers and footers
        content_files.extend(word_dir.glob("header*.xml"))
        content_files.extend(word_dir.glob("footer*.xml"))
        
        for content_file in content_files:
            if content_file.exists():
                used.update(self._find_styles_in_file(content_file))
        
        return used
    
    def _find_styles_in_file(self, file_path: Path) -> set[str]:
        """Find style references in a single XML file."""
        used = set()
        
        try:
            parser = etree.XMLParser(remove_blank_text=False)
            tree = etree.parse(str(file_path), parser)
            root = tree.getroot()
            
            # Paragraph styles
            for pstyle in root.iter(f"{W}pStyle"):
                val = pstyle.get(f"{W}val")
                if val:
                    used.add(val)
            
            # Character/run styles
            for rstyle in root.iter(f"{W}rStyle"):
                val = rstyle.get(f"{W}val")
                if val:
                    used.add(val)
            
            # Table styles
            for tblstyle in root.iter(f"{W}tblStyle"):
                val = tblstyle.get(f"{W}val")
                if val:
                    used.add(val)
            
            # Numbering styles (in numPr)
            for numstyle in root.iter(f"{W}numStyleLink"):
                val = numstyle.get(f"{W}val")
                if val:
                    used.add(val)
                    
        except Exception as e:
            if self.verbose:
                print(f"  Warning: Could not parse {file_path}: {e}")
        
        return used
    
    def _expand_dependencies(
        self, 
        defined_styles: dict[str, StyleInfo], 
        directly_used: set[str]
    ) -> set[str]:
        """
        Expand used styles to include all dependencies.
        Uses transitive closure over basedOn, link, and next relationships.
        """
        all_used = set(directly_used)
        
        # Keep expanding until no new styles are added
        changed = True
        while changed:
            changed = False
            for style_id in list(all_used):
                if style_id not in defined_styles:
                    continue
                    
                info = defined_styles[style_id]
                
                # Add basedOn dependency
                if info.based_on and info.based_on not in all_used:
                    all_used.add(info.based_on)
                    changed = True
                
                # Add linked style
                if info.linked and info.linked not in all_used:
                    all_used.add(info.linked)
                    changed = True
                
                # Add next style
                if info.next_style and info.next_style not in all_used:
                    all_used.add(info.next_style)
                    changed = True
        
        return all_used
