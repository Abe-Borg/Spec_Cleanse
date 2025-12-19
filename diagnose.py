#!/usr/bin/env python3
"""
SpecCleanse Diagnostic Tool

Inspects a DOCX file to show formatting details for each paragraph,
helping identify how specifier notes and editorial content are formatted.
"""

import argparse
import zipfile
import sys
from pathlib import Path
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def get_run_formatting(run: etree._Element) -> dict:
    """Extract all formatting from a run."""
    fmt = {
        "bold": False,
        "italic": False,
        "hidden": False,
        "color": None,
        "highlight": None,
        "font": None,
        "size": None,
        "style": None,
    }
    
    rpr = run.find(f"{W}rPr")
    if rpr is None:
        return fmt
    
    if rpr.find(f"{W}b") is not None:
        fmt["bold"] = True
    if rpr.find(f"{W}i") is not None:
        fmt["italic"] = True
    if rpr.find(f"{W}vanish") is not None:
        fmt["hidden"] = True
    
    color = rpr.find(f"{W}color")
    if color is not None:
        fmt["color"] = color.get(f"{W}val")
    
    highlight = rpr.find(f"{W}highlight")
    if highlight is not None:
        fmt["highlight"] = highlight.get(f"{W}val")
    
    fonts = rpr.find(f"{W}rFonts")
    if fonts is not None:
        fmt["font"] = fonts.get(f"{W}ascii") or fonts.get(f"{W}hAnsi")
    
    sz = rpr.find(f"{W}sz")
    if sz is not None:
        # Size is in half-points
        half_pts = sz.get(f"{W}val")
        if half_pts:
            fmt["size"] = int(half_pts) / 2
    
    style = rpr.find(f"{W}rStyle")
    if style is not None:
        fmt["style"] = style.get(f"{W}val")
    
    return fmt


def get_paragraph_style(para: etree._Element) -> str | None:
    """Get paragraph style name."""
    ppr = para.find(f"{W}pPr")
    if ppr is None:
        return None
    pstyle = ppr.find(f"{W}pStyle")
    if pstyle is None:
        return None
    return pstyle.get(f"{W}val")


def get_paragraph_text(para: etree._Element) -> str:
    """Get all text from paragraph."""
    texts = []
    for t in para.iter(f"{W}t"):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def analyze_document(docx_path: Path, search_text: str = None, show_all: bool = False):
    """Analyze document and show formatting details."""
    
    with zipfile.ZipFile(docx_path, 'r') as zf:
        doc_xml = zf.read('word/document.xml')
    
    root = etree.fromstring(doc_xml)
    
    print(f"\nAnalyzing: {docx_path}")
    print("=" * 70)
    
    para_count = 0
    for para in root.iter(f"{W}p"):
        para_count += 1
        text = get_paragraph_text(para)
        
        if not text.strip():
            continue
        
        # Filter by search text if provided
        if search_text and search_text.lower() not in text.lower():
            continue
        
        # Get paragraph style
        para_style = get_paragraph_style(para)
        
        # Get formatting from all runs
        run_formats = []
        for run in para.iter(f"{W}r"):
            run_text = "".join(t.text or "" for t in run.iter(f"{W}t"))
            if run_text.strip():
                fmt = get_run_formatting(run)
                run_formats.append((run_text[:30], fmt))
        
        # Check if any run has notable formatting
        has_hidden = any(rf[1]["hidden"] for rf in run_formats)
        has_color = any(rf[1]["color"] for rf in run_formats)
        has_italic = any(rf[1]["italic"] for rf in run_formats)
        has_char_style = any(rf[1]["style"] for rf in run_formats)
        
        # Show if it has interesting formatting or matches search
        if show_all or search_text or has_hidden or has_color or has_char_style:
            preview = text[:80] + "..." if len(text) > 80 else text
            print(f"\n[Para {para_count}] {preview}")
            print(f"  Paragraph Style: {para_style or '(none)'}")
            
            if has_hidden:
                print(f"  ⚠️  HIDDEN TEXT DETECTED")
            
            # Show unique formatting found
            colors = set(rf[1]["color"] for rf in run_formats if rf[1]["color"])
            styles = set(rf[1]["style"] for rf in run_formats if rf[1]["style"])
            
            if colors:
                print(f"  Colors: {', '.join(colors)}")
            if styles:
                print(f"  Character Styles: {', '.join(styles)}")
            if has_italic:
                print(f"  Has italic text")
            
            # Show first run's full formatting as example
            if run_formats and (show_all or search_text):
                print(f"  First run formatting: {run_formats[0][1]}")
    
    print(f"\n\nTotal paragraphs: {para_count}")


def find_editorial_content(docx_path: Path):
    """Find likely editorial content based on common patterns."""
    
    with zipfile.ZipFile(docx_path, 'r') as zf:
        doc_xml = zf.read('word/document.xml')
    
    root = etree.fromstring(doc_xml)
    
    print(f"\nScanning for editorial content: {docx_path}")
    print("=" * 70)
    
    editorial_keywords = [
        "retain", "delete", "verify", "coordinate", "revise",
        "edit", "specifier", "architect", "section title",
        "project-specific", "inserting text", "subparagraph below"
    ]
    
    findings = []
    
    for para in root.iter(f"{W}p"):
        text = get_paragraph_text(para)
        text_lower = text.lower()
        
        if not text.strip():
            continue
        
        # Check for editorial keywords
        matched_keywords = [kw for kw in editorial_keywords if kw in text_lower]
        
        if matched_keywords:
            para_style = get_paragraph_style(para)
            
            # Get formatting
            has_hidden = False
            colors = set()
            char_styles = set()
            is_italic = False
            
            for run in para.iter(f"{W}r"):
                fmt = get_run_formatting(run)
                if fmt["hidden"]:
                    has_hidden = True
                if fmt["color"]:
                    colors.add(fmt["color"])
                if fmt["style"]:
                    char_styles.add(fmt["style"])
                if fmt["italic"]:
                    is_italic = True
            
            findings.append({
                "text": text,
                "keywords": matched_keywords,
                "para_style": para_style,
                "hidden": has_hidden,
                "colors": colors,
                "char_styles": char_styles,
                "italic": is_italic,
            })
    
    # Print findings grouped by formatting
    print(f"\nFound {len(findings)} paragraphs with editorial keywords:\n")
    
    for i, f in enumerate(findings, 1):
        preview = f["text"][:70] + "..." if len(f["text"]) > 70 else f["text"]
        print(f"{i}. \"{preview}\"")
        print(f"   Keywords: {', '.join(f['keywords'])}")
        print(f"   Para Style: {f['para_style'] or '(default)'}")
        if f["hidden"]:
            print(f"   ⚠️  HIDDEN (w:vanish)")
        if f["colors"]:
            print(f"   Colors: {', '.join(f['colors'])}")
        if f["char_styles"]:
            print(f"   Char Styles: {', '.join(f['char_styles'])}")
        if f["italic"]:
            print(f"   Italic: Yes")
        print()
    
    # Summary
    print("\n" + "=" * 70)
    print("SUMMARY - Common formatting for editorial content:")
    print("=" * 70)
    
    all_colors = set()
    all_styles = set()
    all_para_styles = set()
    hidden_count = 0
    italic_count = 0
    
    for f in findings:
        all_colors.update(f["colors"])
        all_styles.update(f["char_styles"])
        if f["para_style"]:
            all_para_styles.add(f["para_style"])
        if f["hidden"]:
            hidden_count += 1
        if f["italic"]:
            italic_count += 1
    
    print(f"\nParagraph styles used: {all_para_styles or '(none)'}")
    print(f"Character styles used: {all_styles or '(none)'}")
    print(f"Colors used: {all_colors or '(none)'}")
    print(f"Hidden text count: {hidden_count}")
    print(f"Italic count: {italic_count}")
    
    if all_colors:
        print(f"\n💡 Suggestion: Add these colors to patterns.yaml under")
        print(f"   specifier_notes.formatting_signals.colors:")
        for c in all_colors:
            print(f"     - \"{c}\"")
    
    if all_para_styles:
        print(f"\n💡 Suggestion: Add these to style_based_detection.paragraph_styles:")
        for s in all_para_styles:
            print(f"     - \"{s}\"")


def main():
    parser = argparse.ArgumentParser(
        description="Diagnose DOCX formatting to help configure SpecCleanse"
    )
    parser.add_argument("input", type=Path, help="Input DOCX file")
    parser.add_argument(
        "-s", "--search", 
        help="Search for paragraphs containing this text"
    )
    parser.add_argument(
        "-a", "--all",
        action="store_true",
        help="Show all paragraphs (verbose)"
    )
    parser.add_argument(
        "-e", "--editorial",
        action="store_true", 
        help="Find likely editorial content and show its formatting"
    )
    
    args = parser.parse_args()
    
    if not args.input.exists():
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)
    
    if args.editorial:
        find_editorial_content(args.input)
    else:
        analyze_document(args.input, args.search, args.all)


if __name__ == "__main__":
    main()
