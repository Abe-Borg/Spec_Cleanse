#!/usr/bin/env python3
"""
SpecCleanse - Specification Document Content Stripper

CLI tool that removes unnecessary content from specification Word documents,
leaving only the actual specification content while preserving all formatting.

Features:
- Shallow clean (default): Remove specifier notes, copyright, hidden text, etc.
- Deep clean (--deep): Remove orphaned resources and cruft at ZIP/XML level
"""

import argparse
import os
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import yaml

from detection import DetectionEngine, ContentType
from processor import DocxProcessor, ProcessingResult
from style_cleaner import StyleCleaner, StyleCleanResult


def load_config(config_path: Path) -> dict:
    """Load configuration from YAML file."""
    if not config_path.exists():
        print(f"Error: Config file not found: {config_path}", file=sys.stderr)
        sys.exit(1)
    
    with open(config_path, 'r') as f:
        return yaml.safe_load(f)


def print_result(result: ProcessingResult, verbose: bool = False, 
                 style_result: StyleCleanResult = None, deep_result=None):
    """Print processing result summary."""
    print()
    print("=" * 60)
    print("SpecCleanse Processing Report")
    print("=" * 60)
    print()
    print(f"Input:  {result.input_path}")
    print(f"Output: {result.output_path}")
    print()
    
    if result.errors:
        print("ERRORS:")
        for error in result.errors:
            print(f"  ✗ {error}")
        print()
        return
    
    # Separate preserved vs removed detections
    removed_detections = [d for d in result.detections if d.content_type != ContentType.PRESERVE]
    preserved_detections = [d for d in result.detections if d.content_type == ContentType.PRESERVE]
    
    # Summary by type
    type_counts = {}
    for d in result.detections:
        type_name = d.content_type.value
        type_counts[type_name] = type_counts.get(type_name, 0) + 1
    
    if type_counts:
        print("Content Detections by Type:")
        for type_name, count in sorted(type_counts.items()):
            status = "preserved" if type_name == "preserve" else "removed"
            print(f"  • {type_name}: {count} ({status})")
        print()
    
    print(f"Total Content Removed:   {len(removed_detections)}")
    print(f"Total Content Preserved: {len(preserved_detections)}")
    print()
    
    # Always show removed content details (not just in verbose mode)
    if removed_detections:
        print("-" * 60)
        print("REMOVED CONTENT:")
        print("-" * 60)
        for i, d in enumerate(removed_detections, 1):
            preview = d.text[:80] + "..." if len(d.text) > 80 else d.text
            preview = preview.replace('\n', '↵').replace('\r', '')
            print(f"\n{i}. [{d.content_type.value}]")
            print(f"   \"{preview}\"")
            if verbose:
                print(f"   Confidence: {d.confidence:.0%}")
                print(f"   Reason: {d.reason}")
    
    # Show preserved content in verbose mode
    if verbose and preserved_detections:
        print()
        print("-" * 60)
        print("PRESERVED CONTENT (matched whitelist):")
        print("-" * 60)
        for i, d in enumerate(preserved_detections, 1):
            preview = d.text[:80] + "..." if len(d.text) > 80 else d.text
            preview = preview.replace('\n', '↵').replace('\r', '')
            print(f"  {i}. \"{preview}\"")
    
    # Style cleaning results
    if style_result:
        print()
        print("-" * 60)
        print("Style Cleaning:")
        print(f"  Total styles defined:  {style_result.total_styles}")
        print(f"  Styles in use:         {len(style_result.used_styles)}")
        print(f"  Unused (removable):    {len(style_result.unused_styles)}")
        print(f"  Unused (protected):    {len(style_result.protected_styles)}")
        print(f"  Styles removed:        {len(style_result.removed_styles)}")
        
        if style_result.removed_styles:
            print()
            print("  Removed styles:")
            for style_id in sorted(style_result.removed_styles):
                print(f"    - {style_id}")
        
        if verbose and style_result.unused_styles:
            print()
            print("  Removable (not yet removed in dry-run):")
            for style_id in sorted(style_result.unused_styles):
                print(f"    - {style_id}")
        print()
    
    # Deep cleaning results
    if deep_result:
        print()
        print("-" * 60)
        print("Deep Cleaning:")
        print("-" * 60)
        print()
        print("  Orphans Removed:")
        print(f"    Media files:         {deep_result.media_removed}")
        print(f"    Styles:              {deep_result.styles_removed}")
        print()
        print("  Cruft Removed:")
        print(f"    RSID attributes:     {deep_result.rsids_removed}")
        print(f"    Empty elements:      {deep_result.empty_elements_removed}")
        print(f"    Non-English fonts:   {deep_result.font_mappings_removed}")
        print(f"    Compat settings:     {deep_result.compat_settings_removed}")
        print(f"    Internal bookmarks:  {deep_result.bookmarks_removed}")
        print(f"    Proof state:         {deep_result.proof_elements_removed}")
        print()
        print(f"  Estimated bytes saved: {deep_result.bytes_saved:,} ({deep_result.bytes_saved/1024:.1f} KB)")
        
        if deep_result.warnings:
            print()
            print("  Warnings:")
            for w in deep_result.warnings:
                print(f"    - {w}")
        
        if deep_result.errors:
            print()
            print("  Errors:")
            for e in deep_result.errors:
                print(f"    - {e}")
        print()
    
    if verbose and result.detections:
        print("-" * 60)
        print("Detailed Detections:")
        print("-" * 60)
        for i, d in enumerate(result.detections, 1):
            status = "PRESERVED" if d.content_type == ContentType.PRESERVE else "REMOVED"
            preview = d.text[:60] + "..." if len(d.text) > 60 else d.text
            preview = preview.replace('\n', '↵')
            print(f"\n{i}. [{status}] {d.content_type.value}")
            print(f"   Text: \"{preview}\"")
            print(f"   Confidence: {d.confidence:.0%}")
            print(f"   Reason: {d.reason}")
    
    print()
    if result.success:
        print("✓ Processing completed successfully!")
    else:
        print("✗ Processing completed with errors.")


def run_deep_clean(unpacked_dir: Path, args, verbose: bool = False):
    """Run deep cleaning on an unpacked DOCX directory."""
    # Import here to avoid circular imports and allow standalone use
    from deep_cleaner import analyze_and_clean, get_analysis_only
    
    if args.dry_run:
        # Just analyze, don't clean
        report = get_analysis_only(unpacked_dir, verbose=verbose)
        
        # Create a pseudo-result for reporting
        from deep_cleaner import DeepCleanResult
        result = DeepCleanResult(success=True)
        result.media_removed = len(report.orphaned_media)
        result.styles_removed = len(report.orphaned_styles)
        result.rsids_removed = report.total_rsid_attributes
        result.empty_elements_removed = report.total_empty_elements
        result.font_mappings_removed = len(report.non_english_font_mappings)
        result.compat_settings_removed = len(report.compatibility_settings)
        result.bookmarks_removed = len(report.internal_bookmarks)
        result.proof_elements_removed = len(report.proof_state_elements)
        result.bytes_saved = report.estimated_savings_bytes
        return result
    
    # If --only is specified, disable everything except that one operation
    only = getattr(args, 'only', None)
    if only:
        # Start with everything disabled
        flags = {
            'remove_media': False,
            'remove_styles': False,
            'strip_rsids': False,
            'remove_empty_elements': False,
            'remove_non_english_fonts': False,
            'remove_compat_settings': False,
            'remove_internal_bookmarks': False,
            'remove_proof_state': False,
        }
        # Enable only the specified operation
        only_map = {
            'media': 'remove_media',
            'styles': 'remove_styles',
            'rsids': 'strip_rsids',
            'empty': 'remove_empty_elements',
            'fonts': 'remove_non_english_fonts',
            'compat': 'remove_compat_settings',
            'bookmarks': 'remove_internal_bookmarks',
            'proof': 'remove_proof_state',
        }
        if only in only_map:
            flags[only_map[only]] = True
            if verbose:
                print(f"  [DEBUG] Running ONLY: {only}")
        else:
            print(f"Warning: Unknown --only value '{only}', running all operations")
            flags = {k: True for k in flags}
    else:
        # Normal mode: use --no-* flags
        flags = {
            'remove_media': not args.no_media,
            'remove_styles': not args.no_deep_styles,
            'strip_rsids': not args.no_rsids,
            'remove_empty_elements': not args.no_empty,
            'remove_non_english_fonts': not args.no_fonts,
            'remove_compat_settings': not args.no_compat,
            'remove_internal_bookmarks': not args.no_bookmarks,
            'remove_proof_state': not args.no_proof,
        }
    
    return analyze_and_clean(
        unpacked_dir,
        **flags,
        verbose=verbose,
    )


def main():
    parser = argparse.ArgumentParser(
        prog="speccleanse",
        description="Remove unnecessary content from specification Word documents.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic shallow clean (content removal)
  speccleanse input.docx output.docx

  # Shallow clean + deep clean (full optimization)
  speccleanse input.docx output.docx --deep

  # Preview what would be removed
  speccleanse input.docx output.docx --deep --dry-run

  # Deep clean only (skip content removal)
  speccleanse input.docx output.docx --deep-only

  # Selective deep clean (RSIDs only)
  speccleanse input.docx output.docx --deep --no-media \\
      --no-deep-styles --no-empty --no-fonts --no-compat --no-bookmarks --no-proof

Shallow Clean (default):
  • Specifier notes (editorial comments for specifiers)
  • Copyright notices (boilerplate copyright text)
  • Hidden text (Word's vanish property)
  • SpecAgent references (watermarks, URLs, attribution)
  • Editorial artifacts (placeholders, instructions)

Deep Clean (--deep):
  • Orphaned media files (images from deleted sections)
  • Orphaned styles (unused style definitions)
  • RSID attributes (revision tracking IDs - biggest savings)
  • Empty elements (empty runs, properties)
  • Non-English fonts (font mappings for unused scripts)
  • Compatibility settings (Word 97/2002/2003 cruft)
  • Internal bookmarks (_GoBack, _Hlk*, _Ref*)
  • Proof state (spell/grammar check markers)
        """
    )
    
    parser.add_argument(
        "input",
        type=Path,
        help="Input DOCX file to process"
    )
    
    parser.add_argument(
        "output",
        type=Path,
        nargs="?",
        default=None,
        help="Output DOCX file path (not required for --dry-run)"
    )
    
    parser.add_argument(
        "-c", "--config",
        type=Path,
        default=Path(__file__).parent / "patterns.yaml",
        help="Path to patterns configuration file (default: patterns.yaml)"
    )
    
    parser.add_argument(
        "-d", "--dry-run",
        action="store_true",
        help="Detect content without modifying (preview mode)"
    )
    
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Show detailed detection information"
    )
    
    parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="Suppress all output except errors"
    )
    
    # Shallow clean options
    parser.add_argument(
        "--clean-styles",
        action="store_true",
        help="Also remove unused styles from the document"
    )
    
    parser.add_argument(
        "--styles-only",
        action="store_true", 
        help="Only clean unused styles, skip content removal"
    )
    
    # Deep clean options
    parser.add_argument(
        "--deep",
        action="store_true",
        help="Enable deep cleaning (remove orphans and cruft at ZIP/XML level)"
    )
    
    parser.add_argument(
        "--deep-only",
        action="store_true",
        help="Only perform deep cleaning, skip shallow content removal"
    )
    
    # Granular deep clean flags (all default to False = enabled)
    deep_group = parser.add_argument_group("Deep clean options (use with --deep)")


    deep_group.add_argument(
        "--only",
        choices=['media', 'styles', 'rsids', 'empty', 'fonts', 'compat', 'bookmarks', 'proof'],
        help="Run ONLY this single deep clean operation (for debugging)"
    )
    
    deep_group.add_argument(
        "--no-media",
        action="store_true",
        help="Skip removing orphaned media files"
    )
    
    deep_group.add_argument(
        "--no-deep-styles",
        action="store_true",
        help="Skip removing orphaned styles (deep clean)"
    )
    
    deep_group.add_argument(
        "--no-rsids",
        action="store_true",
        help="Skip stripping RSID tracking attributes"
    )
    
    deep_group.add_argument(
        "--no-empty",
        action="store_true",
        help="Skip removing empty elements"
    )
    
    deep_group.add_argument(
        "--no-fonts",
        action="store_true",
        help="Skip removing non-English font mappings"
    )
    
    deep_group.add_argument(
        "--no-compat",
        action="store_true",
        help="Skip removing compatibility settings"
    )
    
    deep_group.add_argument(
        "--no-bookmarks",
        action="store_true",
        help="Skip removing internal bookmarks"
    )
    
    deep_group.add_argument(
        "--no-proof",
        action="store_true",
        help="Skip removing proof state elements"
    )
    
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s 2.0.0"
    )
    
    args = parser.parse_args()
    
    # Validate input
    if not args.input.exists():
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)
    
    if args.input.suffix.lower() != ".docx":
        print(f"Warning: Input file does not have .docx extension: {args.input}", 
              file=sys.stderr)
    
    # Validate output argument
    if not args.dry_run and args.output is None:
        print("Error: Output file required (unless using --dry-run)", file=sys.stderr)
        sys.exit(1)
    
    # For dry-run without output, use a placeholder path for reporting
    if args.output is None:
        args.output = Path("(dry-run)")
    elif args.output.suffix.lower() != ".docx":
        args.output = args.output.with_suffix(".docx")
    
    # Load config
    config = load_config(args.config)
    
    # Create engine and processor
    engine = DetectionEngine(config)
    processor = DocxProcessor(engine, verbose=args.verbose)
    style_cleaner = StyleCleaner(verbose=args.verbose)
    
    # Determine mode
    do_shallow = not args.styles_only and not args.deep_only
    do_styles = args.clean_styles or args.styles_only
    do_deep = args.deep or args.deep_only
    
    if not args.quiet:
        mode_parts = []
        if do_shallow:
            mode_parts.append("SHALLOW CLEAN")
        if do_styles:
            mode_parts.append("STYLE CLEAN")
        if do_deep:
            mode_parts.append("DEEP CLEAN")
        mode = " + ".join(mode_parts) if mode_parts else "PROCESSING"
        if args.dry_run:
            mode = f"DRY RUN: {mode}"
        print(f"\n{mode}: {args.input}")
    
    # Process document
    style_result = None
    deep_result = None
    
    # Create temp directory for all operations
    temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_"))
    
    try:
        if args.styles_only or args.deep_only:
            # Skip shallow clean, create minimal result
            result = ProcessingResult(input_path=args.input, output_path=args.output)
            
            # Unpack for styles/deep clean
            unpacked = temp_dir / "unpacked"
            with zipfile.ZipFile(args.input, 'r') as zf:
                zf.extractall(unpacked)
            
            if args.styles_only:
                style_result = style_cleaner.clean(unpacked, dry_run=args.dry_run)
            
            if args.deep_only:
                deep_result = run_deep_clean(unpacked, args, verbose=args.verbose)
            
            # Repack if not dry run
            if not args.dry_run:
                args.output.parent.mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(args.output, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for root, dirs, files in os.walk(unpacked):
                        for file in files:
                            file_path = Path(root) / file
                            arcname = file_path.relative_to(unpacked)
                            zf.write(file_path, arcname)
        else:
            # Normal shallow content processing
            result = processor.process(
                input_path=args.input,
                output_path=args.output,
                dry_run=args.dry_run
            )
            
            # Style cleaning if requested
            if do_styles and result.success:
                if args.dry_run:
                    unpacked = temp_dir / "unpacked"
                    with zipfile.ZipFile(args.input, 'r') as zf:
                        zf.extractall(unpacked)
                    style_result = style_cleaner.analyze(unpacked)
                else:
                    unpacked = temp_dir / "unpacked"
                    with zipfile.ZipFile(args.output, 'r') as zf:
                        zf.extractall(unpacked)
                    
                    style_result = style_cleaner.clean(unpacked, dry_run=False)
                    
                    # Repack
                    with zipfile.ZipFile(args.output, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for root, dirs, files in os.walk(unpacked):
                            for file in files:
                                file_path = Path(root) / file
                                arcname = file_path.relative_to(unpacked)
                                zf.write(file_path, arcname)
            
            # Deep cleaning if requested
            if do_deep and result.success:
                if args.dry_run:
                    # Analyze the input file
                    unpacked = temp_dir / "unpacked_deep"
                    with zipfile.ZipFile(args.input, 'r') as zf:
                        zf.extractall(unpacked)
                    deep_result = run_deep_clean(unpacked, args, verbose=args.verbose)
                else:
                    # Clean the output file
                    unpacked = temp_dir / "unpacked_deep"
                    with zipfile.ZipFile(args.output, 'r') as zf:
                        zf.extractall(unpacked)
                    
                    deep_result = run_deep_clean(unpacked, args, verbose=args.verbose)
                    
                    # Repack
                    with zipfile.ZipFile(args.output, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for root, dirs, files in os.walk(unpacked):
                            for file in files:
                                file_path = Path(root) / file
                                arcname = file_path.relative_to(unpacked)
                                zf.write(file_path, arcname)
    
    finally:
        # Cleanup temp directory
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
    
    # Output results
    if not args.quiet:
        print_result(result, verbose=args.verbose, style_result=style_result, deep_result=deep_result)
    
    # Exit code
    sys.exit(0 if result.success else 1)


if __name__ == "__main__":
    main()