#!/usr/bin/env python3
"""
SpecCleanse - Specification Document Content Stripper

CLI tool that removes unnecessary content from specification Word documents,
leaving only the actual specification content while preserving all formatting.
"""

import argparse
import sys
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


def print_result(result: ProcessingResult, verbose: bool = False, style_result: StyleCleanResult = None):
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
            print(f"  ❌ {error}")
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
        print("✅ Processing completed successfully!")
    else:
        print("❌ Processing completed with errors.")


def main():
    parser = argparse.ArgumentParser(
        prog="speccleanse",
        description="Remove unnecessary content from specification Word documents.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  speccleanse input.docx output.docx
  speccleanse input.docx output.docx --dry-run
  speccleanse input.docx output.docx --config custom_patterns.yaml -v

Content Types Removed:
  • Specifier notes (editorial comments for specifiers)
  • Copyright notices (boilerplate copyright text)
  • Hidden text (Word's vanish property)
  • SpecAgent references (watermarks, URLs, attribution)
  • Editorial artifacts (placeholders, instructions)

Content Preserved:
  • All actual specification content
  • "END OF SECTION" text
  • All formatting and Styles
  • Document structure
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
    
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s 1.0.0"
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
    # Ensure output has .docx extension
    elif args.output.suffix.lower() != ".docx":
        args.output = args.output.with_suffix(".docx")
    
    # Load config
    config = load_config(args.config)
    
    # Create engine and processor
    engine = DetectionEngine(config)
    processor = DocxProcessor(engine, verbose=args.verbose)
    style_cleaner = StyleCleaner(verbose=args.verbose)
    
    if not args.quiet:
        mode = "DRY RUN" if args.dry_run else "PROCESSING"
        if args.styles_only:
            mode += " (styles only)"
        elif args.clean_styles:
            mode += " (with style cleaning)"
        print(f"\n{mode}: {args.input}")
    
    # Process document
    style_result = None
    
    if args.styles_only:
        # Only clean styles, create a minimal result
        result = ProcessingResult(input_path=args.input, output_path=args.output)
        
        # Need to unpack, clean, repack manually
        import tempfile
        import shutil
        import zipfile
        
        temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_"))
        try:
            # Unpack
            unpacked = temp_dir / "unpacked"
            with zipfile.ZipFile(args.input, 'r') as zf:
                zf.extractall(unpacked)
            
            # Clean styles
            style_result = style_cleaner.clean(unpacked, dry_run=args.dry_run)
            
            # Repack if not dry run
            if not args.dry_run:
                import os
                args.output.parent.mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(args.output, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for root, dirs, files in os.walk(unpacked):
                        for file in files:
                            file_path = Path(root) / file
                            arcname = file_path.relative_to(unpacked)
                            zf.write(file_path, arcname)
        finally:
            shutil.rmtree(temp_dir)
    else:
        # Normal content processing
        result = processor.process(
            input_path=args.input,
            output_path=args.output,
            dry_run=args.dry_run
        )
        
        # Also clean styles if requested
        if args.clean_styles and result.success and not args.dry_run:
            # Need to unpack the output, clean, repack
            import tempfile
            import shutil
            import zipfile
            
            temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_"))
            try:
                unpacked = temp_dir / "unpacked"
                with zipfile.ZipFile(args.output, 'r') as zf:
                    zf.extractall(unpacked)
                
                style_result = style_cleaner.clean(unpacked, dry_run=False)
                
                # Repack
                import os
                with zipfile.ZipFile(args.output, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for root, dirs, files in os.walk(unpacked):
                        for file in files:
                            file_path = Path(root) / file
                            arcname = file_path.relative_to(unpacked)
                            zf.write(file_path, arcname)
            finally:
                shutil.rmtree(temp_dir)
        
        elif args.clean_styles and args.dry_run:
            # Dry run - just analyze styles
            import tempfile
            import shutil
            import zipfile
            
            temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_"))
            try:
                unpacked = temp_dir / "unpacked"
                with zipfile.ZipFile(args.input, 'r') as zf:
                    zf.extractall(unpacked)
                style_result = style_cleaner.analyze(unpacked)
            finally:
                shutil.rmtree(temp_dir)
    
    # Output results
    if not args.quiet:
        print_result(result, verbose=args.verbose, style_result=style_result)
    
    # Exit code
    sys.exit(0 if result.success else 1)


if __name__ == "__main__":
    main()