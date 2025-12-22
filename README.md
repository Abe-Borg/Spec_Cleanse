# SpecCleanse

A Python CLI tool that removes unnecessary content from specification Word documents (.docx), leaving only the actual specification content while preserving all formatting and styles.

## Features

### Shallow Clean (Default)
Content-level cleaning that removes editorial/specifier content:

- **Specifier notes** - Editorial comments like `[Specifier: ...]`, `NOTE TO SPECIFIER`, etc.
- **Copyright notices** - Boilerplate copyright and licensing text
- **Hidden text** - Content marked with Word's vanish property
- **SpecAgent references** – Visible watermarks, footers, and text-level URLs from specagent.com  
  (Note: embedded hyperlinks and metadata are handled during deep clean)

- **Editorial artifacts** - Placeholders like `<insert>`, `RETAIN OR DELETE`, `[TBD]`
- **Unused styles** – Styles defined but never referenced in document content  
  (content-aware removal; deeper orphan analysis is available with `--deep`)


### Deep Clean (`--deep`)
ZIP/XML-level cleaning that removes accumulated cruft:

| Category | What Gets Removed | Typical Savings |
|----------|-------------------|-----------------|
| **Orphaned Media** | Images in `word/media/` that nothing references | Variable (KB-MB) |
| **Orphaned Styles** | Style definitions never used in document content | ~500 bytes each |
| **RSID Tracking** | RSID attributes and RSID registries in settings.xml | ~25 bytes each, often 1000s per doc |
| **Empty Elements** | Empty runs `<w:r/>`, empty properties `<w:rPr/>` | ~20 bytes each |
| **Non-English Fonts** | Font mappings for Japanese, Arabic, Hebrew, etc. in theme | ~60 bytes each |
| **Compat Settings** | Word compatibility behaviors and legacy layout rules | ~50 bytes each |
| **Internal Bookmarks** | `_GoBack`, `_Hlk*`, `_Ref*` bookmarks | ~80 bytes each |
| **Proof State** | Spell/grammar check state markers | ~40 bytes each |
| **External Links (Domains)** | External hyperlink relationships and cached link metadata for specified domains (e.g. specagent.com) | Variable (often KBs) |


## Installation

### Requirements

- Python 3.10+
- Windows (primary), Linux/macOS (compatible)

### Dependencies

```bash
pip install lxml pyyaml
```

### Setup

1. Clone or download the `speccleanse` directory
2. Install dependencies: `pip install lxml pyyaml`
3. Run: `python speccleanse.py input.docx output.docx`

## Usage

### Basic Usage

```bash
# Shallow clean only (remove specifier notes, copyright, etc.)
python speccleanse.py input.docx cleaned.docx

# Shallow + deep clean (full optimization)
python speccleanse.py input.docx cleaned.docx --deep

# Preview what would be removed (dry run)
python speccleanse.py input.docx cleaned.docx --deep --dry-run

# Deep clean only (skip content removal)
python speccleanse.py input.docx cleaned.docx --deep-only

# Verbose output with detailed detections
python speccleanse.py input.docx cleaned.docx --deep -v
```

### Command Line Options

| Option | Description |
|--------|-------------|
| `input` | Input DOCX file to process |
| `output` | Output DOCX file path |
| `-c, --config` | Path to patterns YAML file (default: patterns.yaml) |
| `-d, --dry-run` | Detect without modifying (preview mode) |
| `-v, --verbose` | Show detailed detection information |
| `-q, --quiet` | Suppress output except errors |
| `--clean-styles` | Also remove unused styles from document |
| `--styles-only` | Only clean unused styles, skip content removal |
| `--deep` | Enable deep cleaning (orphans + cruft) |
| `--deep-only` | Only perform deep cleaning, skip shallow clean |
| `--version` | Show version number |

### Deep Clean Options

Use these with `--deep` to selectively disable specific cleaning operations:

| Option | Description |
|--------|-------------|
| `--no-media` | Keep orphaned media files |
| `--no-deep-styles` | Keep orphaned style definitions |
| `--no-rsids` | Keep RSID tracking attributes |
| `--no-empty` | Keep empty runs/elements |
| `--no-fonts` | Keep non-English font mappings |
| `--no-compat` | Keep backwards compatibility settings |
| `--no-bookmarks` | Keep internal Word bookmarks |
| `--no-proof` | Keep spell/grammar check state |
| `--strip-links-domain DOMAIN` | Remove external hyperlinks and metadata for a given domain (repeatable) |
| `--no-links` | Disable external link domain scrubbing |
| `--aggressive-compat` | Remove entire `<w:compat>` block (higher risk, opt-in) |


> ⚠ Aggressive Compatibility Removal  
> `--aggressive-compat` removes the entire Word compatibility block.
> This may affect layout consistency across Word versions.
> Recommended only for finalized specifications.



### Examples

```bash
# Basic processing
python speccleanse.py "Division 23 - HVAC.docx" "Division 23 - HVAC_cleaned.docx"

# Preview shallow + deep clean
python speccleanse.py spec.docx out.docx --deep --dry-run -v

# Full clean with style removal
python speccleanse.py spec.docx out.docx --deep --clean-styles

# Deep clean: only RSIDs (safest, biggest savings)
python speccleanse.py spec.docx out.docx --deep \
   --no-media --no-deep-styles \
    --no-empty --no-fonts --no-compat --no-bookmarks --no-proof

# Deep clean: everything except non-English fonts (for multilingual docs)
python speccleanse.py spec.docx out.docx --deep --no-fonts

# Quiet mode for scripting
python speccleanse.py spec.docx out.docx --deep -q && echo "Success"
```

# Remove all SpecAgent hyperlinks and metadata
python speccleanse.py spec.docx out.docx --deep --strip-links-domain specagent.com

# Remove multiple external domains
python speccleanse.py spec.docx out.docx --deep \
  --strip-links-domain specagent.com \
  --strip-links-domain example.com




## How It Works

### Shallow Clean Pipeline

1. **Unpack** - DOCX files are ZIP archives; SpecCleanse extracts them
2. **Parse** - XML content is parsed with lxml, preserving structure
3. **Detect** - Each paragraph and run is analyzed against patterns
4. **Decide** - Confidence scores determine if content is removed
5. **Remove** - Elements are surgically removed, preserving surrounding content
6. **Repack** - Modified XML is repacked into a new DOCX

### Deep Clean Pipeline

1. **Analyze** - Scan all XML files to find defined vs. used resources
2. **Compute Orphans** - Resources defined but never referenced
3. **Scan Cruft** - RSIDs, empty elements, compatibility settings, etc.
4. **Remove** - Delete orphaned files, strip attributes, clean XML
5. **Validate** – Ensure required relationships and core document structure remain intact
6. **Repack** - Reconstruct DOCX with cleaned content

### Confidence Scoring

Each detection has a confidence score (0.0 - 1.0):
- Pattern match alone: ~0.6
- Pattern + italic: ~0.8
- Pattern + color: ~0.9
- Style name match: ~0.8+

Content is removed if confidence ≥ 0.5 and no preserve pattern matches.

## Detection Patterns

SpecCleanse uses configurable patterns defined in `patterns.yaml`. You can customize detection rules for your specific needs.

### Content Types

1. **Specifier Notes** (`specifier_notes`)
   - Pattern-based: `[Specifier: ...]`, `NOTE TO SPECIFIER`, etc.
   - Style-based: Paragraphs/runs with specific style names
   - Format-based: Italic text with red/blue color

2. **Copyright Notices** (`copyright_notices`)
   - `©`, `Copyright 2024`, `All rights reserved`
   - Licensing and reproduction restrictions

3. **Hidden Text** (`hidden_text`)
   - Any content with Word's `<w:vanish/>` property

4. **SpecAgent References** (`specagent_references`)
   - URLs, watermarks, attribution text from specagent.com

5. **Editorial Artifacts** (`editorial_artifacts`)
   - `RETAIN OR DELETE`, `<INSERT>`, `[TBD]`
   - Empty placeholders: `[ ]`, `[___]`

6. **Preserve Patterns** (`preserve_patterns`)
   - Content that should NEVER be removed
   - `END OF SECTION`, `PART 1`, `SECTION 1`

### Customizing Patterns

Edit `patterns.yaml` to add or modify detection rules:

```yaml
specifier_notes:
  enabled: true
  text_patterns:
    - '\[specifier[:\s].*?\]'    # [Specifier: ...]
    - 'note to specifier'         # NOTE TO SPECIFIER
    - '\[my custom pattern\]'     # Add your own!
  
  formatting_signals:
    italic: true
    colors:
      - "FF0000"  # Red (hex without #)
      - "0000FF"  # Blue

# Add patterns to never remove
preserve_patterns:
  enabled: true
  text_patterns:
    - 'end\s+of\s+section'
    - 'part\s+\d+'
    - 'my\s+important\s+pattern'  # Add your own!
```

## Understanding DOCX Internals

A DOCX file is a ZIP archive containing:

```
docx_file.docx
├── [Content_Types].xml      # Maps file types to MIME types
├── _rels/
│   └── .rels               # Root relationships (points to main doc)
├── docProps/
│   ├── app.xml             # Application properties
│   └── core.xml            # Core metadata (author, dates)
└── word/
    ├── _rels/
    │   └── document.xml.rels  # Document relationships (images, hyperlinks)
    ├── document.xml        # THE ACTUAL CONTENT
    ├── styles.xml          # Style definitions
    ├── settings.xml        # Document settings
    ├── fontTable.xml       # Font declarations
    ├── numbering.xml       # List/numbering definitions
    ├── webSettings.xml     # Web-related settings (paste artifacts)
    ├── theme/
    │   └── theme1.xml      # Theme colors/fonts
    └── media/              # Embedded images
        ├── image1.png
        └── image2.jpg
```

### Relationship IDs (rIds)

Content in `document.xml` references external resources via `rId` attributes:

```xml
<!-- Hyperlink using rId8 -->
<w:hyperlink r:id="rId8">
  <w:r><w:t>Click here</w:t></w:r>
</w:hyperlink>

<!-- Image using rId5 -->
<a:blip r:embed="rId5"/>
```


## File Structure

```
speccleanse/
├── speccleanse.py     # CLI entry point
├── detection.py       # Detection engine and detectors
├── processor.py       # DOCX content processing logic
├── style_cleaner.py   # Unused style detection and removal
├── deep_cleaner.py    # Orphan analysis and deep cleaning
├── patterns.yaml      # Configurable detection patterns
├── diagnose.py        # Diagnostic tool for inspecting documents
└── README.md          # This file
```

## Diagnostic Tool

Use `diagnose.py` to inspect document formatting and help configure patterns:

```bash
# Find paragraphs containing specific text
python diagnose.py input.docx -s "specifier"

# Show all paragraphs with formatting details
python diagnose.py input.docx -a

# Find likely editorial content and show its formatting
python diagnose.py input.docx -e
```


## Sanitization Use Cases

SpecCleanse can also be used to sanitize DOCX files before issuing or archiving:

- Remove embedded tracking identifiers (RSIDs)
- Strip external hyperlink domains (e.g., vendor tracking URLs)
- Remove authoring and compatibility metadata
- Reduce document fingerprinting

This is especially useful for issued-for-bid or issued-for-construction specifications.



## Common Specification Cruft

Master spec templates (like MasterSpec, BSD SpecLink, ARCOM) often accumulate:

1. **SpecAgent hyperlinks** - Product lookup URLs that should be removed for final specs
2. **Unused styles** - Template styles for sections you deleted
3. **Paste artifacts** - `<w:div>` elements in webSettings.xml from copy/paste
4. **Revision history** - Tracked changes that were accepted but leave cruft
5. **Dead media** - Images from deleted sections

## Troubleshooting

### Common Issues

**"Invalid DOCX file"**
- Ensure file is `.docx` not `.doc`
- File may be corrupted; try opening/saving in Word

**Content not detected**
- Check patterns in `patterns.yaml`
- Use `--dry-run -v` to see what's being detected
- Add custom patterns for your content

**Too much removed**
- Add patterns to `preserve_patterns` section
- Adjust confidence thresholds in detection.py
- Use `--dry-run` to preview before processing

**Document won't open after cleaning**
- Check the error messages in the output
- Try selective deep cleaning (e.g., `--no-deep-styles`) to isolate the problem
- Deep clean validates structure before completing

**Still seeing SpecAgent references after deep clean**
- Visible text or footers: ensure shallow clean ran (default behavior)
- Clickable links or metadata: ensure `--strip-links-domain specagent.com` is enabled


**File size didn't change much with deep clean**
- Orphan removal typically saves a few KB
- The big wins come from RSID stripping (often 25-125 KB)
- Running after shallow clean creates more orphans to remove

## Limitations

- Only processes `.docx` format (not `.doc`)
- Complex nested tables may have edge cases
- Embedded objects (OLE) are not scanned for text
- Very large documents may be slow (processes all XML)

## Contributing

To add new detection types:

1. Add patterns to `patterns.yaml`
2. Create detector class in `detection.py` (extend `BaseDetector`)
3. Register in `DetectionEngine._create_detectors()`

## License

MIT License - Use freely in your MEP specification workflow.