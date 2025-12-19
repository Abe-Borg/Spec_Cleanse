# SpecCleanse

A Python CLI tool that removes unnecessary content from specification Word documents (.docx), leaving only the actual specification content while preserving all formatting and styles.

## Features

- **Removes specifier notes** - Editorial comments like `[Specifier: ...]`, `NOTE TO SPECIFIER`, etc.
- **Removes copyright notices** - Boilerplate copyright and licensing text
- **Removes hidden text** - Content marked with Word's vanish property
- **Removes SpecAgent references** - Watermarks, footers, URLs from specagent.com
- **Removes editorial artifacts** - Placeholders like `<insert>`, `RETAIN OR DELETE`, `[TBD]`
- **Removes unused styles** - Styles defined but never referenced in document content
- **Preserves legitimate content** - "END OF SECTION", actual spec text, all formatting

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
# Process a specification document
python speccleanse.py input.docx cleaned.docx

# Preview what would be removed (dry run)
python speccleanse.py input.docx cleaned.docx --dry-run

# Verbose output with detailed detections
python speccleanse.py input.docx cleaned.docx -v

# Use custom patterns configuration
python speccleanse.py input.docx cleaned.docx --config my_patterns.yaml
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
| `--version` | Show version number |

### Examples

```bash
# Basic processing
python speccleanse.py "Division 23 - HVAC.docx" "Division 23 - HVAC_cleaned.docx"

# Preview mode to see what would be removed
python speccleanse.py spec.docx out.docx --dry-run -v

# Clean content AND unused styles
python speccleanse.py spec.docx out.docx --clean-styles

# Only remove unused styles (no content changes)
python speccleanse.py spec.docx out.docx --styles-only

# Quiet mode for scripting
python speccleanse.py spec.docx out.docx -q && echo "Success"
```

## Style Cleaning

SpecCleanse can detect and remove unused styles from your specification documents. This is useful for cleaning up documents that have accumulated many unused styles over time.

### How It Works

1. Parses `word/styles.xml` to find all defined styles
2. Scans document content (including headers, footers, footnotes) for style references
3. Builds a dependency graph (basedOn, link, next relationships)
4. Identifies styles that are not used and not required by used styles
5. Removes unused styles while preserving protected/built-in styles

### Protected Styles

Certain built-in styles are never removed even if unused:
- Normal, DefaultParagraphFont, TableNormal, NoList
- Heading1-9, Title, Subtitle
- Header, Footer, TOC styles
- And other Word essential styles

### Usage

```bash
# Analyze styles without modifying
python speccleanse.py doc.docx out.docx --styles-only --dry-run -v

# Remove unused styles only
python speccleanse.py doc.docx out.docx --styles-only

# Remove content AND unused styles
python speccleanse.py doc.docx out.docx --clean-styles
```

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

### Pattern Syntax

- Patterns use Python regex syntax (case-insensitive)
- Use `.*?` for non-greedy matching
- Use `\s` for whitespace
- Colors are Word's internal hex codes (without #)

## How It Works

1. **Unpack** - DOCX files are ZIP archives; SpecCleanse extracts them
2. **Parse** - XML content is parsed with lxml, preserving structure
3. **Detect** - Each paragraph and run is analyzed against patterns
4. **Decide** - Confidence scores determine if content is removed
5. **Remove** - Elements are surgically removed, preserving surrounding content
6. **Repack** - Modified XML is repacked into a new DOCX

### Confidence Scoring

Each detection has a confidence score (0.0 - 1.0):
- Pattern match alone: ~0.6
- Pattern + italic: ~0.8
- Pattern + color: ~0.9
- Style name match: ~0.8+

Content is removed if confidence ≥ 0.5 and no preserve pattern matches.

## File Structure

```
speccleanse/
├── speccleanse.py     # CLI entry point
├── detection.py       # Detection engine and detectors
├── processor.py       # DOCX content processing logic
├── style_cleaner.py   # Unused style detection and removal
├── patterns.yaml      # Configurable detection patterns
└── README.md          # This file
```

## Limitations

- Only processes `.docx` format (not `.doc`)
- Complex nested tables may have edge cases
- Embedded objects (OLE) are not scanned for text
- Very large documents may be slow (processes all XML)

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

## Contributing

To add new detection types:

1. Add patterns to `patterns.yaml`
2. Create detector class in `detection.py` (extend `BaseDetector`)
3. Register in `DetectionEngine._create_detectors()`

## License

MIT License - Use freely in your MEP specification workflow.
