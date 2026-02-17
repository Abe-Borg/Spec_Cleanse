# SpecCleanse

A Python GUI tool that removes unnecessary content from specification Word documents (.docx), leaving only the actual specification content while preserving all formatting and styles.

## Features

SpecCleanse performs a maximum-strength clean in three stages:

### 1. Shallow Clean — Content Removal
- **Specifier notes** - Editorial comments like `[Specifier: ...]`, `NOTE TO SPECIFIER`, etc.
- **Copyright notices** - Boilerplate copyright and licensing text
- **Hidden text** - Content marked with Word's vanish property
- **SpecAgent references** - Visible watermarks, footers, and text-level URLs from specagent.com
- **Editorial artifacts** - Placeholders like `<insert>`, `RETAIN OR DELETE`, `[TBD]`

### 2. Style Clean — Unused Style Removal
- Styles defined but never referenced in document content
- Dependency-aware: preserves styles needed by other used styles (basedOn, link, next)
- Protected styles (Normal, Heading1-9, etc.) are never removed

### 3. Deep Clean — ZIP/XML-Level Optimization

| Category | What Gets Removed | Typical Savings |
|----------|-------------------|-----------------|
| **Orphaned Media** | Images in `word/media/` that nothing references | Variable (KB-MB) |
| **Orphaned Styles** | Style definitions never used in document content | ~500 bytes each |
| **RSID Tracking** | RSID attributes and RSID registries in settings.xml | ~25 bytes each, often 1000s per doc |
| **Empty Elements** | Empty runs `<w:r/>`, empty properties `<w:rPr/>` | ~20 bytes each |
| **Non-English Fonts** | Font mappings for Japanese, Arabic, Hebrew, etc. in theme | ~60 bytes each |
| **Compat Blocks** | Entire `<w:compat>` compatibility blocks from settings.xml | ~200 bytes each |
| **Internal Bookmarks** | `_GoBack`, `_Hlk*`, `_Ref*` bookmarks | ~80 bytes each |
| **Proof State** | Spell/grammar check state markers | ~40 bytes each |

After cleaning, a verification pass compares input vs output to flag any unexpected removals.

## Installation

### Requirements

- Python 3.10+
- Windows (primary), Linux/macOS (compatible)

### Setup

1. Clone or download this repository
2. Install dependencies: `pip install -r requirements.txt`
3. Run: `python gui.py`

## Usage

```bash
python gui.py
```

1. Click **Add Files...** to select one or more `.docx` files
2. Optionally choose an **Output Folder** (defaults to same folder as input, with `_cleaned` suffix)
3. Click **CLEAN**

The log area shows live progress per file, including item counts for each cleaning stage and before/after file sizes. Processing runs in a background thread so the UI stays responsive.

## How It Works

### Pipeline

```
input.docx
  → Shallow: unpack ZIP → detect editorial content → remove → repack
  → Styles: unpack → analyze dependency graph → remove unused → repack
  → Deep: unpack → orphan analysis → cruft scan → remove all → validate → repack
  → Verify: compare input vs output, classify removals
  → output_cleaned.docx
```

### Confidence Scoring

Each detection has a confidence score (0.0 - 1.0):
- Pattern match alone: ~0.6
- Pattern + italic: ~0.8
- Pattern + color: ~0.9
- Style name match: ~0.8+

Content is removed if confidence >= 0.5 and no preserve pattern matches.

## Detection Patterns

SpecCleanse uses configurable patterns defined in `patterns.yaml`. You can customize detection rules for your specific needs.

### Content Types

1. **Specifier Notes** (`specifier_notes`) - `[Specifier: ...]`, `NOTE TO SPECIFIER`, italic+red/blue text
2. **Copyright Notices** (`copyright_notices`) - `(c)`, `Copyright 2024`, `All rights reserved`
3. **Hidden Text** (`hidden_text`) - Any content with Word's `<w:vanish/>` property
4. **SpecAgent References** (`specagent_references`) - URLs, watermarks, attribution from specagent.com
5. **Editorial Artifacts** (`editorial_artifacts`) - `RETAIN OR DELETE`, `<INSERT>`, `[TBD]`
6. **Preserve Patterns** (`preserve_patterns`) - Content that should NEVER be removed (`END OF SECTION`, `PART 1`)

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

## File Structure

```
Spec_Cleanse/
├── gui.py             # Tkinter GUI — entry point and orchestration
├── detection.py       # Detection engine and detector classes
├── processor.py       # DOCX unpacking/repacking and content removal
├── deep_cleaner.py    # Orphan analysis and deep cleaning
├── style_cleaner.py   # Unused style detection and removal
├── verify.py          # Post-processing verification
├── patterns.yaml      # Configurable detection patterns
├── requirements.txt   # Pinned Python dependencies
├── CLAUDE.md          # AI assistant guide for the codebase
└── README.md          # This file
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
    ├── theme/
    │   └── theme1.xml      # Theme colors/fonts
    └── media/              # Embedded images
        ├── image1.png
        └── image2.jpg
```

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
- Add custom patterns for your content

**Too much removed**
- Add patterns to `preserve_patterns` section in `patterns.yaml`
- Check the verification output in the GUI log for unexpected removals

**Document won't open after cleaning**
- Check the error/warning messages in the GUI log
- Deep clean validates structure before completing

**File size didn't change much**
- Orphan removal typically saves a few KB
- The big wins come from RSID stripping (often 25-125 KB)

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

See `CLAUDE.md` for a detailed codebase guide covering architecture, conventions, and common modification scenarios.

## Copyright Notice

**Copyright © 2025 Abraham Borg. All Rights Reserved.**

This software and associated documentation files (the "Software") are the proprietary property of Andrew Gossman. 

**Unauthorized copying, modification, distribution, or use of this Software, via any medium, is strictly prohibited without express written permission from the copyright holder.**

This Software is provided for review and reference purposes only. No license or right to use, copy, modify, or distribute this Software for any purpose, commercial or non-commercial, is granted.
