# SpecCleanse - Deep DOCX Cleaning Extension

## Overview

This extension adds **deep DOCX cleaning** capabilities to SpecCleanse, allowing you to remove accumulated cruft from Word documents at the internal ZIP/XML level.

## What It Cleans

### Orphaned Resources

| Resource Type | What Gets Removed | Typical Savings |
|---------------|-------------------|-----------------|
| **Relationships** | Orphaned hyperlinks (like SpecAgent.com refs), unused internal links | ~200 bytes each |
| **Media Files** | Images in `word/media/` that nothing references | Variable (KB-MB) |
| **Styles** | Style definitions never used in document content | ~500 bytes each |
| **Fonts** | Font declarations not referenced anywhere | Minimal |
| **Numbering** | List definitions without any paragraphs using them | Minimal |
| **WebSettings Divs** | Copy/paste artifacts that bloat the file | ~200 bytes each |

### Cruft (New!)

| Cruft Type | What Gets Removed | Typical Savings |
|------------|-------------------|-----------------|
| **RSID Attributes** | Revision tracking IDs on every element (`w:rsidR`, `w:rsidRPr`, etc.) | ~25 bytes each, often 1000s per doc |
| **Empty Elements** | Empty runs `<w:r/>`, empty properties `<w:rPr/>` | ~20 bytes each |
| **Non-English Fonts** | Font mappings for Japanese, Arabic, Hebrew, etc. in theme | ~60 bytes each, 40+ per theme |
| **Compat Settings** | Backwards compatibility for Word 97/2002/2003 | ~50 bytes each |
| **Internal Bookmarks** | `_GoBack`, `_Hlk*`, `_Ref*` bookmarks | ~80 bytes each |
| **Proof State** | Spell/grammar check state markers | ~40 bytes each |

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Deep Clean Pipeline                       │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌──────────────────┐    ┌─────────────────────┐           │
│  │ docx_obliterate  │───▶│ docx_orphan_analyzer │           │
│  │    .py           │    │        .py           │           │
│  │                  │    │                      │           │
│  │ • Extract DOCX   │    │ • Find orphaned rIds │           │
│  │ • Full analysis  │    │ • Find unused styles │           │
│  │ • Reconstruct    │    │ • Find dead media    │           │
│  └──────────────────┘    └──────────┬──────────┘           │
│                                      │                       │
│                                      ▼                       │
│                          ┌─────────────────────┐            │
│                          │   orphans.json      │            │
│                          │   orphans.yaml      │            │
│                          └──────────┬──────────┘            │
│                                      │                       │
│                                      ▼                       │
│                          ┌─────────────────────┐            │
│                          │ docx_deep_cleaner   │            │
│                          │        .py          │            │
│                          │                     │            │
│                          │ • Remove orphans    │            │
│                          │ • Validate structure│            │
│                          │ • Backup/rollback   │            │
│                          │ • Repack DOCX       │            │
│                          └─────────────────────┘            │
│                                                              │
└─────────────────────────────────────────────────────────────┘
```

## Usage

### Step 1: Extract and Analyze

```bash
# Extract the DOCX and generate full analysis
python docx_obliterate.py MECH_SPEC.docx

# Output:
#   MECH_SPEC_extracted/        (extracted contents)
#   MECH_SPEC_extracted_analysis.md  (full analysis report)
```

### Step 2: Find Orphans

```bash
# Analyze for orphaned resources
python docx_orphan_analyzer.py MECH_SPEC_extracted/

# Output:
#   MECH_SPEC_extracted_orphans.json  (machine-readable report)
#   MECH_SPEC_extracted_orphans.yaml  (human-readable manifest)
```

### Step 3: Review the Manifest

Check `MECH_SPEC_extracted_orphans.yaml` before cleaning:

```yaml
# DOCX Cleanup Manifest
# Safe to remove the following orphaned resources

orphaned_relationships:
  - rId: rId79
    target: http://www.specagent.com/LookUp/?uid=123456811839
    type: hyperlink
    reason: "No r:id reference found in content files"

orphaned_styles:
  - styleId: TB4
    name: TB4
    type: paragraph
    reason: "Never referenced in document content or by other styles"

estimated_savings_bytes: 12400
```

### Step 4: Deep Clean (Dry Run First!)

```bash
# See what would be removed without making changes
python docx_deep_cleaner.py MECH_SPEC_extracted/ MECH_SPEC_extracted_orphans.json --dry-run

# Actually perform the cleaning
python docx_deep_cleaner.py MECH_SPEC_extracted/ MECH_SPEC_extracted_orphans.json

# Output:
#   MECH_SPEC_cleaned.docx
```

### Selective Cleaning

Skip certain types of cleaning if needed:

```bash
# Only remove orphaned resources, skip cruft stripping
python docx_deep_cleaner.py extracted/ orphans.json \
    --no-rsids --no-empty --no-fonts --no-compat --no-bookmarks --no-proof

# Only strip RSIDs (biggest savings, safest operation)
python docx_deep_cleaner.py extracted/ orphans.json \
    --no-relationships --no-media --no-styles \
    --no-empty --no-fonts --no-compat --no-bookmarks --no-proof

# Full clean except non-English fonts (if multilingual doc)
python docx_deep_cleaner.py extracted/ orphans.json --no-fonts
```

### CLI Flags

**Orphan Removal:**
- `--no-relationships` - Keep orphaned hyperlinks/relationships
- `--no-media` - Keep orphaned media files
- `--no-styles` - Keep orphaned style definitions

**Cruft Removal:**
- `--no-rsids` - Keep RSID tracking attributes
- `--no-empty` - Keep empty runs/elements
- `--no-fonts` - Keep non-English font mappings
- `--no-compat` - Keep backwards compatibility settings
- `--no-bookmarks` - Keep internal Word bookmarks
- `--no-proof` - Keep spell/grammar check state

## Integration with SpecCleanse

The deep cleaner is designed to work with the existing SpecCleanse pipeline:

```bash
# Full pipeline: content strip → deep clean → output
python speccleanse.py input.docx -o stripped.docx
python docx_obliterate.py stripped.docx
python docx_orphan_analyzer.py stripped_extracted/
python docx_deep_cleaner.py stripped_extracted/ stripped_extracted_orphans.json -o final.docx
```

Or integrate into a single workflow (future enhancement).

## Safety Features

### Backup & Rollback
- Automatic backup created before any modifications
- If validation fails, automatically restores from backup

### Validation Checks
Before completing, the cleaner validates:
- `[Content_Types].xml` exists and parses correctly
- `_rels/.rels` exists with valid structure  
- `word/document.xml` exists and parses correctly

### Conservative Defaults
- Essential relationship types are never removed (styles, settings, theme, etc.)
- Built-in Word styles (`Normal`, `DefaultParagraphFont`, etc.) are protected
- Style dependency chains are fully resolved before marking anything orphaned

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
    ├── webSettings.xml     # Web-related settings (paste artifacts live here)
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

These map to entries in `document.xml.rels`:

```xml
<Relationship Id="rId8" Type="...hyperlink" Target="http://example.com"/>
<Relationship Id="rId5" Type="...image" Target="media/image1.png"/>
```

**Orphaned relationships** occur when an rId exists in the `.rels` file but nothing in the document references it.

## Common Specification Cruft

Master spec templates (like MasterSpec, BSD SpecLink, ARCOM) often accumulate:

1. **SpecAgent hyperlinks** - Product lookup URLs that should be removed for final specs
2. **Unused styles** - Template styles for sections you deleted
3. **Paste artifacts** - `<w:div>` elements in webSettings.xml from copy/paste
4. **Revision history** - Tracked changes that were accepted but leave cruft
5. **Dead media** - Images from deleted sections

## Troubleshooting

### "Document won't open after cleaning"

1. Check the error messages in the cleaning output
2. The backup directory (`*_backup`) should still exist - restore from there
3. Run with `--dry-run` first to see what would be removed
4. Try selective cleaning (e.g., `--no-styles`) to isolate the problem

### "Still seeing SpecAgent references"

The deep cleaner removes orphaned relationships. If SpecAgent hyperlinks are still in the document content, you need to run SpecCleanse first to remove the actual hyperlink elements.

### "File size didn't change much"

Orphan removal typically saves a few KB. The big wins come from:
- Removing large unused media files
- Running after SpecCleanse removes lots of content (which creates orphans)

## Files

| File | Purpose |
|------|---------|
| `docx_obliterate.py` | Extract DOCX, generate full analysis, reconstruct |
| `docx_orphan_analyzer.py` | Find orphaned resources, generate cleanup manifest |
| `docx_deep_cleaner.py` | Remove orphans, validate, repack |
