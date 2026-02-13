# CLAUDE.md — AI Assistant Guide for SpecCleanse

## Project Overview

SpecCleanse is a Python GUI tool that removes unnecessary content from specification Word documents (.docx) while preserving all formatting and styles. It targets architectural/engineering specification workflows where master spec templates (MasterSpec, BSD SpecLink, ARCOM) accumulate editorial content, tracking metadata, and other cruft that should be removed before final publication.

The application runs through a Tkinter GUI and always performs a maximum-strength clean: shallow content removal + unused style cleanup + aggressive deep clean with all operations enabled.

## Architecture

### Three-Stage Cleaning Pipeline

Every clean runs all three stages in sequence:
1. **Shallow Clean** — Content-level removal of editorial/specifier content using confidence-scored pattern detection
2. **Style Clean** — Unused style removal via dependency graph analysis
3. **Deep Clean** — ZIP/XML-level optimization that removes orphaned resources, RSID tracking, empty elements, compat blocks, and other accumulated cruft

### Module Responsibilities

| Module | Purpose |
|--------|---------|
| `gui.py` | Tkinter GUI — entry point, orchestrates all three cleaning stages, verification, progress/logging |
| `detection.py` | Pattern matching engine with confidence scoring; all detector classes |
| `processor.py` | DOCX unpacking/repacking (`repack_docx` utility), XML content walking, element removal |
| `deep_cleaner.py` | Orphan analysis, RSID stripping, cruft removal at ZIP/XML level |
| `style_cleaner.py` | Unused style detection via dependency graph analysis |
| `verify.py` | Post-processing verification comparing input vs output |

### Configuration

| File | Purpose |
|------|---------|
| `patterns.yaml` | All detection patterns, formatting signals, preserve rules (fully customizable) |
| `requirements.txt` | Pinned Python dependencies |

### Data Flow

```
input.docx
  → Shallow: unpack ZIP → parse XML → detect → remove → repack
  → Styles: unpack output → analyze dependency graph → remove unused → repack
  → Deep: unpack output → orphan analysis → cruft scan → remove all → validate → repack
  → Verify: compare input vs output, classify removals
  → output_cleaned.docx
```

## Key Design Patterns

### Strategy Pattern for Detectors

All detectors extend `BaseDetector` in `detection.py` and implement `detect(element, text) -> Optional[Detection]`. The `DetectionEngine` orchestrates all registered detectors.

```
BaseDetector
├── SpecifierNoteDetector
├── CopyrightDetector
├── HiddenTextDetector
├── SpecAgentDetector
├── EditorialArtifactDetector
└── PreserveDetector (overrides removals)
```

To add a new detector:
1. Add patterns to `patterns.yaml`
2. Create a detector class in `detection.py` extending `BaseDetector`
3. Register it in `DetectionEngine._create_detectors()`

### Confidence Scoring

Detections are scored 0.0–1.0. Multiple signals boost confidence:
- Text pattern match alone: ~0.6
- Pattern + italic formatting: +0.2
- Pattern + color match (red/blue): +0.3
- Style name match: ~0.8+
- Removal threshold: confidence >= 0.5
- Preserve patterns always override removals regardless of confidence

### Dataclass-Based Results

Processing results are communicated via dataclasses, not exceptions:
- `Detection` — individual content detection with confidence
- `ProcessingResult` — aggregated results with errors list
- `OrphanReport` — structured deep-clean findings
- `DeepCleanResult` — deep cleaning outcomes
- `StyleCleanResult` — style cleaning outcomes

Errors accumulate in result objects; processing doesn't halt on non-fatal issues.

### Direct XML Manipulation

The project uses `lxml` for direct XML manipulation rather than `python-docx`. This gives lower-level control needed for:
- Preserving exact formatting through XML structure
- Handling Word namespace complexity
- Safe element removal with tail-text preservation
- RSID stripping via regex on raw XML strings (performance optimization)

### XML Namespace Handling

Word namespaces are defined consistently across modules:
```python
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}
```

### Safe Element Removal

Elements marked for removal are collected during iteration, then removed in a second pass. Tail text (text nodes after an element) is preserved by appending to the previous sibling or parent. Parent link existence is verified before any removal.

## Code Conventions

### Style

- **Python 3.10+** — uses `str | None` union syntax, `list[str]` generics
- **snake_case** for functions and variables
- **PascalCase** for classes
- **UPPER_CASE** for constants and module-level namespace strings
- **Type hints** on all function signatures
- **Dataclasses** with `@dataclass` for structured data (not plain dicts)
- **Enum** types for categorical constants (`ContentType`)
- **Docstrings** on all modules, classes, and public methods

### Dependencies

Only two external dependencies — keep it minimal:
- `lxml==6.0.2` — XML parsing and manipulation
- `PyYAML==6.0.3` — YAML configuration loading

Do not add new dependencies without strong justification.

### Error Handling

- Results objects accumulate errors without stopping execution
- Graceful degradation: warnings don't fail the entire operation
- Validation of document structure post-cleaning (deep clean)
- GUI continues processing remaining files even if one fails

### File Organization

- All source files are flat in the project root (no `src/` directory)
- No packaging infrastructure (no `setup.py`, `pyproject.toml`)
- Entry point: `python gui.py`
- Imports are relative within the project (e.g., `from detection import DetectionEngine`)

## How to Run

### Prerequisites

```bash
pip install lxml pyyaml
```

### Usage

```bash
python gui.py
```

The GUI runs maximum-strength cleaning (shallow + styles + deep with all options enabled, including aggressive compat removal) on one or more DOCX files. It uses tkinter (Python standard library — no extra dependencies). Processing runs in a background thread with a live log and progress bar.

1. Click **Add Files...** to select one or more `.docx` files
2. Optionally choose an **Output Folder** (defaults to same folder as input, with `_cleaned` suffix)
3. Click **CLEAN**

## Testing

### Current Approach

There is no automated test suite. Testing is done manually with sample DOCX files in the repository root:
- `NVES.docx`
- `the_grove_spec.docx`
- `P247050.00 - TrueCare 1595 - Specs.docx`
- `weird spec.docx`

Test output goes to `spec_testing/` (gitignored).

### Manual Testing Workflow

Run the GUI against sample files. The log output shows:
- Per-stage item counts (shallow removals, styles removed, deep clean stats)
- Before/after file sizes
- Verification results (expected vs unexpected removals, preserve violations)

### When Making Changes

1. Run the GUI against all sample DOCX files
2. Open the output in Word to verify formatting is preserved
3. Compare file sizes before/after
4. Check the verification output for unexpected removals or preserve violations

## Common Modification Scenarios

### Adding a New Detection Pattern

1. Add regex patterns to the appropriate section in `patterns.yaml`
2. Run the GUI and check the log output to verify matches
3. No code changes needed for simple pattern additions

### Adding a New Content Type

1. Add the type to `ContentType` enum in `detection.py`
2. Create a new detector class extending `BaseDetector`
3. Register it in `DetectionEngine._create_detectors()`
4. Add corresponding patterns to `patterns.yaml`

### Adding a New Deep Clean Operation

1. Add analysis logic in `OrphanAnalyzer` in `deep_cleaner.py` (scan phase)
2. Add removal logic in `DeepCleaner` (clean phase)
3. Call the new method from `DeepCleaner.clean()`
4. Update `DeepCleanResult` dataclass if new metrics are tracked
5. Add logging for the new metric in `gui.py` `_clean_one()`

### Modifying XML Processing

- Always test with documents containing headers, footers, and footnotes (not just document.xml)
- Verify namespace handling — Word uses many namespaces and different XML files may need different namespace maps
- When removing elements, handle tail text preservation
- Never modify XML during iteration; collect targets first, then remove

## Important Caveats

- **DOCX only** — does not handle `.doc` (legacy binary format)
- **Direct XML manipulation** — not using `python-docx`, so changes must be XML-aware
- **RSID stripping uses regex** on serialized XML strings for performance, not parsed XML
- **Style dependency resolution** uses transitive closure (basedOn -> link -> next chains)
- **Protected styles** in `style_cleaner.py` (Normal, Heading1-9, etc.) must never be removed
- **Temp files** are created with `tempfile.mkdtemp(prefix="speccleanse_")` and cleaned up in `finally` blocks
- **The patterns.yaml file** must be in the same directory as `gui.py`
- **Deep clean always runs aggressive compat removal** — removes entire `<w:compat>` blocks
