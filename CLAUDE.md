# CLAUDE.md — AI Assistant Guide for SpecCleanse

## Project Overview

SpecCleanse is a Python tool (CLI + GUI) that removes unnecessary content from specification Word documents (.docx) while preserving all formatting and styles. It targets architectural/engineering specification workflows where master spec templates (MasterSpec, BSD SpecLink, ARCOM) accumulate editorial content, tracking metadata, and other cruft that should be removed before final publication.

## Architecture

### Two-Tier Cleaning Model

1. **Shallow Clean (default)** — Content-level removal of editorial/specifier content using confidence-scored pattern detection
2. **Deep Clean (`--deep`)** — ZIP/XML-level optimization that removes orphaned resources, RSID tracking, empty elements, and other accumulated cruft

### Module Responsibilities

| Module | Lines | Purpose |
|--------|-------|---------|
| `gui.py` | ~410 | Tkinter GUI — runs maximum-strength clean (shallow + styles + deep, all options enabled) |
| `speccleanse.py` | ~590 | CLI entry point, argument parsing, orchestration |
| `detection.py` | ~405 | Pattern matching engine with confidence scoring; all detector classes |
| `processor.py` | ~310 | DOCX unpacking/repacking (`repack_docx` utility), XML content walking, element removal |
| `deep_cleaner.py` | ~1360 | Orphan analysis, RSID stripping, cruft removal at ZIP/XML level |
| `style_cleaner.py` | ~315 | Unused style detection via dependency graph analysis |
| `diagnose.py` | ~305 | Standalone diagnostic utility for inspecting document formatting |

### Configuration

| File | Purpose |
|------|---------|
| `patterns.yaml` | All detection patterns, formatting signals, preserve rules (fully customizable) |
| `requirements.txt` | Pinned Python dependencies |

### Data Flow

```
input.docx
  → unpack ZIP to temp dir
  → parse XML (document.xml, headers, footers)
  → run detectors (confidence scoring per element)
  → apply preserve-pattern overrides
  → remove elements where confidence ≥ 0.5
  → [if --deep] orphan analysis → cruft scan → remove → validate
  → repack ZIP to output.docx
  → cleanup temp dir
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
- Removal threshold: confidence ≥ 0.5
- Preserve patterns always override removals regardless of confidence

### Dataclass-Based Results

Processing results are communicated via dataclasses, not exceptions:
- `Detection` — individual content detection with confidence
- `ProcessingResult` — aggregated results with errors list
- `OrphanReport` — structured deep-clean findings
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
- `sys.exit(1)` only at CLI level for fatal errors (missing files, invalid input)

### File Organization

- All source files are flat in the project root (no `src/` directory)
- No packaging infrastructure (no `setup.py`, `pyproject.toml`)
- Designed for direct script execution: `python speccleanse.py ...`
- Imports are relative within the project (e.g., `from detection import DetectionEngine`)

## How to Run

### Prerequisites

```bash
pip install lxml pyyaml
```

### GUI Usage

```bash
python gui.py
```

The GUI runs maximum-strength cleaning (shallow + styles + deep with all options enabled, including aggressive compat removal) on one or more DOCX files. It uses tkinter (Python standard library — no extra dependencies). Processing runs in a background thread with a live log and progress bar.

### CLI Usage

```bash
# Shallow clean (default)
python speccleanse.py input.docx output.docx

# Full clean (shallow + deep)
python speccleanse.py input.docx output.docx --deep

# Preview without modifying
python speccleanse.py input.docx output.docx --deep --dry-run -v

# Deep clean only (skip content removal)
python speccleanse.py input.docx output.docx --deep-only

# Diagnose document formatting
python diagnose.py input.docx -e
```

### Key CLI Flags

| Flag | Purpose |
|------|---------|
| `--deep` | Enable deep cleaning (orphans + cruft) |
| `--deep-only` | Only deep clean, skip shallow |
| `--styles-only` | Only clean unused styles |
| `--dry-run` / `-d` | Preview without modifying |
| `--verbose` / `-v` | Detailed detection output |
| `--quiet` / `-q` | Suppress output except errors |
| `--no-media`, `--no-rsids`, etc. | Selectively disable deep-clean operations |
| `--strip-links-domain DOMAIN` | Remove hyperlinks for specific domains |
| `--only OPERATION` | Run only a single deep-clean operation (for debugging) |

## Testing

### Current Approach

There is no automated test suite. Testing is done manually with sample DOCX files in the repository root:
- `NVES.docx`
- `the_grove_spec.docx`
- `P247050.00 - TrueCare 1595 - Specs.docx`
- `weird spec.docx`

Test output goes to `spec_testing/` (gitignored).

### Manual Testing Workflow

```bash
# Preview what would be removed
python speccleanse.py "test_file.docx" output.docx --deep --dry-run -v

# Run actual clean and inspect output in Word
python speccleanse.py "test_file.docx" output.docx --deep -v
```

### When Making Changes

1. Run against all sample DOCX files with `--dry-run -v` to verify detection behavior
2. Run a full clean and open the output in Word to verify formatting is preserved
3. Compare file sizes before/after for deep clean changes
4. Check that `--deep-only` and `--styles-only` modes still work independently

## Common Modification Scenarios

### Adding a New Detection Pattern

1. Add regex patterns to the appropriate section in `patterns.yaml`
2. Test with `--dry-run -v` to verify matches
3. No code changes needed for simple pattern additions

### Adding a New Content Type

1. Add the type to `ContentType` enum in `detection.py`
2. Create a new detector class extending `BaseDetector`
3. Register it in `DetectionEngine._create_detectors()`
4. Add corresponding patterns to `patterns.yaml`
5. Update `print_result()` in `speccleanse.py` if special reporting is needed

### Adding a New Deep Clean Operation

1. Add analysis logic in `deep_cleaner.py` (scan phase)
2. Add removal logic in the clean phase
3. Add a `--no-<operation>` CLI flag in `speccleanse.py`
4. Update `OrphanReport` dataclass if new data is tracked

### Modifying XML Processing

- Always test with documents containing headers, footers, and footnotes (not just document.xml)
- Verify namespace handling — Word uses many namespaces and different XML files may need different namespace maps
- When removing elements, handle tail text preservation
- Never modify XML during iteration; collect targets first, then remove

## Important Caveats

- **DOCX only** — does not handle `.doc` (legacy binary format)
- **Direct XML manipulation** — not using `python-docx`, so changes must be XML-aware
- **RSID stripping uses regex** on serialized XML strings for performance, not parsed XML
- **Style dependency resolution** uses transitive closure (basedOn → link → next chains)
- **Protected styles** in `style_cleaner.py` (Normal, Heading1-9, etc.) must never be removed
- **Temp files** are created with `tempfile.mkdtemp(prefix="speccleanse_")` and cleaned up in `finally` blocks
- **The patterns.yaml file** must be in the same directory as `speccleanse.py` by default (or specified with `-c`)
