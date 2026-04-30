# CLAUDE.md — AI Assistant Guide for SpecCleanse

## Project Overview

SpecCleanse is a Python GUI tool that removes editorial noise from specification Word documents (.docx) before LLM analysis. It targets architectural/engineering specification workflows where master spec templates (MasterSpec, BSD SpecLink, ARCOM) accumulate specifier notes, copyright boilerplate, hidden text, and editing instructions that should be stripped before further processing.

The application runs through a Tkinter GUI and performs **single-pass shallow content removal** followed by an automatic verification pass.

## Architecture

### Single-Stage Cleaning Pipeline

Every clean run performs:
1. **Shallow content removal** — pattern + formatting detection inside `document.xml`, headers, footers, footnotes, and endnotes
2. **Verification** — input/output comparison classifies every removed paragraph as expected, unexpected, or a preserve violation

There is no longer a "deep clean" or "style clean" stage in the active pipeline. Earlier versions had ZIP/XML structural optimization and unused-style removal stages; both were retired because they changed document metadata without meaningfully improving downstream LLM extraction. Their source still lives in `legacy/` for reference but is not imported by the running app.

### Module Responsibilities

| Module | Purpose |
|--------|---------|
| `gui.py` | Tkinter GUI — entry point, runs preview/clean in a background thread, manages logging and progress |
| `detection.py` | Pattern matching engine with confidence scoring; all detector classes |
| `processor.py` | DOCX unpacking/repacking, XML walking, element removal |
| `verify.py` | Post-processing verification comparing input vs. output paragraphs |
| `legacy/deep_cleaner.py` | Archived; not used |
| `legacy/style_cleaner.py` | Archived; not used |

### Configuration

| File | Purpose |
|------|---------|
| `patterns.yaml` | All detection patterns, formatting signals, preserve rules |
| `requirements.txt` | Pinned Python dependencies (UTF-8) |

### Data Flow

```
input.docx
  → unpack ZIP → parse XML (document/headers/footers/footnotes/endnotes)
  → detect → remove (collect-then-remove with tail-text preservation)
  → repack
  → verify: diff input vs. output, classify removals
  → output_cleaned.docx
```

## Key Design Patterns

### Strategy Pattern for Detectors

All detectors extend `BaseDetector` in `detection.py` and implement `detect(element, text) -> Optional[Detection]`. The `DetectionEngine` orchestrates them.

```
BaseDetector
├── SpecifierNoteDetector
├── CopyrightDetector
├── HiddenTextDetector
├── SpecAgentDetector
├── EditorialArtifactDetector
└── PreserveDetector (short-circuits removals)
```

To add a new detector:
1. Add patterns to `patterns.yaml`
2. Create a detector class in `detection.py` extending `BaseDetector`
3. Register it in `DetectionEngine._create_detectors()`

### Confidence Scoring

Detections are scored 0.0–1.0. Multiple signals combine:
- Text pattern match alone: ~0.6 (specifier notes), 0.7 (copyright), 0.8 (high-confidence editorial), 1.0 (preserve / SpecAgent / hidden)
- Italic formatting: +0.2
- Editorial color (red, dark red, blue, light blue): +0.3
- Editorial paragraph or character style: +0.8
- Removal threshold: confidence ≥ 0.5
- Preserve patterns short-circuit removals regardless of confidence

`editorial_artifacts` uses a two-tier scheme: `text_patterns` are high-confidence and remove on text alone; `low_confidence_patterns` start at 0.3 and require a formatting signal to cross the threshold.

### Dataclass-Based Results

Processing results are communicated via dataclasses, not exceptions:
- `Detection` — individual content detection with confidence
- `ProcessingResult` — aggregated results with errors list
- `RemovedParagraph` / `VerificationResult` — verification output

Errors accumulate in result objects; processing doesn't halt on non-fatal issues.

### Direct XML Manipulation

The project uses `lxml` for direct XML manipulation rather than `python-docx`. This gives lower-level control needed for:
- Preserving exact formatting through XML structure
- Handling Word namespace complexity
- Safe element removal with tail-text preservation

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

Elements marked for removal are collected during iteration, then removed in a second pass. Tail text (text nodes after an element) is preserved by appending to the previous sibling or parent. Parent existence is verified before any removal.

### Files Walked

`processor.py` and `verify.py` both walk the same set of XML files inside the `word/` directory:
- `document.xml`
- `header*.xml`
- `footer*.xml`
- `footnotes.xml`
- `endnotes.xml`

Keep these two lists in sync. If you add coverage for a new XML file in one place, add it to the other.

## Code Conventions

### Style

- **Python 3.10+** — uses `str | None` union syntax, `list[str]` generics
- **snake_case** for functions and variables
- **PascalCase** for classes
- **UPPER_CASE** for constants and module-level namespace strings
- **Type hints** on all function signatures
- **Dataclasses** for structured data (not plain dicts)
- **Enum** types for categorical constants (`ContentType`)
- **Docstrings** on all modules, classes, and public methods

### Dependencies

Only two external dependencies — keep it minimal:
- `lxml==6.0.2` — XML parsing and manipulation
- `PyYAML==6.0.3` — YAML configuration loading

`tkinter` is part of the standard library. Do not add new dependencies without strong justification.

### Error Handling

- Result objects accumulate errors without stopping execution
- Graceful degradation: warnings don't fail the entire operation
- GUI continues processing remaining files even if one fails

### File Organization

- All source files are flat in the project root (no `src/` directory)
- No packaging infrastructure (no `setup.py`, `pyproject.toml`)
- Entry point: `python gui.py`
- Imports are relative within the project (e.g., `from detection import DetectionEngine`)

## How to Run

### Prerequisites

```bash
pip install -r requirements.txt
```

### Usage

```bash
python gui.py
```

1. Click **Add Files...** to select one or more `.docx` files
2. Optionally click **Output Folder...** (defaults to same folder as input, with `_cleaned` suffix)
3. Click **Preview** for a dry-run report, or **CLEAN** to write `*_cleaned.docx`

Processing runs on a background thread with a live log and progress bar.

## Testing

### Current Approach

There is no automated test suite. Testing is done manually with sample DOCX files placed in the repository root. Test output goes to `spec_testing/` (gitignored).

### Manual Testing Workflow

Run the GUI against sample files. The log output shows:
- Per-file removed/preserved counts
- Before/after file sizes
- Verification results (expected, unexpected, preserve violations)

### When Making Changes

1. Run the GUI against representative DOCX files (with and without footnotes/headers)
2. Open the output in Word to verify formatting is preserved
3. Check the verification output for unexpected removals or preserve violations

## Common Modification Scenarios

### Adding a New Detection Pattern

1. Add regex patterns to the appropriate section in `patterns.yaml`
2. Run the GUI's Preview mode and check the log to verify matches
3. No code changes needed for simple pattern additions

### Adding a New Content Type

1. Add the type to `ContentType` enum in `detection.py`
2. Create a new detector class extending `BaseDetector`
3. Register it in `DetectionEngine._create_detectors()`
4. Add corresponding patterns to `patterns.yaml`
5. Add the new category to `_build_removal_patterns` in `verify.py` so verification recognizes it

### Modifying XML Processing

- Always test with documents containing headers, footers, and footnotes
- Verify namespace handling — Word uses many namespaces
- Handle tail text preservation when removing elements
- Never modify XML during iteration; collect targets first, then remove

## Important Caveats

- **DOCX only** — does not handle `.doc` (legacy binary format)
- **Direct XML manipulation** — not using `python-docx`, so changes must be XML-aware
- **No structural/style optimization** — those stages were retired; the cleaner only removes content
- **Temp files** are created with `tempfile.mkdtemp(prefix="speccleanse_")` and cleaned up in `finally` blocks
- **`patterns.yaml`** must be in the same directory as `gui.py`
