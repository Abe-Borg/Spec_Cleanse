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
| `processor.py` | DOCX unpacking/repacking, XML walking, element removal, inline redaction |
| `verify.py` | Post-processing verification: removals, modifications, structural lint |
| `docx_xml.py` | Shared WordprocessingML plumbing: namespaces, iteration, text extraction, structure rules, style resolution, config loading |
| `apppaths.py` | Runtime file locations: which `patterns.yaml` to load from source vs. a frozen build |
| `tests/` | stdlib `unittest` suite; builds synthetic DOCX files with `zipfile` |
| `legacy/deep_cleaner.py` | Archived; not used |
| `legacy/style_cleaner.py` | Archived; not used |

`docx_xml.py` holds document-model plumbing only — it knows how WordprocessingML
nests runs inside paragraphs and which containers Word refuses to open when empty,
but no detection policy. Anything that decides *what* to remove belongs in
`detection.py`.

### Configuration

| File | Purpose |
|------|---------|
| `patterns.yaml` | All detection patterns, formatting signals, styles, preserve rules |
| `requirements.txt` | Pinned Python dependencies (UTF-8) |
| `CHANGELOG.md` | Release history, Keep a Changelog format |
| `requirements-build.txt` | Build-time only (PyInstaller); not needed to run from source |
| `packaging/speccleanse.spec` | PyInstaller build definition |
| `packaging/installer.iss` | Inno Setup installer definition |

### Data Flow

```
input.docx
  → unpack ZIP → bind styles from word/styles.xml
  → parse XML (document/headers/footers/footnotes/endnotes/glossary)
  → optional: accept tracked changes, strip comments
  → detect → collect targets → redact spans, remove runs, remove paragraphs
  → repack (temp file, then os.replace)
  → verify: diff input vs. output, classify removals and modifications,
            compare structure
  → output_cleaned.docx
```

## Key Design Patterns

### Strategy Pattern for Detectors

All detectors extend `BaseDetector` in `detection.py` and implement `detect(element, text) -> Optional[Detection]`. The `DetectionEngine` orchestrates them.

```
BaseDetector
├── SpecifierNoteDetector
├── CopyrightDetector
├── HiddenTextDetector          (run w:vanish, or inherited from a hidden style)
├── SpecAgentDetector
├── EditorialArtifactDetector   (also emits INLINE_PLACEHOLDER span detections)
└── PreserveDetector            (short-circuits removals; patterns and styles)
```

To add a new detector:
1. Add patterns to `patterns.yaml`
2. Create a detector class in `detection.py` extending `BaseDetector`
3. Register it in `DetectionEngine._create_detectors()`

Verification derives its categories from `DetectionEngine.removal_patterns()`, so a
new detector's patterns are recognised there without a second registration.

### Confidence Scoring

Detections are scored 0.0–1.0. Multiple signals combine:
- Text pattern match alone: ~0.6 (specifier notes), 0.7 (copyright), 0.8 (high-confidence editorial), 1.0 (preserve / SpecAgent / hidden)
- Italic formatting: +0.2
- Editorial color (red, dark red, blue, light blue): +0.3
- Editorial paragraph or character style: +0.8
- Removal threshold: confidence ≥ 0.5
- Preserve patterns and preserve styles short-circuit removals regardless of confidence

Evidence is kept in two buckets. *Content* evidence — a text pattern or an editorial
style — says what the text is; *formatting* evidence — italic, editorial colour —
says only how it looks, and the two formatting signals together land on exactly the
0.5 threshold. `specifier_notes.formatting_only_removal` (default true) decides
whether formatting alone is enough; when it is false, formatting only boosts a score
that content evidence already opened. Removals that crossed on formatting alone carry
`Detection.formatting_only` and are labelled in Preview.

`editorial_artifacts` uses three tiers: `text_patterns` are high-confidence and remove
the whole paragraph on text alone; `low_confidence_patterns` start at 0.3 and require a
formatting signal to cross the threshold; `inline_patterns` produce
`ContentType.INLINE_PLACEHOLDER` detections carrying character `spans`, which are cut
out of the paragraph in place. `DetectionEngine.should_remove()` ignores inline
detections — they never remove their element.

The inline tier must stay separable when *judging* a removal, which is why
`removal_patterns()` takes `include_inline`. An inline pattern matching a paragraph
that vanished entirely means only that the paragraph contained a placeholder — never
that losing it was intended. Verification asks the narrower question instead: does
cutting every placeholder leave nothing behind?

### Dataclass-Based Results

Processing results are communicated via dataclasses, not exceptions:
- `Detection` — individual content detection with confidence, spans, formatting flag
- `ProcessingResult` — detections plus an errors list
- `RemovedParagraph` / `ModifiedParagraph` / `StructuralViolation` / `VerificationResult` — verification output. Categories: a pattern name, `formatting_based`, `inline_placeholder`, `tracked_deletion`, `preserve_violation`, or `None` for unexplained
- `StyleInfo` / `StyleIndex` — `word/styles.xml` resolved for name and `w:basedOn` matching

Errors accumulate in result objects; processing doesn't halt on non-fatal issues.

### Direct XML Manipulation

The project uses `lxml` for direct XML manipulation rather than `python-docx`. This gives lower-level control needed for:
- Preserving exact formatting through XML structure
- Handling Word namespace complexity
- Safe element removal with tail-text preservation

### XML Namespace Handling

Word namespaces are defined once, in `docx_xml.py`, and imported from there — never redeclared in a module:
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

Elements marked for removal are collected during iteration, then removed in a second pass — mutating the tree while iterating it skips elements. Targets are applied in a fixed order: redactions (whose spans are offsets into the paragraph text as it stands), then runs, then paragraphs.

Deleting a `w:p` element is wrong in four situations, and `_remove_paragraph` empties the paragraph in place instead of deleting it in each of them:

| Situation | Why | Check |
|-----------|-----|-------|
| A field begins inside and ends later | Unbalanced `w:fldChar` is a file Word refuses to open | `field_chars_balanced()` |
| The paragraph carries `w:pPr/w:sectPr` | It is a section boundary; deleting it merges the section into the next and loses its headers, footers and page setup | `has_section_properties()` |
| It holds a picture, object, field, or note reference | The content is invisible to text patterns but not to the reader | `has_embedded_content()` |
| Its parent would be left with no block content | `CT_HdrFtr` and friends carry `minOccurs="1"`; a table cell must also *end* with a `w:p` | `can_delete_paragraph()` |

Half-open `w:bookmarkStart`/`w:bookmarkEnd` and comment-range markers are relocated beside the paragraph before it is deleted — they are legal at block level, so the range keeps spanning the same content.

There is no tail-text handling: in WordprocessingML an element tail is only inter-element whitespace, so preserving it protects nothing.

### Files Walked

`processor.py` and `verify.py` both walk the parts returned by `docx_xml.collect_content_parts()`, inside the `word/` directory:
- `document.xml`
- `header*.xml`
- `footer*.xml`
- `footnotes.xml`
- `endnotes.xml`
- `glossary/document.xml`

Both go through that one function, so the lists cannot drift. Add new coverage there, not in either caller.

### Iteration and Nesting

Every run lives inside a paragraph — directly, or through `w:hyperlink`, `w:ins`, `w:sdtContent`, `w:fldSimple`, `w:smartTag`. `iter_own_runs()` finds all of them without crossing into a paragraph nested in a text box, because that inner paragraph is visited in its own right by `iter_paragraphs()`. Never use `para.iter(w:r)` or `para.iter(w:t)` directly: it double-counts text-box content and attributes it to the wrong paragraph.

`mc:AlternateContent` stores the same shape twice, under `mc:Choice` and `mc:Fallback`. Processing walks both (each has to be cleaned); text extraction passes `skip_alternate_fallback=True` so the content is counted once.

### Text Extraction

`element_text()` yields `w:t` text plus the whitespace that tabs (`\t`), breaks (`\n`) and non-breaking hyphens render as, in document order. Without that, `PART 1<w:tab/>GENERAL` extracts as `PART 1GENERAL` and both `\s+` patterns and the `^`-anchored preserve patterns fail. `iter_text_nodes()` yields `(element, text)` pairs so character offsets map back onto the nodes they came from, which is what inline redaction relies on.

`w:delText` (tracked deletions) and `w:instrText` (field codes) are deliberately not extracted.

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

### Automated Tests

```bash
python -m unittest discover -s tests -t .
```

Stdlib `unittest`, no new dependency. `tests/docx_builder.py` assembles synthetic `.docx` files with `zipfile`, so structural cases — a sole paragraph in a footer, a cell that would stop ending with a paragraph, a `w:sectPr` paragraph, an unbalanced field — are covered without binary fixtures. `tests/support.py` provides `DocxTestCase` with `build()`, `clean()`, `detect()`, and parsed-XML assertions. GUI tests skip where `tkinter` is unavailable.

Add a test whenever you touch removal safety, the pattern tiers, or verification classification.

### Manual Testing Workflow

Run the GUI against sample files. The log output shows:
- Per-file removed/preserved counts
- Before/after file sizes
- Verification results (expected, unexpected, preserve violations)

### When Making Changes

1. Run `python -m unittest discover -s tests -t .`
2. Run the GUI against representative DOCX files (with and without footnotes/headers)
3. Open the output in Word — there must be no "unreadable content" prompt
4. Check the verification output for unexpected removals, unexpected modifications, preserve violations, and structural violations

## Releases

Releases are git tags of the form `vMAJOR.MINOR.PATCH` on `master`, following
Semantic Versioning. There is no version constant in the source and no packaging
metadata — the tag is the version.

To cut one: confirm `python -m unittest discover -s tests -t .` is green on
`master`, add the version's section to `CHANGELOG.md` (newest first, with its
link reference at the bottom), commit, then tag that commit and push the tag.
Publish the GitHub release from the tag using the changelog section as its body.

Pushing the tag also starts `.github/workflows/release.yml`, which builds the
Windows assets and attaches them. It runs on `windows-latest` because PyInstaller
does not cross-compile — nothing in this repo can produce a Windows `.exe` from
Linux or macOS. The workflow also accepts a `workflow_dispatch` with a tag name,
which retries a build without moving the tag. It runs on pull requests touching
the packaging or what goes into it as well, uploading to the workflow run rather
than to a release, so a build break is found before merge.

A tag can only be built if its tree contains the packaging: a tag build checks
out that tag alone, so anything added later is simply absent. The workflow
verifies `requirements-build.txt`, both files under `packaging/` and `apppaths.py`
are present straight after checkout and stops with the reason if they are not.
`v1.0.0` therefore has no assets and is not going to get any — it also predates
`apppaths.py`, so an executable built from it would read `patterns.yaml` out of
PyInstaller's temporary extraction directory and silently discard every edit.
Assets start at the first tag cut after packaging landed.

Two assets are produced: `SpecCleanse-<version>-portable.exe` (a single windowed
executable) and `SpecCleanse-Setup-<version>.exe` (an Inno Setup installer that
installs per-user, so it needs no administrator rights). Neither is code-signed,
so SmartScreen warns on first run.

### patterns.yaml in a frozen build

PyInstaller unpacks the bundle into a temporary directory and deletes it on exit,
so the bundled `patterns.yaml` is not somewhere a user can usefully edit — but
editing it is a documented workflow. `apppaths.resolve_config_path()` handles the
difference: from source it returns the file beside the modules, unchanged; frozen,
a copy beside the `.exe` wins if present, otherwise a per-user copy under
`%APPDATA%\SpecCleanse` is used and seeded from the bundled default on first run.
Seeding is best-effort — a read-only profile falls back to the bundled copy rather
than raising, because a traceback from a windowed build goes nowhere anyone can
see. `gui.py` logs the resolved path at startup and names it in full in
configuration errors.

Anything that changes where files live at runtime belongs in `apppaths.py`, and
needs a test in `tests/test_apppaths.py` — those rules only ever execute inside a
bundle, which the suite simulates by patching `sys.frozen` and `sys._MEIPASS`.

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

Verification picks the category up automatically from `DetectionEngine.removal_patterns()`.

### Adding an Inline Placeholder

Add the pattern to `editorial_artifacts.inline_patterns`. No code change is needed: matches become `INLINE_PLACEHOLDER` detections with character spans, the processor cuts them out in place, and verification classifies the resulting modification.

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
- **`patterns.yaml`** must be in the same directory as `gui.py`, and is always read as UTF-8 — it contains `©`, `–` and `—`, which the Windows default encoding silently mangles
- **Toggle properties** (`w:i`, `w:b`, `w:vanish`) are on when present *without* `w:val`, and off when `w:val` is `0`/`false`/`off`. Use `docx_xml.is_on()`/`toggle_on()`, never a bare `find(...) is not None`
- **Toggle properties inherited through styles do not accumulate.** Along a `w:basedOn` chain they XOR: a style that repeats its base style's `<w:vanish/>` switches hidden back *off*, and Word renders that text normally. `StyleIndex.is_hidden()` implements that, and treats an explicit `w:val="0"` anywhere in the chain as off — where the spec leaves room, take the reading that keeps text
- **Tracked deletions are not all marked up the same way.** A deleted run holds `w:delText`, which no extractor reads; a deleted table *row* keeps ordinary `w:t` and records the deletion only in `w:trPr` (cells use `w:cellDel`). Accepting revisions therefore has to remove the row or cell whole, not just the marker
- **Repacking** writes to a temp file and `os.replace`s it into place, so an interrupted run cannot leave a truncated `.docx`
- **Tracked changes and comments** are only touched when `DocxProcessor(strip_revisions=True)`, which the GUI exposes as a checkbox, default off
