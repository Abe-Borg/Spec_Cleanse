# SpecCleanse

SpecCleanse is a Python desktop app that removes editorial noise from `.docx` specification documents before LLM analysis.

## What it does

SpecCleanse performs a **single-pass content clean** and keeps document structure intact.

It removes five categories:

1. **Specifier notes / editorial comments** (`[Specifier: ...]`, note-to-specifier content)
2. **Copyright and proprietary boilerplate** (`© 2026`, `all rights reserved`, licensing text)
3. **Hidden text** (`w:vanish` on the run, or inherited from a hidden style)
4. **SpecAgent references** (watermarks/attribution text)
5. **Editorial artifacts** (`retain or delete`, `select one of the following`, etc.)

It also **redacts inline placeholders**. MasterSpec writes placeholders inside real
requirements — "Provide two [Verify quantity with Owner] spare sprinklers per type."
Removing that paragraph would remove a requirement, so the placeholder is cut out
where it stands and the sentence around it survives. A paragraph left with nothing
but placeholders is removed in full.

Every clean run is followed by an automatic verification pass (see [Verification](#verification)).

## What it does NOT touch

- Document formatting, style definitions, numbering structure
- Media/images — a paragraph holding a picture is emptied of text rather than deleted
- XML metadata or structural optimization passes
- Section breaks, headers, footers, footnotes and table cells as structure
  (they are scanned for removable content, but never left empty)
- Tracked changes and comments, unless you turn that option on

## Installation

### Requirements

- Python 3.10+
- Tkinter-capable Python environment

### Setup

```bash
pip install -r requirements.txt
python gui.py
```

## Usage (GUI)

1. Click **Add Files...** and select one or more `.docx` files.
   Use **Remove Selected** to drop files from the list, or **Clear** to empty it.
2. (Optional) choose **Output Folder...**; **Open Output Folder** opens it in the
   file browser.
3. (Optional) tick **Strip comments and accept tracked changes** — see below.
4. Click **Preview** to run a dry-run detection report.
5. Click **CLEAN** to write cleaned documents (`*_cleaned.docx`).

Processing runs on a background thread with a live log and progress bar. Files that
already exist are only overwritten after you confirm, and a file left open in Word
is reported as such instead of as an error code.

### Strip comments and accept tracked changes (optional, off by default)

Word stores deleted text in `w:delText`, which no text extractor reads — so an
ordinary clean leaves it in the file along with its revision markup, where a
downstream LLM may well read it back. With this option on, SpecCleanse accepts
every tracked change (insertions keep their text, deletions go), removes comment
anchors, and deletes the comment parts of the package along with their
relationships, sidecar `.rels`, and content-type entries.

Deleted table rows and cells are a special case worth knowing about: their text is
not marked up at all — it stays ordinary text, and only a marker in the row's
properties records the deletion — so they are removed whole.

A tracked deletion of a *paragraph mark* is left alone: merging the two paragraphs
would move content you never asked to move.

## Detection configuration

Detection rules are configured in `patterns.yaml`. The file is UTF-8 and contains
`©`, `–` and `—`; it is always read as UTF-8, so keep it that way when editing.
Every regex in it is compiled when the app builds its engine, so a bad pattern is
reported once — with its section and index — instead of failing every document.

### Editorial artifacts use three tiers

- `editorial_artifacts.text_patterns`: high-confidence patterns. The whole paragraph
  goes on a text match alone.
- `editorial_artifacts.low_confidence_patterns`: prose patterns that could plausibly
  be real content. They start below the removal threshold and need an editorial
  formatting signal (style, colour, italic) to cross it.
- `editorial_artifacts.inline_patterns`: placeholders that sit inside real
  requirement text. They are cut out where they stand; the paragraph stays.

### Confidence scoring

| Signal | Contribution |
|--------|--------------|
| Text pattern match (specifier notes) | 0.6 |
| Text pattern match (copyright) | 0.7 |
| Text pattern match (high-confidence editorial) | 0.8 |
| Low-confidence editorial pattern | 0.3 |
| Editorial paragraph or character style | 0.8 |
| Italic | 0.2 |
| Editorial colour (red, dark red, blue, light blue) | 0.3 |
| Hidden text, SpecAgent, preserve | 1.0 |

Content is removed at 0.5 and above. Preserve patterns and preserve styles
short-circuit removal regardless of confidence.

### Removing on formatting alone

Italic plus an editorial colour adds up to exactly 0.5, so text formatted that way
is removed even when no pattern matches it. Some firms mark their notes only that
way; other firms' real specification text is red and italic. The behaviour is the
`specifier_notes.formatting_only_removal` switch, on by default. Such removals are
labelled `formatting-only` in the Preview report, and verification trusts that
signal only while the switch is on.

### Style-based detection

Style names in `style_based_detection` are matched against the style **ID** and
against the **display name** in `word/styles.xml`, and a style that inherits from a
listed style through `w:basedOn` matches too. So `CMT`, `Specifier Note`, and a
firm's own style derived from either are all recognised. A style that declares
`w:vanish` marks its text hidden — which is how MasterSpec hides its notes — unless
a run un-hides itself with `<w:vanish w:val="0"/>`.

Hidden is a *toggle* property, so declarations along a `w:basedOn` chain cancel
rather than accumulate: a style that repeats its base style's `w:vanish` is
rendered visible by Word, and SpecCleanse leaves that text alone.

`style_based_detection.preserve_styles` protects headings whose text alone gives
nothing away: MasterSpec numbers parts automatically, so the heading "PART 1 -
GENERAL" extracts as just "GENERAL". A paragraph in one of those styles is
protected whole. The requirement styles (`PR1`–`PR5`) are deliberately **not**
listed — they carry the body of the specification, which is exactly where the
placeholders that need redacting live.

## Verification

After cleaning, SpecCleanse compares the input and output and reports:

- **Removed paragraphs** — classified as expected (matched a rule, an editorial
  formatting signal, or an accepted tracked deletion), unexpected (nothing accounts
  for the loss), or a preserve violation (content that should never be removed).
  An inline placeholder inside a paragraph never explains losing the whole
  paragraph; only a paragraph that is nothing but placeholders does.
- **Modified paragraphs** — a paragraph that survived but lost text, with each lost
  fragment classified the same way. Inline redactions land here. A change that is
  not a pure deletion is never expected: if the text was altered rather than
  trimmed, it is reported.
- **Structural violations** — an emptied header, footer, footnote, text box or table
  cell; a cell that no longer ends with a paragraph; an unbalanced field; a lost
  section break. The input is inspected too, so a document's own pre-existing
  oddities are not blamed on the clean.
- **Added paragraphs** — text in the output that was not in the input.

A run passes only when all four come back clean. The verdict is advisory: a FAIL
is written to the log with its supporting detail, but the cleaned file is still
produced and still counts as processed, so read the log before trusting an output.

Verification borrows the detection engine's compiled patterns rather than
recompiling its own copy, so the two can never drift apart. That makes the pattern
layer a consistency check rather than an independent one — a rule that removes the
wrong thing will be reported as expected. The independent checks are the formatting
signals read from the source document and the structural inspection.

## Testing

The test suite is stdlib `unittest` and builds synthetic `.docx` files in memory,
so it needs nothing beyond the runtime dependencies:

```bash
python -m unittest discover -s tests -t .
```

The GUI tests are skipped where `tkinter` is unavailable.

## Project structure

```text
Spec_Cleanse/
├── gui.py              # Tkinter interface and per-file workers
├── detection.py        # Pattern/format/style detectors and the engine
├── processor.py        # Unpack, remove, redact, repack
├── verify.py           # Input/output comparison and structural lint
├── docx_xml.py         # Shared WordprocessingML plumbing and config loading
├── patterns.yaml       # Detection patterns, styles, preserve rules
├── requirements.txt
├── tests/
│   ├── docx_builder.py # Synthetic .docx fixtures
│   ├── support.py      # Shared test-case base
│   └── test_*.py
├── legacy/
│   ├── style_cleaner.py
│   ├── deep_cleaner.py
│   └── README.md
├── README.md
├── CHANGELOG.md
└── LICENSE.md
```

## Releases

Tagged releases follow [Semantic Versioning](https://semver.org/). See
[CHANGELOG.md](./CHANGELOG.md) for what changed in each one, and the
[releases page](https://github.com/Abe-Borg/Spec_Cleanse/releases) for
downloadable source archives.

## License

SpecCleanse is **source available**, not open source, under the
[PolyForm Noncommercial License 1.0.0](./LICENSE.md).

You may use, modify, fork, and redistribute SpecCleanse **for any noncommercial
purpose** — personal projects, study, research, experimentation, and use by
charitable organizations, educational institutions, public research bodies, and
government institutions.

**Commercial use requires a separate license from the copyright holder.** This
includes use by architecture, engineering, and construction firms on billable or
client work. To request a commercial license, open an issue on this repository.

### Third-party components

SpecCleanse depends on the following, all under permissive licenses. `lxml` and
`PyYAML` are installed via `pip` rather than vendored, so no additional notices
ship with this source distribution.

| Component | License |
|-----------|---------|
| [lxml](https://lxml.de/) | BSD-3-Clause (bundles libxml2/libxslt, MIT) |
| [PyYAML](https://pyyaml.org/) | MIT |
| tkinter (Python standard library) | PSF-2.0 |
| [Tcl and Tk](https://www.tcl.tk/) (runtimes behind `tkinter`) | TCL/TK License (BSD-style) |

If you build a standalone binary (e.g. PyInstaller), it bundles all of the above
and you must include their license texts in your distribution. Note that the Tcl
and Tk runtimes are separately copyrighted from the PSF-licensed `tkinter`
wrapper, and their license requires its notice be reproduced **verbatim** in any
distribution.
