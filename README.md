# SpecCleanse

SpecCleanse is a Python desktop app that removes editorial noise from `.docx` specification documents before LLM analysis.

## What it does

SpecCleanse performs a **single-pass content clean** and keeps document structure intact.

It removes five categories:

1. **Specifier notes / editorial comments** (`[Specifier: ...]`, note-to-specifier content)
2. **Copyright and proprietary boilerplate** (`©`, `all rights reserved`, licensing text)
3. **Hidden text** (`w:vanish` in WordprocessingML)
4. **SpecAgent references** (watermarks/attribution text)
5. **Editorial artifacts** (`[TBD]`, `retain or delete`, `<insert ...>`, etc.)

Every clean run is followed by an automatic verification pass that compares input/output text removals.

## What it does NOT touch

- Document formatting, style definitions, numbering structure
- Media/images
- XML metadata or structural optimization passes
- Header/footer structure (headers/footers are only scanned for removable content)

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
2. (Optional) choose **Output Folder...**.
3. Click **Preview** to run a dry-run detection report.
4. Click **CLEAN** to write cleaned documents (`*_cleaned.docx`).

## Detection configuration

Detection rules are configured in `patterns.yaml`.

### Editorial artifacts use two confidence tiers

- `editorial_artifacts.text_patterns`: high-confidence patterns (remove on text match alone).
- `editorial_artifacts.low_confidence_patterns`: lower-confidence prose patterns that require editorial formatting signals (style/color/italic combinations) before removal.

This reduces false positives for legitimate spec language.

## Verification

After cleaning, SpecCleanse runs `verify.py` to classify removed paragraphs as:

- **Expected removals** (matched known patterns/signals)
- **Unexpected removals** (needs review)
- **Preserve violations** (content that should never be removed)

A run is considered pass/fail based on unexpected removals and preserve violations.

## Project structure

```text
Spec_Cleanse/
├── gui.py
├── detection.py
├── processor.py
├── verify.py
├── patterns.yaml
├── legacy/
│   ├── style_cleaner.py
│   ├── deep_cleaner.py
│   └── README.md
├── README.md
└── LICENSE.md
```

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
