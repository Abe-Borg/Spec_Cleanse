# SpecCleanse

SpecCleanse is a Python desktop app that removes editorial noise from `.docx` specification documents before LLM analysis.

## What it does

SpecCleanse offers two workflows, both of which keep document structure intact:

- **CLEAN** — single-pass automatic removal of every detected category (below).
- **Inspect Notes** — list every specifier note with its location, tick the ones you want gone, and the tool writes a new `.docx` with only those notes removed. Word does not need to be open (and on Windows must NOT be open on the input file).

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
3. Pick a workflow:
   - **Preview** — dry-run detection report (does not write any files).
   - **Inspect Notes** — opens a review window listing every specifier note with its location (e.g. *"Main body, ¶47 — under PART 2 - PRODUCTS"*) and a checkbox. Tick the notes to delete, then click **Delete Selected** to write a new `.docx`. Inspect Notes works on one file at a time.
   - **CLEAN** — full single-pass cleanup that removes every detected category and writes `*_cleaned.docx`.

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
├── notes.py
├── verify.py
├── patterns.yaml
├── legacy/
│   ├── style_cleaner.py
│   ├── deep_cleaner.py
│   └── README.md
└── README.md
```
