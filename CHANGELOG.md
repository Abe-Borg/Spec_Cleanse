# Changelog

All notable changes to SpecCleanse are documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-09-08

First tagged release. SpecCleanse strips editorial noise from `.docx` specification
documents so the remaining text is fit for LLM analysis, and verifies every clean
against its own input.

### Cleaning

- Single-pass shallow content removal across `document.xml`, headers, footers,
  footnotes, endnotes, and the glossary part.
- Five removal categories: specifier notes, copyright and proprietary boilerplate,
  hidden text, SpecAgent references, and editorial artifacts.
- Inline placeholder redaction. A placeholder embedded in a real requirement
  ("Provide two [Verify quantity with Owner] spare sprinklers per type.") is cut
  out where it stands so the surrounding sentence survives; a paragraph left with
  nothing but placeholders is removed in full.
- Confidence scoring over content evidence (text patterns, editorial styles) and
  formatting evidence (italic, editorial colour), with a 0.5 removal threshold.
  `specifier_notes.formatting_only_removal` decides whether formatting alone can
  carry a removal.
- Preserve patterns and preserve styles short-circuit removal regardless of score.
- Optional revision handling: accept tracked changes and strip comments, exposed
  as a GUI checkbox and off by default. Deleted table rows and cells are removed
  whole rather than left behind with their markup stripped.

### Document safety

- Paragraphs are emptied in place rather than deleted when deleting them would
  corrupt the file: an unbalanced field, a `w:sectPr` section boundary, embedded
  pictures, objects, fields or note references, or a parent left with no block
  content.
- Half-open bookmark and comment-range markers are relocated beside a paragraph
  before it is deleted, so ranges keep spanning the same content.
- Style resolution walks `w:basedOn` chains and treats toggle properties as
  toggles: repeated `w:vanish` along a chain XORs back to visible, and an explicit
  `w:val="0"` anywhere in the chain reads as off.
- Repacking writes to a temp file and `os.replace`s it into place, so an
  interrupted run cannot leave a truncated `.docx`.

### Verification

- Every clean run is followed by an automatic pass that diffs input against output
  and classifies each removed paragraph as expected, unexpected, or a preserve
  violation, alongside modification and structural-lint checks.
- Categories derive from `DetectionEngine.removal_patterns()`, so a new detector's
  patterns are recognised without a second registration.
- Verification can fail a run rather than only reporting.

### Interface

- Tkinter GUI with multi-file selection, optional output folder, Preview (dry run)
  and CLEAN modes, background-thread processing, live log and progress bar.
- Overwrites are confirmed before they happen, and a file held open in Word is
  reported as such instead of as an error code.

### Configuration

- All patterns, formatting signals, styles and preserve rules live in
  `patterns.yaml`, read as UTF-8. Simple pattern additions need no code change.
- Malformed configuration is reported instead of silently falling back.

### Project

- 85 stdlib `unittest` tests building synthetic `.docx` files with `zipfile`,
  so structural edge cases are covered without binary fixtures.
- Licensed under PolyForm Noncommercial 1.0.0, with Tcl/Tk notices documented
  separately from tkinter.
- Retired ZIP/XML structural optimization and unused-style removal stages; their
  source remains in `legacy/` for reference and is not imported by the running app.

[1.0.0]: https://github.com/Abe-Borg/Spec_Cleanse/releases/tag/v1.0.0
