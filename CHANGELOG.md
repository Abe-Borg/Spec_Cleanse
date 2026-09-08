# Changelog

All notable changes to SpecCleanse are documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Windows packaging. Each release now carries two assets built on a Windows
  runner: `SpecCleanse-Setup-<version>.exe`, an Inno Setup installer that
  installs per-user and so needs no administrator rights, and
  `SpecCleanse-<version>-portable.exe`, a single windowed executable that runs
  from a folder or network share. Neither requires Python. Neither is signed, so
  SmartScreen warns on first run.
- `.github/workflows/release.yml` builds and attaches those assets when a
  `vX.Y.Z` tag is pushed, builds them on pull requests that touch the packaging
  so a break is caught before merge, and can be re-run manually against an
  existing tag. Only tags containing the packaging can be built; `v1.0.0`
  predates both it and the `patterns.yaml` fix below, so it has no assets.
- `apppaths.py`, holding the runtime file locations, with tests that simulate a
  PyInstaller bundle.

### Fixed

- `patterns.yaml` was located as `Path(__file__).parent / "patterns.yaml"`, which
  in a frozen build resolves inside PyInstaller's temporary extraction directory
  — deleted when the process exits. Editing it, the documented way to add
  detection patterns without touching code, would have silently done nothing for
  anyone running an installed build. A frozen SpecCleanse now prefers a copy
  beside the executable, otherwise a per-user copy under `%APPDATA%\SpecCleanse`
  seeded from the shipped defaults on first run.
- Configuration errors named only `patterns.yaml`; they now give the full path,
  since a frozen build has more than one copy. The GUI also logs the file it
  loaded as its first line.

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
- A run is judged PASS only when unexpected removals, unexpected modifications,
  preserve violations, structural violations and added paragraphs all come back
  empty; the verdict and its supporting detail are written to the log. The verdict
  is advisory — a FAIL is reported but the cleaned file is still written and still
  counts as processed, so review the log before trusting an output.

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

[Unreleased]: https://github.com/Abe-Borg/Spec_Cleanse/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/Abe-Borg/Spec_Cleanse/releases/tag/v1.0.0
