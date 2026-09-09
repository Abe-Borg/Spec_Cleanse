# Changelog

All notable changes to SpecCleanse are documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- Rules that were deleting real requirement text have been narrowed. Three
  copyright patterns (`may not be reproduced`, `duplication.*?prohibited`,
  `unauthorized.*?reproduction`) matched ordinary specification prose;
  reproduction and duplication language is not by itself a copyright notice, so
  the narrowed forms require either an unambiguous marker or the boilerplate
  clause a notice actually uses, with bounded gaps that keep the words adjacent.
  `select one` now requires an editorial referent — above, below, following,
  paragraph, option, article — so "Select one of the following paragraphs" is
  cleaned and "Select one of the listed manufacturers" is not. `retain or
  delete` is anchored to the start of the paragraph, so it fires when the
  paragraph *is* the instruction rather than when a requirement contains the
  marker. The SpecAgent pattern gained word boundaries.

  More editorial text survives as a result. That is the intended trade: a
  retained note costs noise in whatever reads the cleaned file, a deleted
  requirement is not recoverable from it.
- **`specifier_notes.formatting_only_removal` now defaults to `false`.** It is
  the only mechanism that removes text on no content evidence at all — no
  pattern, no style, only italic plus an editorial colour — and specification
  text is routinely red and italic where a decision is pending. The opt-in still
  works, is still labelled `formatting-only` in Preview, and
  `tools/census_formatting` reports what the setting is worth on real documents.

  **This was a judgement call, not a measurement.** No specification corpus was
  available, so Census A was never run against real files; the decision rests on
  the asymmetric cost of the two errors. Anyone with documents to hand can check
  it in one command and set the switch accordingly.

### Added

- `detection.config_notices()`, reported in the GUI log at startup. `patterns.yaml`
  beside the executable or under `%APPDATA%` is never overwritten by an update, so
  a copy made before a rule was narrowed keeps the broad version and nothing
  otherwise says so. Flags any superseded shipped rule still active, naming the
  requirement it used to delete, and flags formatting-only removal being on.
  Matched on the exact prior string, so an edited rule is left alone.

- `.github/workflows/tests.yml` runs the suite on every pull request and every
  push to `master`, in two lanes: Windows on Python 3.12, the version the
  executable is built with and the platform whose path semantics this code has
  actually been wrong about, and Linux on 3.10, the oldest version supported.
  Both lanes run the GUI tests — `actions/setup-python` ships `tkinter` on the
  Linux runners, so the skips seen in a minimal container do not happen there. The Windows lane asserts that
  `gui.py` imports before running anything, since a skipped test exits zero and
  a lane missing tkinter would otherwise be indistinguishable from a passing
  one. Previously the release workflow was the only one, and its `*.py`
  pull-request filter matches root-level files only — so a change touching just
  `tests/` or `tools/` triggered no workflow at all and could merge without its
  tests having run anywhere. The release workflow's filter stays root-level, now
  with a comment saying why: nothing under `tests/` or `tools/` goes into the
  executable.
- `tools/`, developer measurement utilities. Nothing in the application imports
  them, they only read documents, and every clean they run is a dry run. All of
  them rest on `tools/actions.py`, which reports what a build would *do* — one
  row per action, mirroring the processor's own decision order — rather than
  what the detectors found. The two are different, and the difference decides
  whether the numbers are worth anything.
  - `census_formatting` measures what turning `formatting_only_removal` off
    would cost, by cleaning each document twice and reporting the paragraphs
    that would newly survive. The proposal to flip that default has so far
    rested on an argument rather than a measurement.
  - `census_references` measures how much of a clean sits inside a bookmark
    range that some supported `REF`/`PAGEREF`/`NOTEREF` field or internal
    hyperlink actually names — the question that decides how far reference
    protection can reasonably go.
  - `corpus_compare` records the decisions a build makes and diffs two
    recordings, so a change to the patterns can be reviewed as "what did this
    decide differently" rather than "do the tests still pass".
- `tests/test_shipped_policy.py` pins what the shipped patterns do to
  representative specification prose. Five requirement sentences the current
  rules wrongly remove are carried as expected failures until the rules are
  narrowed; five unambiguous editorial positives are pinned so that narrowing
  cannot go too far.

### Fixed

- Two selected documents sharing a basename mapped to one destination when a
  common output folder was chosen, and the second clean silently replaced the
  first. Nothing on disk showed the loss: the surviving file was a valid cleaned
  document of the wrong source. Every destination is now worked out and the whole
  set checked before any file is opened, and a batch whose destinations collide —
  or one whose destination is itself a selected input — is rejected with nothing
  written. Paths are compared as the filesystem sees them, so a symlink, a
  relative spelling or a different case is recognised as the same file. Whether
  case matters is probed on the volume rather than assumed from the platform —
  `normcase` answers for the operating system, and a default macOS volume ignores
  case while POSIX conventionally does not.
- A failed verification was counted as a success. `_clean_one()` returned `True`
  after logging `FAIL`, so a file with a preserve violation landed in the
  "succeeded" total — the one line most people read. Each file now ends as
  **verified**, **needs review**, or **failed**, counted separately
  (`Done: 8 verified, 2 need review, 1 failed.`). A needs-review file is still
  written and its path is named. Verification raising is told apart from
  processing failing: when the write succeeded and only the check blew up, the log
  says so and names the unverified file.
- A tab, line break or non-breaking hyphen inside an inline placeholder was
  counted for offsets and then left behind, so
  `Provide [Verify quantity-with Owner] units.` cleaned to `Provide -units.`
  Separators covered by a redaction are now removed with it. Page and column
  breaks are the exception: they render as `\n` and so can fall inside a match,
  but they are page setup rather than content. A placeholder straddling one is
  abandoned whole and the reason reported, rather than cut around the break —
  stranding a page break mid-requirement produced a paragraph no rule could
  explain, so every such file was reported as needing review for a decision the
  cleaner made on purpose. Other placeholders in the same paragraph are still cut.
- A correct long redaction was reported as damage.
  `Provide [Verify quantity with the Owner and the AHJ prior to bid] units.`
  cleans correctly to `Provide units.`, and verification called that an
  unexplained removal plus an invented paragraph. The similarity threshold reads
  as "about half the characters survived" but for a pure deletion rejects
  anything keeping less than a third, so it failed on the redactions that worked
  best. Verification now computes from the source what the paragraph becomes when
  every authorized placeholder is cut, and an exact match settles the pairing
  before similarity is consulted.

### Added

- `batch.py`, holding destination planning, the collision rules and `FileOutcome`.
  It is free of Tk on purpose: `gui.py` imports `tkinter` at module level, so
  anything defined there cannot be tested where Tk is absent — every Linux run.
- `ProcessingResult.warnings`, for things worth reporting that did not stop the
  run, such as a page break kept inside otherwise removed text. Separate from
  `errors`, which decide `success`.
- `docx_xml.is_layout_break()` and `docx_xml.spans_cover()`.

### Known

- No real specification corpus was available when the census tools were built,
  so they have not yet been pointed at one. Any decision taken about the
  formatting-only default or the scope of reference protection must record
  whether it had census data or was a judgement call without it.
- Suite baseline at the time of writing: 190 passed, 7 skipped, 6 expected
  failures on Linux with Python 3.11 and lxml 6.0.2; on Windows the seven GUI
  skips run, so the counts there are 190 passed, 0 skipped, 6 expected
  failures. An expected failure is not a failure and an unexpected success is;
  they are tracked separately for that reason.
- An inline placeholder still excuses deleting a requirement word beside it:
  `Provide two [Verify quantity] spare filters per unit.` reduced to
  `Provide two filters per unit.` verifies as expected, because classification
  asks whether a lost fragment *contains* a pattern match rather than whether
  matches *cover* it. The case is pinned by an `unittest.expectedFailure` test, so
  the suite turns red when the behaviour changes.

## [1.1.0] - 2026-09-08

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

[1.1.0]: https://github.com/Abe-Borg/Spec_Cleanse/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/Abe-Borg/Spec_Cleanse/releases/tag/v1.0.0
