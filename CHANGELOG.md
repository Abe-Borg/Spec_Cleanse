# Changelog

All notable changes to SpecCleanse are documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- **Verification now judges against source evidence rather than a flattened list
  of patterns.** It asks what the configured policy permits losing from the
  *input* document, then checks the actual output against that; it never asks the
  cleaner what it did. `DetectionEngine.paragraph_evidence()` produces that
  evidence — a preserve reason, whole-paragraph authority, placeholder intervals,
  and per-run authority with offsets into the paragraph text — and
  `removal_patterns()`, `inline_patterns()` and `preserve_patterns()` are gone
  with the flattening they existed for.

  The thing the flattened list could not carry is **scope**, and five wrong
  verdicts followed from that. Each is now pinned by a test that builds the
  damaged output by hand rather than by running the cleaner:

  - A hidden note run beside a requirement no longer excuses losing the whole
    paragraph. It authorizes losing its own text. Whole-paragraph loss needs a
    rule that qualified against the paragraph, or permitted intervals covering
    every substantive character in it.
  - The low-confidence tier carries the formatting it requires. `Provide pumps
    and revise as required.` is text the cleaner leaves alone, and deleting it
    is now reported instead of accepted.
  - Preserve **styles** are honoured, not only preserve patterns. A paragraph
    styled `ART` is protected even when its text matches a removal rule.
  - A protected paragraph that loses a fragment is a preserve violation, judged
    on the original paragraph — previously it was merely unexplained.
  - A lost fragment must be *covered by* a permitted interval, not merely
    contain something that matches one. `[Verify quantity]` no longer vouches
    for the word `spare` beside it, and the check is positional, so an editorial
    run cannot account for an identical requirement elsewhere in the paragraph.

- An accepted tracked deletion now outranks a preserve rule. A protected
  heading inside a row the author explicitly deleted was reported as a preserve
  violation; with `strip_revisions` on, that deletion is the instruction the run
  was asked to carry out, so it is an accepted tracked deletion instead. The
  extent is validated — the marker must be on the enclosing row or cell — so a
  revision somewhere nearby is not blanket permission, and with the option off
  the authority does not exist at all.

- **Verification compares within a location, never across one.** A location is a
  package-relative part name plus, inside `footnotes.xml`, the individual note.
  Paragraphs were previously compared as one flat list spanning every part, which
  made authority transferable between them: a requirement deleted from the body
  was reported as verified because an identical *hidden* paragraph in a header
  carried authority the body's copy did not. `word/document.xml` and
  `word/glossary/document.xml` are now distinct, and one footnote can no longer
  account for another.

- Attribution runs only once every established pairing is known. A paragraph
  that survives in shortened form is paired, but its original text no longer
  appears anywhere, which also made it look missing — so another paragraph's
  removal could be attributed to it, counting it twice and leaving a genuinely
  deleted requirement unclassified. The run was reported as verified. Paragraphs
  matched inside an unchanged block are deliberately *not* reserved, because
  there the differ matched on text alone, which for repeated text says nothing
  about which paragraph is which.

- Which of several identical paragraphs disappeared is decided by evidence rather
  than by the difference algorithm's alignment. Removing a hidden note beside an
  identical plain requirement was blamed on the plain one, reporting a correct
  clean as damage; the multiset of paragraph signatures now says which paragraph
  the output is genuinely missing. The signature includes the paragraph style,
  because two paragraphs can hold identical runs and still differ in what policy
  permits — an editorial paragraph style is authority the runs do not carry.

- The expected transformation is the whole intended removal, not one mechanism at
  a time. A paragraph carrying both a run policy permits losing and a placeholder
  inside another run had no expectation to match, so it fell through to the
  differ, which blamed the wrong copy of a repeated run. Three such correct cleans
  were reported as damage and now verify. Matching text alone is still never
  enough: the surviving runs must agree before the intervals are taken as known.

- Repeated identical text no longer makes a correct clean look like damage. A
  hidden note followed by an identical visible requirement extracts the same
  characters twice; removing the note is correct, but the difference algorithm
  assigns the deletion to the first occurrence and left the second unexplained.
  Verification now compares the output's own runs against the runs a correct
  clean would leave, which is the only evidence that separates this from the
  case where the *visible* requirement was lost and the hidden copy survived —
  the two produce byte-identical text. The comparison only ever accepts; where
  it does not match, the positional check runs unchanged.

- Disabling a detector now removes its permission as well as its removals.
  Turning hidden-text detection off means hidden formatting no longer excuses a
  loss in the output.

- Verification reports the detector's own category for a removal — `hidden_text`
  where hidden text explains it — rather than folding everything formatting into
  `formatting_based`. That label is now reserved for what it says: a removal that
  crossed the threshold on formatting alone.


- Rules that were deleting real requirement text have been narrowed. The
  copyright section now fires only on unambiguous markers — ©, "copyright",
  "all rights reserved", the ARCOM distribution line, and an anchored "licensed
  for use by". Three patterns that keyed on reproduction or duplication
  language (`may not be reproduced`, `duplication.*?prohibited`,
  `unauthorized.*?reproduction`) are gone rather than narrowed: that language
  says the same thing, about the same act, in a notice and in a requirement,
  and what differs — who imposes the restriction — is not in the text. Narrowed
  forms were tried and each one still took routine specification prose, such as
  "Provide keys designed to prevent unauthorized duplication." and "Shop
  Drawings may not be reproduced in whole or in part without the Architect's
  written consent." A notice carrying no marker now survives; that is one line
  of boilerplate against a deleted requirement, and the notice blocks this tool
  targets all carry a marker. `licensed for use by` was also anchored, having
  taken "Software licensed for use by the Owner shall be transferable.".
  `select one` now requires an editorial referent — above, below, following,
  paragraph, option, article — so "Select one of the following paragraphs" is
  cleaned and "Select one of the listed manufacturers" is not. `retain or
  delete` is anchored to the start of the paragraph, so it fires when the
  paragraph *is* the instruction rather than when a requirement contains the
  marker. The SpecAgent pattern gained word boundaries.

  More editorial text survives as a result. That is the intended trade: a
  retained note costs noise in whatever reads the cleaned file, a deleted
  requirement is not recoverable from it.
- **Documentation corrections.** Three statements contradicted the code and are
  fixed rather than softened: that element removal preserves tail text (it does
  not, and preserving it would protect nothing — an element tail in
  WordprocessingML is only inter-element whitespace); that `patterns.yaml` must
  sit beside `gui.py` (an installed build finds it beside the `.exe` or under
  `%APPDATA%`); and that the clean "keeps document structure intact", which is a
  guarantee no content remover can make. `verify.py`'s module description
  likewise still described the pre-source-evidence contract.

- **The test fixtures were not valid packages.** Every generated part marked the
  `w14` namespace ignorable without declaring it, which Markup Compatibility does
  not allow. Nothing in the suite noticed, and it would have mattered most when
  opening a fixture in Word as a control: a rejection would have said nothing
  about the change under test.

- **Large documents verify in a fraction of the time.** Verification of a
  2,000-paragraph document full of repeated headings went from 21 seconds to
  0.2; a 12,000-paragraph one is now under 1.5 seconds where the old code would
  have taken roughly three quarters of an hour by its own measured growth curve.
  Cleaning the same document went from 8.1 seconds to 0.5. A document of mostly
  unique requirements — the shape a real specification takes — verifies about
  three times faster.

  Three costs, all of them quadratic in the number of paragraphs. The
  paragraph-alignment search treated every occurrence of a repeated text as a
  candidate for every other; where the output is a pure deletion, that search
  is now skipped entirely for a single ordered pass. Behind it sat two more:
  the per-text signature inventory rescanned both documents for every distinct
  paragraph, and removal attribution rescanned every paragraph for every
  removal. Both are single passes now. Finally `can_delete_paragraph` listed a
  container's entire contents on every removal, and now stops as soon as the
  answer is known.

  **No verdict changed.** Every difference is still classified the same way:
  the fast path matches on the paragraph's full signature rather than its text,
  so a surviving hidden note can never stand in for a deleted requirement, and
  anything it declines — a modification, an addition, a reordering, a reshaped
  run — takes the original path unchanged. The injected-damage cases all still
  fail for their stated reasons.

- **`patterns.yaml` is checked for the mistakes that used to fail silently.** A
  Boolean option must now be a real Boolean: YAML makes `formatting_only_removal:
  'false'` a *string*, and every reader asks a plain truthiness question, so
  writing that the option should be off switched it **on** — the one path that
  removes text with no content evidence behind it. `enabled: 'false'` had the
  same shape and left a section running that the user had turned off. Both are
  refused now, with a message that says why quoting broke it.

  Editorial colours are checked too, and the three ways to get one wrong all
  went differently: `255` raised an `AttributeError` per file at detection time
  naming no section, key or line; `'#FF0000'` and `'bright red'` were accepted
  and matched nothing, forever. A leading `#` is now normalised away — it has one
  possible meaning and it is how every other tool writes a hex colour — and
  anything that is not six hexadecimal digits is refused, naming the section,
  key and index. Style names are deliberately not validated against any one
  document, since styles differ across templates.

- **Every file that needs review now says why.** The outcome carried one
  undifferentiated verdict, so a file needing a look because a cross-reference
  broke, because the configuration is unusual, or because a requirement went
  missing all read identically — three different pieces of work, reported the
  same way. Each outcome now names its categories: **ambiguous alignment**,
  **detected damage**, **configuration notice**, **reference/numbering
  warning**. They are counted separately and never pooled, and a file that is
  genuinely in more than one is reported in all of them rather than reduced to
  whichever sounded worst. The run summary carries the same split, and states
  whether a failed file nonetheless left something on disk.

  The split between the first two turns on how the verdict was reached, not on
  how bad it sounds. A preserve violation, an invented paragraph or a
  structural problem is a claim about the document. An unexplained removal or
  modification is such a claim only where the alignment behind it was exact;
  where the pairing came from the similarity fallback, what the report
  establishes is that the comparison could not follow the change. On a real
  pair that path reported the fragments `nd hangers a` and ` on drawings` —
  artefacts of where the differ happened to align, not text anyone edited out.

- **A configuration notice now makes a file need review.** Verification shares
  its patterns with the cleaner, so a pass means the output agrees with the
  rules it was given. When those rules include removal on formatting alone —
  the one path that removes text on no content evidence — or a rule still
  active that was narrowed because it deleted requirements, that agreement is
  not evidence the file can be handed on unread. With such a configuration
  every file in the run needs review, which is the intended reading; the
  category is what tells a user to go and look at their `patterns.yaml` rather
  than hunt through a document for a loss that never happened.

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

- **Numbering notices**, a category of their own on `VerificationResult`. A
  removed paragraph that took part in automatic numbering makes a run need review
  without claiming anything was lost: nothing went missing, but the numbers a
  reader sees may now read differently. Direct `w:numPr` and numbering inherited
  through the paragraph style are both recognised, including the default
  paragraph style Word applies where none is named, and `w:numId` `"0"` is
  treated as the override it is rather than a list called zero. A notice is raised only
  when the list still has surviving members, since a list whose every paragraph
  went renumbers nothing. Nothing is renumbered and no cross-reference is
  rewritten.
- `docx_xml.reference_target()`, `bookmark_names()`, `referenced_names()` and
  `numbering_id()`.

- `docx_xml.run_signature()`, `run_profile()` and `paragraph_signature()` — a run's
  identity as document fact (its text plus its raw `w:rPr` properties, no style
  resolution and no policy), the profile of a paragraph's text-carrying runs, and
  the paragraph's full identity including its `w:pStyle`. They answer whether two
  runs, or two paragraphs, are interchangeable.
- `docx_xml.note_identity()` and `paragraph_style()`.
- `docx_xml.field_instructions()` and `in_tracked_deletion()`.
- `docx_xml.layout_break_offsets()`, shared by the processor (which refuses to
  cut a placeholder across a page or column break) and by verification (which
  must not then claim the cut was permitted). The processor's private copy is
  gone.
- `tests/test_evidence.py` — cleans each case for real and checks that the source
  evidence predicted the output, so the evidence and the cleaner cannot drift
  apart unnoticed.
- The full V01–V12 injected-damage set in `tests/test_verify.py`. Seven of the
  twelve fail against the previous verifier and pass now; the other five were
  already correct and are carried as regression guards. V09 asserts that a
  *correct* clean still passes — without it the contract could be satisfied by a
  verifier that distrusts everything.


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
- `batch.py`, holding destination planning, the collision rules and `FileOutcome`.
  It is free of Tk on purpose: `gui.py` imports `tkinter` at module level, so
  anything defined there cannot be tested where Tk is absent — every Linux run.
- `ProcessingResult.warnings`, for things worth reporting that did not stop the
  run, such as a page break kept inside otherwise removed text. Separate from
  `errors`, which decide `success`.
- `docx_xml.is_layout_break()` and `docx_xml.spans_cover()`.

### Fixed

- **The verifier resolves `patterns.yaml` the same way the app does.** Asked to
  verify without an engine, it read the file beside its own module, which in a
  frozen build is PyInstaller's temporary extraction directory rather than the
  copy a user can edit — so it would have judged the output against the shipped
  defaults while the cleaner used the user's rules, and reported the difference
  as damage. It now goes through `apppaths.resolve_config_path()`. The GUI always
  supplies an engine, so no shipped run was affected. Precedence is unchanged and
  now tested: an engine wins outright, an explicit path comes next, and only with
  neither is the location resolved — lazily, so supplying an engine never seeds a
  per-user configuration file.

- **One file's failure no longer ends the batch.** `DocxProcessor.process()`
  turns its own exceptions into errors, but anything raised around it reached
  the run's outer handler and stopped it, leaving every remaining file
  unprocessed with nothing in the log to say why. A file that fails
  unexpectedly is now counted as that file's failure and the run continues. The
  loop moved from `gui.py` into `batch.py` so that this is testable where
  `tkinter` is absent, which is every Linux run.

- **A cross-reference this run broke is now reported.** A bookmark inside a
  removed paragraph went with it while the `REF` field naming it survived,
  leaving a live reference pointing at nothing — and lint, verification and the
  structural checks all reported the file clean. Reference fields (`REF`,
  `PAGEREF`, `NOTEREF`), simple or complex, and internal hyperlink anchors are
  now inventoried document-wide and compared. Only references this run broke are
  reported; one already dangling on the way in is the document's own problem, and
  a bookmark nothing points at may go silently. Each surviving consumer is named
  with its part and with the instruction or anchor that points at the target —
  one missing bookmark can break references in the body, a header and a note at
  once, and each is a separate place to repair. There is deliberately no
  tracked-deletion exemption: accepting a revision that deletes a referenced
  target is a requested deletion with an unrequested consequence. Nothing is
  repaired — no target is invented, no field retargeted, no replacement bookmark
  created — and the output is still written.

- **A table left with no rows by accepted revisions is removed.** Deleting the
  last row already worked; the `w:tbl` stayed, holding nothing. Word does not
  accept that, and neither the lint nor verification could see it. Where the
  table was its parent's only block, or a cell's last block, an empty paragraph
  takes its place — the minimum the container requires, not a repair. Only tables
  this acceptance actually emptied are removed: one that arrived rowless is the
  document's own problem, and removing it would rewrite a file that had no
  revisions to accept. A zero-row table and a zero-cell row are now both linted, so the shape is observable
  rather than silently absent from the rules.


- **A simple field no longer protects nothing.** Word writes a field two ways: as
  one `w:fldSimple` carrying its instruction in an attribute, or as a run sequence
  delimited by `w:fldChar` with the instruction in `w:instrText`. Only the second
  was recognised as a carrier, so an editorial paragraph whose one field was
  simple was deleted whole and took a live cross-reference with it. Because the
  paragraph's text was exactly what the rule asked to remove, nothing reported
  the loss. Such a paragraph is now emptied in place, keeping the instruction and
  its wrapper.

- A field nested inside another field's result is counted as its own carrier.
  Instructions were accumulated into one buffer and emitted when the nesting
  closed, which merged the two, so an output that lost the inner field's
  `w:fldChar` pair — keeping its instruction text and cached result — produced an
  identical inventory and the loss was invisible.

- Accepting a tracked **move** is no longer reported as damage. `w:moveFrom` is
  the one revision whose content reaches the text comparison as ordinary `w:t`,
  because a deleted run hides its text in `w:delText` that no extractor reads. So
  a correct revision-accepting run was reported as an unexplained removal, and any
  field inside the move as a lost carrier. The check now reads
  `REVISION_DELETE_TAGS` rather than naming `w:del` itself, and asks at paragraph
  scope whether every text-carrying run sits inside a revision whose content goes.

- Verification can see a field carrier disappear even when the text is unchanged.
  A stripped field leaves the same characters behind, so no text comparison
  notices; what is lost is the live reference. Fields are compared by instruction
  and **per part**, not as a document-wide count, so a field lost from a header is
  not answered by an identical one in the body. Instructions are normalised,
  because Word splits a complex one across `w:instrText` nodes at arbitrary
  points. A field inside a tracked deletion is exempt when revisions are being
  accepted — the source revision explains its absence.

  Keeping a wrapper is not a promise about its value: Word recalculates fields on
  refresh, so a preserved-but-emptied result may come back. What is promised is
  that the instruction is still there to recalculate from.

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

### Known

- No real specification corpus was available when the census tools were built,
  so they have not yet been pointed at one. Any decision taken about the
  formatting-only default or the scope of reference protection must record
  whether it had census data or was a judgement call without it.
- Performance figures quoted above are from one machine and one interpreter,
  on synthetic fixtures. They are review targets, not thresholds asserted in
  CI, and no real specification corpus has been measured.
- Suite baseline at the time of writing: 446 passed, 15 skipped, 0 expected
  failures on Linux with Python 3.11 and lxml 6.0.2; on Windows the fifteen GUI
  skips run, so the counts there are 446 passed and 0 skipped. Every case that
  was carried under `unittest.expectedFailure` has since been closed by the
  package its docstring named. An expected failure is not a failure and an
  unexpected success is; they are tracked separately for that reason.

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
