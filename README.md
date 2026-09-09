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

### Windows: installer or portable executable

Download either from the [latest release](https://github.com/Abe-Borg/Spec_Cleanse/releases/latest):

| Asset | Use it when |
|-------|-------------|
| `SpecCleanse-Setup-<version>.exe` | You want it installed, with a Start Menu entry and an uninstaller |
| `SpecCleanse-<version>-portable.exe` | You want to run it from a folder or a network share with nothing installed |

The installer is per-user: it writes to `%LOCALAPPDATA%\Programs\SpecCleanse`,
needs no administrator rights, and asks for no elevation prompt. Neither build
requires Python — everything is bundled.

Both are unsigned, so Windows SmartScreen will warn the first time you run one.
Choose **More info → Run anyway**, or have whoever manages your machines
whitelist it.

### Running from source

```bash
pip install -r requirements.txt
python gui.py
```

Requires Python 3.10+ in a Tkinter-capable environment.

### Where patterns.yaml lives

Editing `patterns.yaml` is the supported way to add detection patterns, so an
installed build keeps an editable copy outside the executable:

| How you run it | File it reads |
|----------------|---------------|
| From source | `patterns.yaml` beside the modules |
| Installed or portable | `patterns.yaml` beside the `.exe`, if you put one there |
| Installed or portable, otherwise | `%APPDATA%\SpecCleanse\patterns.yaml`, created from the shipped defaults on first run |

The app logs the file it loaded as its first line — `Patterns: <path>` — so
there is no guessing about which copy is in effect.

Dropping a `patterns.yaml` next to the executable overrides the per-user copy,
which is the simpler arrangement when a team shares one tuned pattern set from a
network folder.

## Usage (GUI)

1. Click **Add Files...** and select one or more `.docx` files.
   Use **Remove Selected** to drop files from the list, or **Clear** to empty it.
2. (Optional) choose **Output Folder...**; **Open Output Folder** opens it in the
   file browser.
3. (Optional) tick **Strip comments and accept tracked changes** — see below.
4. Click **Preview** to run a dry-run detection report.
5. Click **CLEAN** to write cleaned documents (`*_cleaned.docx`).

Processing runs on a background thread with a live log and progress bar. A file left
open in Word is reported as such instead of as an error code.

Before anything is written, SpecCleanse works out every destination and checks the
whole set. Two conflicts stop the run:

- **Two selected files would be written to one destination.** Two documents named
  `230500 Fire Suppression.docx` in different project folders both map to
  `230500 Fire Suppression_cleaned.docx` when you choose a common output folder, and
  the second clean would silently replace the first. Nothing on disk would show the
  loss — the surviving file is a perfectly valid cleaned document, of the wrong
  source.
- **A destination is itself one of the selected files.** Processing is sequential, so
  that input would be destroyed before its turn came.

Either one rejects the batch with nothing written, and the log names the files
involved. Choose a different output folder, or clean them in separate runs.

Paths are compared as the filesystem sees them, so two spellings of one file — a
symlink, a relative path, a different case — are recognised as the same file.
Whether case matters is asked of the volume rather than assumed from the operating
system: a default macOS volume ignores case even though POSIX conventionally does
not, and a Windows volume can be case-sensitive per directory. The check is
read-only and writes nothing.

Cleaned files that already exist from an earlier run are still only overwritten
after you confirm, once the batch's own destinations are known to be distinct.

### What the summary means

Each file ends in one of three states, and the run's summary counts them
separately — `Done: 8 verified, 2 need review, 1 failed.`

| Outcome | Meaning |
|---|---|
| **Verified** | Written, and every difference between input and output was accounted for. |
| **Needs review** | Written, but verification did not pass. The file is still produced and its path is named in the log; read the log before using it. |
| **Failed** | Processing failed, or verification could not run at all. If the file was written before the check failed, the log says so and names it as unverified. |

A successful write and a passing verification are separate facts. There is
deliberately no "succeeded" total, because that word used to cover both.

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
could be removed even when no pattern matches it. That is the
`specifier_notes.formatting_only_removal` switch, and it is **off by default**.

It is the only mechanism here that removes text on no content evidence at all — no
pattern, no style, only how the text looks. Some firms do mark their notes only that
way, but specification text is routinely red and italic where a decision is still
pending, and the two are indistinguishable to this tool. A retained note costs some
noise in whatever reads the cleaned file; a deleted requirement is not recoverable
from it.

Turn it on in `patterns.yaml` if your notes carry no other distinguishing mark.
Removals it causes are labelled `formatting-only` in the Preview report, and
verification trusts that signal only while the switch is on. To see what the setting
is worth on your own documents before deciding:

```bash
python -m tools.census_formatting /path/to/specs
```

That cleans each file twice — once each way — and reports the paragraphs the setting
accounts for, without writing anything.

### If you are upgrading

`patterns.yaml` beside the executable or under `%APPDATA%\SpecCleanse` is never
overwritten by an update, so a copy made earlier keeps whatever rules it had. On
startup the log names the file in use and flags two things about it: any shipped
rule still present that was later narrowed because it deleted real requirement
text, and formatting-only removal being switched on. Neither stops a run. Your
edits are never touched — compare your file with the shipped one and take what you
want.

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

- **Removed paragraphs** — classified as expected (a rule qualified against the
  whole paragraph, the permitted intervals accounted for every substantive
  character in it, or a tracked deletion was accepted), unexpected (nothing
  accounts for the loss), or a preserve violation (content that should never be
  removed, by pattern *or* by style).

  One editorial signal *inside* a paragraph is not permission to lose the
  paragraph. A hidden note run beside a requirement authorizes losing its own
  text and nothing beside it; an inline placeholder authorizes cutting the
  placeholder. Either explains a whole-paragraph loss only when what it
  authorizes covers everything the paragraph said.
- **Modified paragraphs** — a paragraph that survived but lost text, with each lost
  interval classified the same way. Inline redactions land here. A change that is
  not a pure deletion is never expected: if the text was altered rather than
  trimmed, it is reported.

  Each lost interval has to be *covered by* a permitted one, not merely to
  contain something that matches a rule. `[Verify quantity]` sitting next to
  `spare` permits cutting the placeholder and says nothing about the word beside
  it. The check is positional for the same reason: a paragraph can hold the same
  words twice, once in an editorial run and once in a requirement, and asking
  whether a lost fragment *resembles* something removable cannot tell which
  occurrence actually went.

  When a paragraph holds the same text twice, positions alone are not enough
  either. A hidden note followed by an identical visible requirement extracts
  the same characters twice, so an output holding one copy is consistent with
  either having gone — and those two outcomes are a correct clean and a lost
  requirement, with byte-identical text. Verification therefore compares the
  output's own runs against the runs a correct clean would leave. That
  comparison only ever accepts: where it does not match, the positional check
  runs unchanged, so keeping the hidden copy while the visible requirement
  disappears is still reported.

**Comparison happens inside a location.** A location is a part of the package —
the document, a header, the glossary — plus, within `footnotes.xml`, the
individual note. Nothing is ever compared across one. Text equality is not
identity: a requirement deleted from the body must not be excused by an
identical paragraph in a header, and one footnote cannot account for another.

Where a location holds the same text more than once, which copy disappeared is
decided by comparing paragraph signatures — the style and the runs — rather than
by the difference algorithm's choice of alignment. That matters in both
directions: a correct clean that removes a hidden note beside an identical plain
requirement must pass, and an output that keeps the *hidden* copy while the
plain requirement disappears must not.

  A protected paragraph is not touched by the cleaner at all — not redacted, not
  trimmed — so any loss inside one is a violation whatever the lost text looks
  like, judged on the original paragraph where the protection is visible.

  To decide which output paragraph an input paragraph became, verification first
  computes — from the source text and the configured patterns, independently of
  anything the cleaner reports — what the paragraph becomes when every authorized
  placeholder is cut out. An exact match to that text settles the pairing. Only if
  no exact answer is found does it fall back to character similarity. That
  fallback alone used to fail on the redactions that worked best: a paragraph
  keeping under a third of its characters was rejected as too dissimilar, so
  `Provide [Verify quantity with the Owner and the AHJ prior to bid] units.`
  correctly cleaned to `Provide units.` was reported as an unexplained removal
  plus an invented paragraph.
- **Structural violations** — an emptied header, footer, footnote, text box or table
  cell; a cell that no longer ends with a paragraph; an unbalanced field; a lost
  section break. The input is inspected too, so a document's own pre-existing
  oddities are not blamed on the clean.
- **Added paragraphs** — text in the output that was not in the input.

A run passes only when all four come back clean. The verdict is advisory: a file
that does not pass is still produced, but it is counted as **needs review** rather
than as a success, and its path is named in the log.

A PASS means every difference between input and output was accounted for by a
configured rule, and the structural checks found nothing the input did not already
have. It is not a statement that the document is correct, and not a claim that Word
will open it without complaint.

Verification asks what the configured policy permits losing from the **source**
document, and then checks the actual output against that. It never asks the cleaner
what it did: a program's account of its own work cannot be the evidence that the
work was right, since a bug would report itself as intended.

It does share the detection engine's compiled rules, which is deliberate — the two
must not drift into disagreeing about what the configuration says. The limit that
follows is worth stating plainly: **a rule that removes the wrong thing is still
reported as expected.** Verification checks that the cleaner did what the rules ask,
not that the rules ask for the right thing. Whether a rule is correct is a question
for the patterns themselves, and for the fixtures that pin them.

Disabling a detector removes its permission too, not just its removals: turning
hidden-text detection off means hidden formatting no longer excuses a loss.

## Testing

The test suite is stdlib `unittest` and builds synthetic `.docx` files in memory,
so it needs nothing beyond the runtime dependencies:

```bash
python -m unittest discover -s tests -t .
```

The GUI tests skip themselves where `tkinter` is unavailable, which is true of a
minimal Linux install though not of the CI runners. The rules that decide whether a
batch is safe to write live in `batch.py` rather than `gui.py` for that reason, so
they are exercised wherever the suite runs at all.

Two suites carry the verification contract, and they ask opposite questions.
`tests/test_evidence.py` cleans each case for real and checks that what the source
evidence predicted is what the output shows — if the evidence and the cleaner ever
disagree, verification has quietly become a second opinion about a different
program. `tests/test_verify.py` builds damaged outputs **by hand, never by running
the cleaner**, because agreement between a broken verifier and the cleaner that
produced its input proves nothing. One of those cases asserts that a *correct* clean
still passes, which is what stops the whole contract being satisfied by a verifier
that simply distrusts everything.

Continuous integration runs the suite on every pull request and every push to
`master`, in two lanes: Windows on the Python version the executable is built with,
and Linux on 3.10, the oldest version this project claims to support. The Windows
lane checks that `gui.py` imports before running anything, because a skip exits zero
and a lane that skipped the GUI tests would otherwise look exactly like one that
passed them.

Some tests are carried as `unittest.expectedFailure` — cases that describe
behaviour a later change will fix. The suite stays green while they fail, and turns
red if one ever starts passing, which is what prompts removing the decorator along
with the defect. Today they cover an inline placeholder excusing the deletion of a
requirement word beside it, and five requirement sentences the shipped patterns
currently remove.

## Developer tools

Measurement utilities under `tools/`. Nothing in the application imports them, they
only read documents, and every clean they run is a dry run. They exist to answer
questions that decisions about the patterns depend on.

```bash
# What would turning formatting-only removal off actually cost?
python -m tools.census_formatting /path/to/specs

# How much of a clean sits inside a bookmark range some REF field points at?
python -m tools.census_references /path/to/specs

# Record what this build decides, change something, and diff the two.
python -m tools.corpus_compare record baseline.json /path/to/specs
python -m tools.corpus_compare diff baseline.json candidate.json
```

Recordings and census output carry excerpts of the documents they measured.
Specifications are usually proprietary — keep them in a scratch directory and do not
commit one. `corpus_compare --no-text` omits excerpts and leaves a short content
digest in their place, which is what keeps two recordings comparable. The digest is
a fingerprint, not encryption: it stops a recording being readable, which is what
committing one would leak.

All three read what the processor would actually *do*, not what the detectors found.
Those differ — a note matching two rules is two detections and one removed
paragraph — and the numbers are only worth taking a decision on because of it.

## Project structure

```text
Spec_Cleanse/
├── gui.py              # Tkinter interface and per-file workers
├── batch.py            # Destination planning, collision rules, file outcomes
├── detection.py        # Pattern/format/style detectors and the engine
├── processor.py        # Unpack, remove, redact, repack
├── verify.py           # Input/output comparison and structural lint
├── docx_xml.py         # Shared WordprocessingML plumbing and config loading
├── apppaths.py         # Where patterns.yaml lives, source vs. frozen build
├── patterns.yaml       # Detection patterns, styles, preserve rules
├── tools/              # Developer measurement utilities (not imported by the app)
│   ├── census_formatting.py   # Cost of turning formatting-only removal off
│   ├── census_references.py   # Removals inside referenced bookmark ranges
│   └── corpus_compare.py      # Record decisions, diff two builds
├── requirements.txt
├── requirements-build.txt
├── tests/
│   ├── docx_builder.py # Synthetic .docx fixtures
│   ├── support.py      # Shared test-case base
│   └── test_*.py
├── legacy/
│   ├── style_cleaner.py
│   ├── deep_cleaner.py
│   └── README.md
├── packaging/
│   ├── speccleanse.spec  # PyInstaller build
│   └── installer.iss     # Inno Setup installer
├── .github/workflows/
│   └── release.yml     # Builds and attaches release assets on a tag
├── README.md
├── CHANGELOG.md
└── LICENSE.md
```

## Building the Windows executable

PyInstaller does not cross-compile, so a Windows build has to happen on Windows.
Pushing a `vX.Y.Z` tag runs `.github/workflows/release.yml`, which builds both
assets on a Windows runner and attaches them to the release; the same workflow
can be re-run manually against an existing tag from the Actions tab.

Only tags that contain the packaging can be built. `v1.0.0` predates it and has
no assets: it also predates the fix for `patterns.yaml` being read out of
PyInstaller's temporary extraction directory, so an executable built from that
tag would ship a build whose pattern edits silently do nothing. The workflow
checks for the build tooling after checkout and stops with that explanation
rather than failing obscurely later.

The same workflow also runs on pull requests that touch the packaging or what
goes into it. Those builds attach the two executables to the workflow run instead
of to a release, so a broken build shows up before merge and the result can be
downloaded and tried from the PR's Checks tab.

To build locally on Windows:

```bat
pip install -r requirements.txt -r requirements-build.txt
pyinstaller packaging\speccleanse.spec --noconfirm
iscc /DAppVersion=1.0.0 packaging\installer.iss
```

Both land in `dist\`. The installer step needs [Inno Setup 6](https://jrsoftware.org/isdl.php);
skip it if you only want the portable executable.

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
