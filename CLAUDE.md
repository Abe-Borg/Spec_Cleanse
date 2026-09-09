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
| `batch.py` | Destination planning and collision rules, the run loop, and the outcome vocabulary — `FileOutcome`, `ReviewCategory`, `FileReport`, `RunTally`, `run_batch`. Deliberately free of Tk *and* of every other project module, so the rules are testable where `tkinter` is absent. Case folding is decided per *volume*, not per platform — `normcase` answers the wrong question, since a default macOS APFS volume ignores case while `posixpath.normcase` is the identity |
| `detection.py` | Pattern matching engine with confidence scoring; all detector classes; and `paragraph_evidence()`, which states what policy permits losing from a source paragraph |
| `processor.py` | DOCX unpacking/repacking, XML walking, element removal, inline redaction |
| `verify.py` | Post-processing verification: comparison and classification against the source evidence — removals, modifications, structural lint |
| `docx_xml.py` | Shared WordprocessingML plumbing: namespaces, iteration, text extraction, structure rules, style resolution, config loading |
| `apppaths.py` | Runtime file locations: which `patterns.yaml` to load from source vs. a frozen build |
| `tests/` | stdlib `unittest` suite; builds synthetic DOCX files with `zipfile` |
| `tools/` | Developer measurement utilities. Never imported by the application, read-only, dry runs. `actions` (what a build would *do*, one row per action — the base the others rest on), `census_formatting` (what turning formatting-only removal off would cost), `census_references` (removals inside referenced bookmark ranges), `corpus_compare` (record decisions, diff two builds) |
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

Verification derives its categories from `DetectionEngine.paragraph_evidence()`, so a
new detector's decisions are recognised there without a second registration.

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
0.5 threshold. `specifier_notes.formatting_only_removal` (**default false**) decides
whether formatting alone is enough; when it is false, formatting only boosts a score
that content evidence already opened. Removals that crossed on formatting alone carry
`Detection.formatting_only` and are labelled in Preview.

The default is off because this is the only path that removes text on no content
evidence whatsoever, and real specification text is routinely red and italic. It is
set in three places that must agree — the shipped YAML, `PatternConfig` and
`_make_pattern_config`'s omitted-key default, and verification's omitted-key default
in `verify_clean` — and a test asserts each. Tests that depend on the switch state it
explicitly rather than leaning on whichever way the default points.

`detection.config_notices()` reports what a user's own `patterns.yaml` is doing that
they cannot otherwise see: a shipped rule still present that was later narrowed for
deleting requirements, and formatting-only removal being on. `apppaths` prefers an
existing user file over the bundled one, so an update never changes their rules.
Superseded rules are matched on the **exact** prior string, so someone who edited a
rule themselves is not told their own work is stale.

`editorial_artifacts` uses three tiers: `text_patterns` are high-confidence and remove
the whole paragraph on text alone; `low_confidence_patterns` start at 0.3 and require a
formatting signal to cross the threshold; `inline_patterns` produce
`ContentType.INLINE_PLACEHOLDER` detections carrying character `spans`, which are cut
out of the paragraph in place. `DetectionEngine.should_remove()` ignores inline
detections — they never remove their element.

The inline tier must stay separable when *judging* a removal. An inline pattern
matching a paragraph that vanished entirely means only that the paragraph contained a
placeholder — never that losing it was intended. So inline matches enter the evidence
as *intervals*, never as whole-paragraph authority, and a paragraph is explained by
them only when they cover all of its substantive text.

### Pairing an Input Paragraph With Its Survivor

`_pair_with_survivor()` takes exact answers before guessing. An unchanged paragraph
pairs outright; so does one matching `ParagraphInfo.expected_after_redaction()`, which
computes from the source text and the configured patterns — never from anything the
processor reports — what the paragraph becomes when every authorized placeholder is cut.

Only when no exact answer exists does `MIN_PAIR_SIMILARITY` apply. That threshold
reads as "about half the characters survived" and is not: for a pure deletion the
ratio is `2*len(after)/(len(before)+len(after))`, which falls below 0.5 once *more
than two-thirds* of the characters go. It therefore rejected exactly the redactions
that worked best, which is the bug the exact path fixes.

Recognising the exact result does not widen what counts as permitted — anything
other than that text still faces the similarity path unchanged, and whatever it lost
still has to be covered by an authorized interval before it counts as expected.

### The Verification Contract: Source Evidence

Verification asks what the *configured policy* permits losing from the *source*
document, then checks the actual output against that. It never asks the processor
what it did. The cleaner's account of its own work cannot be the evidence that the
work was right — a bug in the processor would report itself as intended.

`DetectionEngine.paragraph_evidence(para)` produces that evidence for one source
paragraph, in the processor's decision order: preserve, then a whole-paragraph rule,
then placeholder intervals, then individual runs. It is a pure read — no mutation,
no processor call.

The distinction it carries, and that a flattened list of `(category, regex)` pairs
could not, is **scope**:

| Evidence | Authorizes |
|---|---|
| `preserve_reason` | nothing — the paragraph is protected outright, by pattern *or* by style |
| an accepted tracked deletion | losing the paragraph, outranking even protection — see below |
| `whole_category` | losing the entire paragraph, because a rule qualified against the paragraph |
| `runs[i]` with a category | losing *that run's own characters*, and nothing beside them |
| `inline_spans` | losing those intervals |

Three consequences follow, and each closed a real defect:

- **One editorial signal inside a paragraph is not permission to lose the
  paragraph.** A hidden note run beside a requirement authorizes its own text.
  Whole-paragraph loss needs either a qualifying whole-paragraph rule or
  `covers_all_text()` — the authorized intervals accounting for every substantive
  character.
- **The low-confidence tier carries its formatting prerequisite.** It arrived at
  verification as a bare regex, so `Provide pumps and revise as required.` — which
  the cleaner leaves alone, because 0.3 does not reach the threshold — was reported
  as an expected removal when something else deleted it.
- **Authority is positional, not textual.** `authority_for(start, end)` asks which
  rule covers an interval. Asking whether a lost fragment *resembles* something
  removable cannot tell two identical occurrences apart, and a paragraph may hold
  the same words in a hidden run and in a requirement.

Coverage is also what `_classify_modification` requires: every lost interval must be
**covered by** an authorized one, not merely *contain* something that matches a rule.
`[Verify quantity]` authorizes cutting the placeholder and says nothing about the
word `spare` beside it.

**Repeated identical text breaks the alignment the intervals rest on, so one exact
answer comes first.** A hidden note followed by an identical visible requirement
extracts the same characters twice. When the cleaner removes the note, the survivor is
equally consistent with *either* occurrence having gone, and `SequenceMatcher` simply
picks the first — leaving the second interval unauthorized and a correct clean reported
as damage.

No rule reading text alone can fix this, because the correct clean and the
corresponding damage produce **byte-identical text**; the difference is only which run
survived. So `ParagraphEvidence.surviving_signature()` states the runs a correct clean
would leave behind, and `docx_xml.paragraph_signature()` reads what the output actually
has. Signatures are pure document fact — text plus raw `w:rPr` properties, no style
resolution and no policy — because the question is narrow and syntactic: *are these two
runs interchangeable?*

Reading the output's structure is not a breach of the boundary above. What is forbidden
is trusting the processor's account of its actions; the output package is the artifact
being judged, and §10.3 requires establishing whether it "retains all protected content
in order".

**That path only ever accepts.** A mismatch falls through to the interval reasoning
unchanged, so an output that kept the *hidden* copy while the visible requirement
vanished is still reported. A paragraph whose runs were legitimately reshaped — an
inline redaction rewrites run text — simply misses the fast path rather than being newly
flagged.

A protected paragraph is not touched by the cleaner at all — not redacted, not
trimmed — so any loss inside one is a violation whatever the lost text looks like,
judged on the original paragraph where the protection is visible.

**One thing outranks protection: an explicit tracked deletion, and only while
`strip_revisions` is on.** A preserved heading inside a row the author deleted is
that deletion working, not damage — the run was asked to accept revisions. The
extent is validated rather than assumed: `_in_tracked_deletion` requires the marker
on the enclosing `w:tr` or `w:tc`, so a revision somewhere nearby is not blanket
permission. With the option off the authority does not exist, and the same loss is
a violation again.

### Location, and Which Occurrence Went

**Comparison happens inside a location, never across one.** A location is a
package-relative part name plus, inside a part holding several independent stories,
the note identity from `docx_xml.note_identity()`. `word/document.xml` and
`word/glossary/document.xml` are different locations; so are footnote 3 and footnote 7
within the one `footnotes.xml`.

A flat list across every part made authority transferable between them: a requirement
deleted from the body verified clean because an identical *hidden* paragraph in a
header carried authority the body's copy did not. Text equality is not identity.

**Within a location, repeated text still does not say which occurrence went.**
`difflib` aligns on the longest matching block, not on evidence, so removing a hidden
note and keeping the plain requirement beside it was blamed on the plain one — a
correct clean reported as damage. `_lost_signatures()` takes the multiset difference of
paragraph signatures between the two sides, and `_attribute_removal()` judges the
paragraph the output is genuinely missing rather than whichever index the alignment
left over. It only ever re-attributes among paragraphs whose text is already identical.

Pairing is settled before any attribution. A paragraph that survives in shortened
form is paired, but its original text is gone from the output, so it also looks
*missing* — and offering it as the explanation for another paragraph's loss counts it
twice and leaves the real loss unclassified, which reported a deleted requirement as a
verified clean. Paragraphs matched inside an unchanged block are deliberately not
reserved: there the differ matched on text alone, which for repeated text says nothing
about which paragraph is which.

Two signatures are needed because they answer different questions:

| | Contents | Answers |
|---|---|---|
| `run_profile(para)` | the text-carrying runs' signatures | did the runs a correct clean would leave actually survive? |
| `paragraph_signature(para)` | `(w:pStyle, run_profile)` | which of several identical paragraphs disappeared? |

The style has to be in the second. Two paragraphs can hold character-for-character
identical runs and still differ in what policy permits, because an editorial
*paragraph* style is authority the runs know nothing about.

**The expected transformation is the whole intended removal, not one mechanism.**
`ParagraphEvidence.expected_text()` cuts every authorized interval — runs *and*
placeholder spans — and `expected_profile()` is the run profile that leaves, each
surviving run carrying the text it keeps once the placeholders inside it are cut. A
paragraph holding both a hidden twin and a placeholder has no single-mechanism
expectation to match, so it used to fall through to the differ, which then blamed the
wrong copy.

**Matching text is never sufficient on its own.** Where the output equals the expected
text, the runs must agree as well before the intervals are taken as known — a paragraph
holding a visible requirement and an identical hidden copy produces the same string
whichever one was lost, and only the surviving run's properties say which. When they do
agree the intervals are the authorized spans outright; re-deriving them by diffing would
reintroduce the guess.

**Two places mirror the processor's decision order, on purpose, for different
questions.** `tools/actions.py` asks the processor's own methods what it *would do*,
because a measurement must match the build being measured. `paragraph_evidence()`
asks the policy what is *permitted*, and deliberately never consults the processor,
because that is the independence the contract rests on. `tests/test_evidence.py`
pins the second against real cleans, so the two cannot drift apart unnoticed.

Offsets are into the **raw** paragraph text. A `ParagraphInfo` keeps `raw_text` and
`text` (trimmed, what the comparison aligns on) separately, and `lead` maps between
them. An offset computed against a trimmed or normalized string addresses the wrong
characters. `offsets_reliable` is `False` if run texts ever fail to reconstruct the
paragraph, and interval reasoning is then abandoned rather than guessed at.

### Why a File Needs Review

"Needs review" on its own cannot be acted on. Reading a document for a lost
requirement, fixing a `patterns.yaml`, and checking a cross-reference are three
different jobs, and a verdict that does not say which will be ignored. Every
outcome names its categories, they are counted separately and never pooled, and
a file genuinely in more than one is reported in all of them.

| Category | Means | Comes from |
|---|---|---|
| `AMBIGUOUS_ALIGNMENT` | something is unexplained and where it came from is a guess | an unexplained finding whose pairing was the `MIN_PAIR_SIMILARITY` fallback, or whose offsets were unreliable |
| `DETECTED_DAMAGE` | a claim about the document | preserve violations, added paragraphs, structural damage, and unexplained findings reached on an *exact* alignment |
| `CONFIGURATION` | the rules that produced this file are worth knowing about | `detection.config_notices()` |
| `REFERENCE_NUMBERING` | nothing was lost; what a reader sees may differ | `VerificationResult.numbering`, and structural violations of `kind == "reference"` |

**The first two split on how the verdict was reached, not on how bad it sounds.**
Where the pairing was a guess, what the report establishes is that the comparison
could not follow the change — a real pair on that path reported the fragments
`nd hangers a` and ` on drawings`, artefacts of where the differ happened to
align rather than text anyone edited out. Calling that damage overstates it in
exactly the direction that trains a user to ignore the verdict. §10.6 wants the
two counted apart for the same reason: a Needs-review rate dominated by ambiguous
alignment is a reason to revisit the comparison, and one dominated by damage is a
different problem with a different fix.

**A category is decided by a field, never by reading the message.**
`StructuralViolation.kind` exists so that the sentence shown to the user is free
to change without silently reclassifying anything.

**Configuration is not visible to verification**, because nothing in a comparison
of two documents can see it — `review_categories()` deliberately omits it and
`_clean_one` adds it. It is decided once per run: the configuration cannot change
between two files in one batch. A notice makes *every* file in the run need
review, which is the intended reading. Verification shares its patterns with the
cleaner, so a pass means the output agrees with the rules it was given; when
those rules include removal on formatting alone, agreement is not evidence the
file can be handed on unread.

**`VerificationResult.passed` stays the authority on the comparison, with the
categories as the explanation.** Deriving the verdict from the categories is
equivalent today and would fail silently the moment something new contributes to
`passed` without a matching category — a real failure reported as Verified. A
subTest over every contributor asserts the two agree.

### Running a Batch

`batch.run_batch()` owns the loop, not `gui.py`, for the reason `batch.py` exists:
"the batch kept going after one file failed" is not a claim worth making untested,
and it was not true. `DocxProcessor.process()` turns its own exceptions into
errors, but anything raised around it ended the whole run. One file's unexpected
failure is now that file's.

It is handed the manifest `plan_batch()` validated before the worker started and
never recomputes a destination. The file selection and output folder are live
widgets; a destination worked out mid-run could collide with one already written —
the loss `plan_batch` exists to prevent, reintroduced after its check had passed.

`FileReport.output_written` is separate from the outcome because a failure can
still leave a document on disk — verification raising after a successful write is
exactly that — and a reader told only "1 failed" cannot tell whether there is
something in the output folder to delete.

### Dataclass-Based Results

Processing results are communicated via dataclasses, not exceptions:
- `Detection` — individual content detection with confidence, spans, formatting flag
- `ProcessingResult` — detections, an errors list, and a `warnings` list. Warnings
  are things worth telling the user that did not stop the run — a structure kept
  because removing it would have gone beyond removing content. They are separate
  from `errors`, which decide `success`
- `FileOutcome` — `VERIFIED` / `NEEDS_REVIEW` / `FAILED` (in `batch.py`). A
  successful write and a passing verification are separate facts and one Boolean
  cannot carry both
- `ReviewCategory` / `FileReport` / `RunTally` (in `batch.py`) — why a file needs
  review, what became of one file, and a run's running counts. `_clean_one()`
  returns a `FileReport`, never a bare outcome: "needs review" on its own cannot
  be acted on
- `BatchPlan` / `BatchItem` / `BatchConflict` — the validated input/output pairs for
  a run, built before anything is opened
- `RunEvidence` / `ParagraphEvidence` — what policy permits losing from one *source*
  paragraph, read from the source alone. `RunEvidence` carries offsets into the raw
  paragraph text; `ParagraphEvidence` carries whole-paragraph authority, placeholder
  intervals, and the preserve reason. This is the type that replaced a flattened list
  of `(category, regex)` pairs, and *scope* is the thing it carries that the list could not
- `RemovedParagraph` / `ModifiedParagraph` / `StructuralViolation` / `VerificationResult` — verification output. Categories: a detector name, `formatting_based`, `inline_placeholder`, `tracked_deletion`, `preserve_violation`, or `None` for unexplained. An unexplained finding also carries `ambiguous`, and a `StructuralViolation` carries `kind`
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

### Inline Redaction and Separators

`iter_text_nodes()` yields `w:t` **and** the separator elements that render as a
character — `w:tab`, `w:br`, `w:cr`, `w:ptab`, `w:noBreakHyphen`. `_redact_spans()`
walks that stream with a running offset, so a redaction span can land on any of
them. A `w:t` is edited; a separator has no partial state, so it is removed whole
when a span covers all of it and left alone otherwise. Removal happens *before* the
empty-run sweep, so a run holding nothing but a redacted separator is seen as empty.

Advancing the offset past a separator without removing it is what left
`Provide -units.` behind.

**Page and column breaks are the exception, and the whole redaction is abandoned
rather than half-completed.** Extraction renders every `w:br` as `\n`, which is what
lets one fall inside a match at all, but a break carrying `w:type="page"` or
`"column"` is page setup rather than content — removing it reflows the document from
that point on, a larger claim than any editorial pattern makes.

Cutting the text *around* such a break and stranding it is worse than not cutting:
it leaves a page break mid-requirement, and it produces a paragraph no rule
explains. Verification computes the expected text by cutting the whole placeholder,
the output does not match it, and the diff fallback then sees two fragments — the
text before the break and the text after — neither of which matches the placeholder
pattern alone. Every such file would be reported as needing review for a decision
the cleaner made on purpose.

So `_drop_spans_over_layout_breaks()` discards any span straddling one, before any
mutation, and records why. The placeholder survives; that is the lesser cost. Other
spans in the same paragraph are still cut. A soft break (no type, or
`textWrapping`) is whitespace and goes with the text around it.
`docx_xml.is_layout_break()` states the distinction; the policy is the processor's.
The guard inside `_redact_spans()` is unreachable by the ordinary path and kept only
as a last line of defence, because stranding a break is not recoverable.

### Field Carriers

Word records the same field two ways, and only one of them was protected.

| Form | Instruction lives in | Was protected |
|---|---|---|
| `w:fldSimple` | the `w:instr` **attribute** | no |
| complex field | `w:instrText` between `w:fldChar` begin/end | yes |

A field's cached result reads as ordinary words, so an editorial pattern matching the
paragraph around it deleted the live cross-reference with it — and because the text was
what the rule asked to remove, nothing reported the loss. `w:fldSimple` is now in
`EMBEDDED_CONTENT_TAGS`, so `has_embedded_content()` protects its paragraph and
`_remove_paragraph` empties it in place instead of deleting it.

`docx_xml.field_instructions()` reads both forms and normalises the instruction, because
Word splits a complex instruction across `w:instrText` nodes at arbitrary points — ` REF `
plus `Target ` is the same field as ` REF Target `. It returns a Counter, not a set: a
document may legitimately hold the same field twice and losing one of them is still a loss.

**Nested fields need a stack, not a depth counter.** A field inside another field's result
is its own carrier. Accumulating instructions into one buffer and emitting when the nesting
closed merged the two, so an output that lost the inner field's `w:fldChar` pair — keeping
its instruction text and cached result — produced an *identical* inventory, and the loss was
invisible.

Verification compares those instructions **per part**, not as a document-wide count. A
field lost from a header is not answered by an identical one in the body, and a total
would hide one field going while another arrives. `skip_deleted` leaves out fields inside
a tracked deletion, which a run accepting revisions removes legitimately — the source
revision is the evidence that explains their absence.

`docx_xml.in_tracked_deletion()` answers "would `accept_revisions()` remove this?", and it
reads `REVISION_DELETE_TAGS` rather than naming `w:del` itself. A second, shorter list of
that answer drifted from the first and left `w:moveFrom` out, so a field inside an accepted
*move* was reported lost from a run that had done exactly the right thing.

**`w:moveFrom` is the one revision whose content reaches the text comparison.** A deleted
run hides its text in `w:delText`, which no extractor reads, so it never arrives; the source
half of a move keeps real `w:t` until the move is accepted. `verify._in_tracked_deletion()`
therefore asks two questions — is the *container* marked deleted (a row in `w:trPr`, a cell
in `w:tcPr`), or does every text-carrying run sit inside a revision whose content goes.

**Keeping a wrapper is not a promise about its value.** Word recalculates fields on
refresh, so a preserved-but-emptied field result may come back. What is promised is that
the instruction and its wrapper are still there to recalculate from. Field instructions
are never rewritten and field values are never updated programmatically.

### Reference Integrity, Emptied Tables, and Numbering

Three consequences a clean can have that no text comparison sees.

**A reference this run broke.** A bookmark inside a removed paragraph goes with it
while the `REF` naming it survives, leaving a live cross-reference pointing at nothing.
`docx_xml.reference_target()` reads just enough field grammar to say which bookmark a
field consumes — reusing the field-instruction reader, so a split instruction and a
quoted name both resolve. Two consumers are supported: reference fields (`REF`,
`PAGEREF`, `NOTEREF`), simple or complex, and internal hyperlinks, which name their
target in `w:anchor` rather than through a field.

The inventory is **document-wide**, the opposite of field carriers and for the opposite
reason: a reference in a header legitimately names a bookmark in the body, while a field
lost from a header is not answered by an identical one in the body. Names match
case-insensitively, as Word matches them, and are reported as written.

References are kept **per consumer**, not collapsed by target name. One missing bookmark
can break references in the body, a header and a note at once, and each is a separate
place someone has to go and repair — a report naming only the bookmark says what is wrong
without saying where. Each surviving consumer is reported with its part and with the
instruction or anchor that names the target.

Only what this run broke is reported, and there is deliberately **no tracked-deletion
exemption** — unlike the field inventory. Accepting a revision that deletes a referenced
target is a requested deletion with an unrequested consequence, and the consequence is
what needs review. Nothing is repaired: no target invented, no field retargeted, no
replacement bookmark created. Leaving an empty bookmark behind would not count as fixing
this — it suppresses one error message while letting the `REF` return misleading content.

**A table accepted revisions emptied.** Deleting a table's last row already worked; the
`w:tbl` stayed, holding nothing, and neither lint nor verification could see it. That is
the standing demonstration that **a clean lint is not evidence Word will accept a
package** — the lint had no rule for the shape, so it reported nothing, which is not the
same as reporting that nothing is wrong. Nesting needs no special traversal: a table with
no rows has no cells, so it can hold no inner table, and is always a leaf.

**Only tables this acceptance actually emptied are removed.** A table that arrived rowless
is the document's own problem; removing it would rewrite a file that had no revisions to
accept, merely because the option was on. It is still linted, so it is visible without
being silently repaired — the same rule the structural comparison follows, that only what
this run did is this run's doing.

**A removed paragraph that took part in automatic numbering.** Its own category —
`VerificationResult.numbering` — because no text-integrity claim is being made. Nothing
was lost; what may have changed is the numbers a reader sees. `docx_xml.numbering_id()`
resolves direct `w:numPr` then the style chain, and treats `w:numId` `"0"` as the
override it is rather than a list called zero. A paragraph naming no style still has one —
Word applies the default paragraph style — so the chain starts there when `w:pStyle` is
absent, rather than at `None` and examining nothing. A notice is raised only when the list
still has surviving members: a list whose every paragraph went renumbers nothing, and a
notice about it would be noise dressed as precision.

Nothing is renumbered and no cross-reference is rewritten. The reference check above is
what asserts a specific reference broke, and it names the bookmark.

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

Two suites carry the verification contract, and they ask opposite questions:

- `tests/test_evidence.py` cleans each case **for real** and checks that the evidence
  read from the source predicted the output. It is the anti-drift guard: if
  `paragraph_evidence()` and the processor ever disagree, verification has become a
  second opinion about a different program.
- `tests/test_verify.py`'s `InjectedDamageTests` builds source and damaged output
  **independently, never by running the cleaner**, because agreement between a broken
  verifier and the cleaner that produced its input proves nothing. Each case carries an
  anchor paragraph present on both sides — substituting a placeholder would make the
  case fail on the invented text no matter how the loss was classified, and a tripwire
  that fires for the wrong reason is not a tripwire. V09 is the false-alarm guard: it
  asserts a *correct* clean still passes, which is what stops the contract being
  satisfied by a verifier that simply distrusts everything.

Add a test whenever you touch removal safety, the pattern tiers, or verification classification.

Two conventions matter here:

- **GUI tests skip wherever `tkinter` is absent.** That is a property of the
  environment, not of the platform: a minimal Linux container has no `tkinter`, but
  `actions/setup-python` ships one, so both CI lanes run them. Logic that needs
  testing must still not live in `gui.py` — that is why `batch.py` exists — and when
  changing `gui.py` in an environment without `tkinter`, say plainly that its tests
  did not execute rather than reporting a green suite as though they had. Read the
  skip count; do not infer it from the operating system.
- **A test that anticipates a later fix is carried under `unittest.expectedFailure`,
  never as an ordinary failing test.** A permanently red suite cannot validate
  anything, and hides real regressions. The decorator keeps the suite green while the
  defect stands and turns it red — `unexpected successes=1`, exit code 1 — the moment
  the behaviour changes, so the tripwire fires in both directions and the decorator
  cannot be forgotten. `tests/test_shipped_policy.py` describes how its fixtures were
  carried this way until W02 narrowed the rules, and the V04 case in
  `tests/test_verify.py` how it was carried until W03 replaced the containment
  predicate with interval coverage. Both reported the unexpected success that said the
  decorator could go. No case currently carries one — when you add one, its docstring
  names the package that closes it.
- **A measurement tool is held to the same standard as the code it measures.**
  Everything under `tools/` is tested, because a wrong number is what a decision
  gets taken on. Two traps, both of which produced real bugs here:
  - **Measure actions, not detections.** `ProcessingResult.detections` is what the
    detectors *found*; the processor acts once per paragraph. A note matching a
    specifier rule and a copyright rule is two detections and one removed paragraph,
    and a paragraph of nothing but a placeholder is an inline detection that gets
    *removed* whole. `tools/actions.py` mirrors `_process_xml_file`'s decision order
    and asks the processor's own methods; everything else builds on it rather than
    walking detections.
  - **lxml element identity needs a live reference.** Proxies are built on demand and
    released when nothing holds them, so `id(element)` is stable only while a list of
    those elements is alive. A position map built across two `root.iter()` passes is
    wrong — materialise once and hold it.

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

Verification picks the category up automatically from `DetectionEngine.paragraph_evidence()`.

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
