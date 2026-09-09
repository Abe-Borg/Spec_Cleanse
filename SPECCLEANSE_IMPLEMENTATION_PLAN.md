# SpecCleanse: detailed implementation and agent handoff plan

**Prepared:** 2026-09-08
**Revised:** 2026-09-08 (revision 2 — resequenced after independent verification and review)
**Repository:** `C:\Github-Repos\Spec_Cleanse`
**Reviewed baseline:** `dfb9c8a44d48b79877a097d14ba6d17ec7dcc27d`
**Verification baseline:** `c849938` — plan document only; every application file is byte-identical to `dfb9c8a`
**Status:** proposed implementation work; no application changes have been made as part of preparing or revising this plan.
**Audience:** capable coding agents and the maintainer who will review their work.

## Revision history

**Revision 2** resequences the work and narrows three commitments. It does not overturn any finding —
all seventeen survived independent re-testing, and most were reproduced end to end against real
`.docx` packages rather than in memory. What changed:

1. Four corrections with no dependency on the evidence model were pulled out of the middle of the
   plan into a first package, **W00**.
2. Corpus evidence moved from the last package to the second, and again at completion.
3. Sequential implementation replaced the multi-agent allocation.
4. Reference-integrity work was rescoped from *retain the target's content* to *detect and report*,
   with retention deferred to a later, separately justified decision.
5. The verifier gained explicit false-alarm acceptance criteria — deliberately **not** a percentage target.
6. Performance language was corrected from "approximately cubic" to what the measurements support.

Package IDs were renumbered to match the new order. Mapping from revision 1: old W00 → W01,
W01 → W02, W02 → W03, W03 → W04, W04 → W05 (less F09), W05 → W06, W06 → W07 (less F12/F13 basics),
W07 → W08, W08 → W09, W09 → W10. **W00 is new.**

## 1. Objective and scope

Make SpecCleanse safer at preserving specification requirements, make its verification verdicts more
reliable, prevent batch output loss, and address the measured large-document bottleneck without
replacing the existing architecture.

This plan is the deliverable requested after an analysis-only review. Creating and revising this
Markdown file is authorized. It is not a record that the implementation, full test suite, Word
validation, or release has already happened. A receiving coding agent should implement the work when
assigned this plan, subject to the permissions and instructions in that receiving session.
Publishing a release is outside this implementation plan.

The highest-value outcome is fewer silent false positives. Retaining some uncertain editorial text is
preferable to deleting a requirement. Verification must not interpret a substring match or one
editorial run as permission to discard unrelated text.

Retain the flat Python layout, direct `lxml` manipulation, stdlib `unittest`, Tkinter GUI, and the two
runtime dependencies. Do not bring back the archived structural or style optimization stages. Do not
introduce a service, database, LLM call, dependency-heavy rule framework, or broad package
reorganization.

The plan deliberately separates:

1. Confirmed cleaner behavior that needs correction.
2. Confirmed verifier weaknesses exposed by deliberately damaged outputs.
3. Structural concerns requiring standards and Word validation.
4. Measured performance work.
5. Optional product ideas that are deferred.

## 2. Evidence, confidence, and corrections

### 2.1 What the first review established

The original review brief was supplied at `C:\Users\AbrahamBorg\Downloads\spec-cleanse_REVIEW_BRIEF.md`.
It contains useful leads, but its claims and instructions are not a substitute for independent
investigation.

The first follow-up review read the active cleaner, detection, XML, verification, and GUI paths and
performed synthetic checks on Windows using Python 3.14.6. Python bytecode caching was disabled. The
broader probe explicitly blocked filesystem mutations. XML trees, source metadata, and output
comparisons stayed in memory; file-operation boundaries were replaced for these probes.

Consequently that pass:

- Established behavior of the relevant application functions, including the actual XML redactor and
  classification logic.
- Did **not** exercise ZIP unpacking/repacking end to end.
- Did **not** rerun the normal test suite, which creates temporary files.
- Did **not** measure precision or recall on a real specification corpus.
- Did **not** open or round-trip a generated document in Word.

### 2.2 What independent end-to-end verification established

A second, independent pass re-tested the findings without reusing the first pass's probes. It built
real `.docx` packages with `tests/docx_builder.py`, ran the public `DocxProcessor.process()` and
`verify_clean()` paths, and reread the output packages. Damaged outputs for verifier tests were
constructed by hand, never by running the cleaner.

Environment: Linux, Python 3.11, `lxml` 6.0.2, system PyYAML. Existing suite green at
**92 passed, 4 skipped** (the skips are the GUI tests, which require `tkinter`).

This closes the first pass's largest gap — the findings are no longer in-memory leads; they reproduce
through ZIP round-trip and the public API. It does **not** close the others: no real corpus, no Word
round-trip, and no Windows GUI execution. Those remain outstanding and are scheduled in W01 and W10.

Line numbers cited anywhere in this plan are navigation hints against the verification baseline, not
stable identifiers.

### 2.3 Confirmed findings

| ID | Finding | Status | Evidence and consequence |
|---|---|---|---|
| F01 | Copyright patterns classify plausible requirements as removable paragraphs | Reproduced end to end | All three examples in §8.2 were deleted by a real clean. One pattern match scores 0.7 (`detection.py:254`), clearing the 0.5 threshold alone. |
| F02 | Some high-confidence editorial patterns also overreach | Reproduced end to end | `Select one of the listed manufacturers.` and `Provide two [retain or delete] spare filters per unit.` were both deleted whole. |
| F03 | Verification drops the conditions attached to low-confidence patterns | Reproduced end to end | A plain paragraph containing `revise as required` is **retained by the cleaner** but a hand-damaged output missing it verified **PASS**, categorised `editorial_artifact`. `removal_patterns()` flattens the low-confidence tier without its formatting prerequisite (`detection.py:633`). |
| F04 | Verification promotes partial formatting evidence into whole-paragraph permission | Reproduced end to end | Deleting a paragraph of ordinary requirement text carrying one hidden run verified **PASS**, categorised `formatting_based / hidden text`. `is_hidden |= run_hidden` (`verify.py:356-357`) lets one run speak for the paragraph. |
| F05 | Verification does not carry preserve-style evidence | Reproduced end to end | Deleting an `ART`-styled heading verified **PASS**. `_matches_preserve` (`verify.py:461`) consults preserve *patterns* only; the cleaner's `PreserveDetector` also honours preserve *styles*. |
| F06 | Verification allows an inline match to excuse extra deleted words | Reproduced end to end | `Provide two [Verify quantity] spare filters per unit.` → `Provide two filters per unit.` verified **PASS** as `inline_placeholder`. Classification asks whether a fragment *contains* a match (`verify.py:695`), not whether matches *cover* it. |
| F07 | A correct long inline redaction produces a false verification failure | Reproduced end to end | The real cleaner correctly produced `Provide units.`; verification reported one unexpected removal plus one addition. `MIN_PAIR_SIMILARITY = 0.5` (`verify.py:74`) rejects the pair. |
| F08 | Simple fields are missing from the embedded-content safeguard | Reproduced end to end | An editorial paragraph carrying `w:fldSimple` lost the field (1 → 0); text verification and structural inspection both passed. `EMBEDDED_CONTENT_TAGS` covers `w:fldChar` and `w:instrText` but not `w:fldSimple`. |
| F09 | Non-`w:t` characters within redactions can remain | Reproduced end to end | A placeholder containing `w:noBreakHyphen` produced exactly `Provide -units.`. `_redact_spans` advances the offset for separators but edits only `w:t` (`processor.py:427`). |
| F10 | Complete bookmark targets can be lost | Reproduced end to end | A bookmark fully inside a removed paragraph disappeared (1 → 0) while a `REF` field elsewhere still named it; verification passed. `_relocate_orphaned_markers` rescues only *half-open* ranges. |
| F11 | Accepting the last deleted table row leaves a zero-row table | Reproduced end to end | Rows 1 → 0 with the `w:tbl` still present; structural lint reported clean and verification passed. Whether Word repairs or rejects the package is **still untested**. |
| F12 | Batch destinations can collide | Confirmed by inspection and mapping simulation | `_output_for()` (`gui.py:465`) maps same-named inputs from different folders to one destination. `_confirm_overwrite` runs before processing and checks only files already on disk, so the second write is silent. |
| F13 | Verification failure is counted as GUI success | Confirmed by inspection | `_clean_one()` returns `True` after logging a failed verification; `_run_clean()` counts it as succeeded. The log prints `FAIL`, the summary prints "succeeded". Not executed: GUI tests skip without `tkinter`. |
| F14 | Verification's main paragraph alignment dominates the large-input cost | Reproduced by profiling | At 6,000 paragraphs the outer `difflib.find_longest_match` accounted for ~97% of profiled verification time; `_pair_with_survivor()` accounted for **0.97%**. |
| F15 | Paragraph deletion repeatedly scans the entire parent container | Confirmed by inspection | `can_delete_paragraph()` materialises the parent's full block-child list per candidate removal. Secondary to F14. |
| F16 | Verifier fallback configuration bypasses the existing path resolver | Confirmed by inspection — **latent** | `verify_clean()` computes `Path(__file__).parent / "patterns.yaml"` (`verify.py:538`). The GUI always supplies `engine=`, so the fallback is currently computed and never read. It governs no production run today; it would govern any engine-less caller. |
| F17 | Test-only pull requests can miss CI | Confirmed by inspection | `release.yml` is the only workflow. Its pull-request filter lists `*.py`, which matches root files only, so a PR touching just `tests/**` triggers nothing. The workflow does run the suite when it fires. |

F03–F06 are verifier tests with injected damage. **Do not describe them as observed extra deletions by
the current cleaner.** F03 in particular was re-confirmed in both directions: the cleaner keeps the
paragraph, and the verifier would nonetheless accept its deletion.

### 2.4 Performance evidence

Two independent measurements, on different machines and interpreters, agree on the shape.

First pass — Windows, Python 3.14.6, synthetic paragraph distribution, ZIP and disk excluded:

| Body paragraphs | Cleaning logic | Verification logic |
|---:|---:|---:|
| 500 | 0.0282 s | 0.0073 s |
| 2,000 | 0.1536 s | 0.1445 s |
| 6,000 | 0.7957 s | 2.9939 s |

Second pass — Linux, Python 3.11, real `.docx` packages, **ZIP and disk included**:

| Body paragraphs | Clean | Verify |
|---:|---:|---:|
| 500 | 0.0502 s | 0.0276 s |
| 2,000 | 0.2561 s | 0.2711 s |
| 6,000 | 1.3178 s | 4.6726 s |

**The decisive measurement is the controlled comparison at fixed size.** Six thousand paragraphs with
the brief's repeated part heading verified in **4.6726 s**; the same six thousand paragraphs with
*unique* headings verified in **0.6105 s** — a 7.7× difference attributable to duplicate paragraph
text alone.

The profile explains it: `difflib.SequenceMatcher` builds a `b2j` index of element → positions, and
duplicate elements produce large candidate buckets that `find_longest_match` rescans. The 6,000-paragraph
profile recorded 44.4 million `dict.get` calls inside that function.

Profiled shares at 6,000 paragraphs (cProfile inflates wall clock roughly 3× in both passes):

| Component | Share of verification |
|---|---:|
| Outer paragraph `find_longest_match` | ~97% |
| `_pair_with_survivor()` | 0.97% |
| `extract_paragraphs()` (both sides, incl. unzip) | ~1.7% |

Interpretation, stated at the confidence the data supports:

- Growth over the measured **duplicate-heavy** range is **super-quadratic and accelerating** — an
  observed exponent near 1.65 from 500→2,000 rising to about 2.6 from 2,000→6,000. Do **not** call it
  cubic; the data does not establish a complexity law, and §16.1 forbids extrapolating one.
- Duplicate candidate buckets are the demonstrated mechanism. That makes unique-anchor alignment a
  strong candidate — but it does **not** establish patience/LIS as the final answer. A region with few
  or no unique anchors still needs a correct and efficient fallback, which is exactly fixture family 4
  in §16.1.

The brief's generator repeats the same part heading (when `i % 12 == 1`, `i % 3 + 1` is always 2).
Keep that distribution as an **adversarial regression fixture**, not as the representative case, and
assess performance across the distributions in §16.1.

### 2.5 Corrections that must inform implementation

These correct the original brief. Three were re-verified during the second pass and are marked.

- **Removing global `DOTALL` does not fix the supplied copyright false positives.** *Re-verified:* all
  three match on one line, with `DOTALL` off.
- Shorter wildcard gaps reduce some false positives but do not prove a paragraph is boilerplate.
- **For a pure deletion, similarity `2 * len(after) / (len(before) + len(after)) < 0.5` means more than
  two-thirds of the original characters were lost, not roughly half.** *Re-verified:* the ratio is
  exactly 0.500 at 33% retained; 40% retained scores 0.571 and pairs; 25% retained scores 0.400 and is
  rejected.
- **Eliminating a duplicate character comparison cannot halve total verification time.** *Re-verified:*
  the character-level work is under 1% of the profile. It is a secondary optimization.
- Immediate sibling checks are not automatically equivalent to the current block-container rules:
  markers and other non-block children can intervene.
- A CLI is not a prerequisite for a corpus harness; the harness can import the existing modules.
- Flattening the verifier's paragraph list is not a complete specification text export. Generated
  numbering, table relationships, notes, fields, and part identity need an extraction design.
- Sorting ZIP member names alone does not make archives byte-deterministic; timestamps and metadata
  also matter.
- Automatic numbering consequences depend on list membership, restarts, and inheritance. Do not claim
  every later number necessarily changes.
- Valid-looking XML and a passing custom lint are not proof that Word will open a document without
  repair. F11 is the live example: lint reports clean on a zero-row table.

### 2.6 Corrections to the review of this plan

Revision 2 incorporates an outside review of revision 1. Several of that review's *recommendations*
were adopted; several of its *supporting claims* were not evidenced and are recorded here so they are
not mistaken for findings. The plan holds itself to the same standard it demands elsewhere.

| Reviewer claim | Disposition |
|---|---|
| Specific line counts and a "ships in days" estimate for the W00 fixes | **Not adopted as fact.** No implementation was performed. W00 is ordered first because its dependencies are minimal and its benefit is immediate — not because its size is known. |
| F12 is "the only finding that destroys user data" | **Too narrow.** Silent removal of a requirement is consequential even when the source file survives, because the cleaned output is what downstream analysis consumes. The real distinction is the recovery path: F12 leaves no output at all for the losing file, while F01/F02 corrupt a derived artifact with the source intact. |
| Revision 1 "prescribed five simultaneous agents" | **Overstated.** Revision 1 described roles and later waves, and already folded the performance role into a subsequent wave. The valid part of the criticism — that the proposed ownership overlapped so heavily the coordination rules were compensating for it — is adopted in §6. |
| "Protect referenced ranges" and "do not protect every bookmark" are contradictory | **Not contradictory.** Protecting only *referenced* ranges while not protecting *all* bookmarks is coherent. The valid criticism — revision 1 did not bound the protection's cost or its effect on cleaning — is adopted in §13.1. |
| Word's automatic bookmarks would substantially suppress cleaning | **Plausible risk, not an established outcome.** The mechanism is real; the effect is unmeasured. W01 adds a census that settles it either way (§8.3). |
| A blanket "≥X% Verified" acceptance target | **Rejected.** It has a perverse incentive: the cheapest way to hit a Verified rate is to make the verifier more permissive, which is the exact failure W03 exists to eliminate. Replaced with the criteria in §10.6. |

One reviewer observation is adopted as a **new technical constraint**, not merely as sequencing advice:
F06 and F07 live in the same code path (`_pair_with_survivor` → `_classify_modification`). See §5.2.

## 3. Decisions selected for this implementation

### 3.1 Required work, in execution order

1. Ship the four corrections that depend on nothing else (W00).
2. Establish a baseline and take corpus census measurements that inform the policy decisions (W01).
3. Narrow shipped whole-paragraph rules; decide the formatting-only default on the census evidence (W02).
4. Preserve the scope, prerequisites, and source location of removal evidence during verification (W03).
5. Recognize exact permitted inline transformations regardless of retained-text percentage; preserve
   part and story identity (W04).
6. Complete simple-field preservation and the remaining inline-redaction cases (W05).
7. Detect newly broken internal references; handle tables emptied by accepted revisions; add numbering
   review notices (W06).
8. Complete the batch manifest and the full file-outcome model (W07).
9. Use the shared configuration resolver and cover nested test changes in CI (W08).
10. Optimize paragraph alignment after its correctness contract is established; optimize container
    scanning separately (W09).
11. Update documentation, rerun the corpus comparison, and validate changed XML cases in Word (W10).

### 3.2 Explicit product choices

- Keep producing a cleaned output when verification finds a problem, provided processing completed
  successfully. Mark it **Needs review**, distinguish it from a verified output, and identify its
  path. Do not silently delete it or rename it after the fact.
- A successful write and a passing verification are separate outcomes.
- Ambiguous prose such as a contractor being told to select a manufacturer is retained by default.
  Additional editorial text retained by narrower rules is an accepted tradeoff and must be documented.
- **The `specifier_notes.formatting_only_removal` default flip is now conditional on evidence.** The
  reproduced requirement false positives (F01/F02) justify correcting unsafe default rules on their
  own. The formatting-only flip is a different decision with a different cost, and it is taken in W02
  **after** the W01 census reports how much of the current removal volume depends on formatting-only
  evidence. Choosing a conservative default under asymmetric error costs does not require pretending
  recall was measured — but it does require knowing what the change costs this workflow.
- **Reference integrity is detect-and-report first.** Preserve the original paragraph, report a newly
  broken supported reference as **Needs review**, and defer automatic retention of referenced content
  to a later, separately justified and tested decision. Leaving an empty bookmark behind does not
  count as fixing reference integrity.
- Do not renumber documents or rewrite cross-references automatically. Report numbering-sensitive
  deletions as requiring review when reliably identified.
- Existing user-owned configuration files must not be silently overwritten or migrated. Safer defaults
  do not automatically repair an installed user's old patterns.

### 3.3 Deferred work

Do not include these unless later evidence and a separate task justify them:

- A public cleaning CLI, text/Markdown output sidecar, section chunking, deduplication, token
  estimation, caching, or multiprocessing.
- Converting to `python-docx`, adding pytest, reorganizing into a package, or restoring `legacy/` stages.
- General-purpose OOXML repair or complete schema validation of arbitrary input packages.
- Broad style-resolution redesign, speculative security hardening, or formatting changes unrelated to
  the confirmed defects.
- ZIP byte determinism, per-pattern telemetry, a new persistent settings system, or a general rules DSL.
- **Automatic retention of referenced bookmark target content** (see §3.2 and §13.1).

Small developer census and corpus/benchmark runners under `tools/` are in scope. They are internal
utilities, not a new public CLI product.

## 4. Non-negotiable invariants

1. **Preserve decisions dominate editorial cleaning.** Source text protected by a preserve pattern,
   preserve style, or explicit reference-target protection cannot be deleted or redacted through
   another editorial path. Explicit supported tracked deletions have the separate authority and
   consequence reporting defined in §10.4.
2. **Evidence has a scope.** An inline match permits an interval removal. A qualifying run permits
   removal of its own text. Whole-paragraph removal needs whole-paragraph authority or complete
   coverage by eligible intervals.
3. **Prerequisites travel with rules.** Disabled categories, formatting-only settings, and
   low-confidence formatting requirements apply equally during cleaning and verification.
4. **Verification rereads the actual input and output.** The processor's target list, claimed
   removals, or `ProcessingResult` cannot be the sole authority for acceptance.
5. **Protected text remains in order.** No unexplained insertion, substitution, reordering, or loss of
   multiplicity becomes PASS because a fuzzy match found similar text elsewhere.
6. **Part identity is retained.** Text in a header cannot account for body text that disappeared.
   Notes and distinct content parts must not be flattened into one interchangeable pool.
7. **Offsets have one definition.** Use the same paragraph-owned text stream to interpret source
   intervals. Boundary trimming and any permitted whitespace cleanup must have an explicit mapping.
8. **Nested paragraphs are independent.** Do not cross into text-box paragraphs while handling an
   outer paragraph's runs or text.
9. **Structural carriers survive ordinary editorial removal.** Keep required blocks, section
   properties, fields, pictures, and note anchors. Accepted revisions may remove structures only under
   the explicit revision option and its documented rules.
10. **Mutations remain staged.** Collect targets, apply redactions, then run removals, then paragraph
    removals, or prove a revised order preserves all offsets and detached-element behavior.
11. **No silent batch overwrite.** A batch must have distinct output destinations, and no destination
    may overwrite a selected source document.
12. **Unknown is not PASS.** If alignment or structure cannot be verified, report the limitation as
    requiring review. Do not hide unresolved differences for speed.
13. **Do not improve runtime by removing safety checks.** Preserve-classification, multiplicity, part
    boundaries, and structure checks remain mandatory.
14. **Report what was tested.** Do not claim real-corpus precision, recall, or Word compatibility from
    synthetic tests alone.
15. **A review verdict must be actionable.** Invariant 12 requires reporting uncertainty; this one
    requires that the report be usable. Ambiguous alignment, detected damage, configuration notices,
    and reference/numbering warnings are distinct categories, counted separately and named
    individually in the outcome — never pooled into one undifferentiated rate. A verifier whose
    warnings cannot be acted on is not safe, only noisy.

## 5. Execution sequence and known couplings

### 5.1 Work packages, in order

| # | Package | Purpose | Depends on | Primary files |
|---|---|---|---|---|
| W00 | Immediate corrections | Stop batch output loss, stop false success reporting, stop the orphaned separator, stop the false alarm on long redactions | Nothing | `gui.py`, `processor.py`, `verify.py`, focused tests |
| W01 | Baseline, fixtures, and corpus census | Establish a trustworthy starting point and the evidence the policy decisions need | W00 landed so the baseline reflects it | Tests, `tools/` census and corpus utilities, developer notes |
| W02 | Scoped default policy | Reduce requirement deletion under default policy | W01 census | `patterns.yaml`, `detection.py`, config tests |
| W03 | Source-based verification contract | Stop permissive acceptance of unexplained loss | W01; policy contract agreed with W02 | `detection.py`, `verify.py`, metadata helpers in `docx_xml.py`, verifier tests |
| W04 | Pairing, part identity, ordering | Remove false alarms without hiding real edits | W03 | `verify.py`, verifier tests |
| W05 | Inline redaction and field carriers | Complete the XML redaction and carrier defects W00 did not close | W03 contract | `processor.py`, `docx_xml.py`, inline/processor tests |
| W06 | Reference integrity, revision-empty tables, numbering notices | Report broken references and prevent malformed table residue | W03, W05 | `docx_xml.py`, `processor.py`, `verify.py`, structural/revision tests |
| W07 | Batch manifest and full outcome model | Complete what W00 started: explicit verified/needs-review/failed outcomes | W00; result types from W03/W06 | `gui.py`, GUI tests |
| W08 | Configuration fallback and CI | Fix config selection and test coverage | W02 for default behavior | `verify.py`, `apppaths.py` if necessary, workflow files, config/path tests |
| W09 | Measured performance | Make large manuals practical | W04, W05, W06 correctness settled | `verify.py`, `docx_xml.py`, benchmark utility, targeted tests |
| W10 | Integrated validation and documentation | Establish what is actually ready to trust | All implemented packages | Tests, `README.md`, `CLAUDE.md`, `CHANGELOG.md`, final report |

Safety corrections must remain deliverable even if performance work needs another iteration. Do not
make a faster diff algorithm a prerequisite for fixing false-positive deletion.

### 5.2 Known couplings that must be managed, not discovered

**F06 and F07 share one code path.** Both are decided in `_pair_with_survivor()` →
`_classify_modification()`. The W00 fix for F07 was verified not to worsen F06: a damaged output whose
text does not equal the independently computed expected text simply falls through to the existing
similarity path, where F06's permissiveness still lives. They can therefore ship separately — but only
deliberately:

- W00 implements the **narrow** F07 correction only: compute the expected text from authorized spans
  and accept an exact match *before* consulting similarity. It must not lower `MIN_PAIR_SIMILARITY`,
  accept all short survivors, drop `added` from the pass predicate, or otherwise create a shortcut
  that admits extra deletions.
- W00 also lands F06's injected-damage test, marked with the package that closes it (W03) and the
  reason. Land it under **`unittest.expectedFailure`**, not as an ordinary failing test. A suite left
  red from W00 through W02 cannot validate a package boundary — every run fails for a known reason,
  so a genuine new regression is obscured and "was it green?" stops answering anything, which is the
  same objection invariant 15 makes about a warning nobody can act on. `expectedFailure` avoids that
  and is a *stronger* tripwire than a red test, because it fires in both directions:

  | State | unittest reports | Exit |
  |---|---|---|
  | V04 still failing (W00–W02) | `OK (expected failures=1)` | 0 — suite green |
  | V04 starts passing (W03, or a wrong F07 fix) | `FAILED (unexpected successes=1)` | 1 — suite red |

  Both rows were verified against stdlib `unittest`. A plain red test says nothing when someone
  changes the behavior; this one fails loudly the moment V04's verdict changes for any reason. W03
  removes the decorator when it lands the real fix — the unexpected-success failure is what forces
  that, and it is a feature, not a chore.
- W03 closes F06 by changing one predicate in the same function: require that every deleted fragment
  be **covered by** authorized spans, not merely that it **contains** a match.

Whoever takes W00 will be reading the code that contains F06. Taking F06 there instead is acceptable
if it can be done without widening the change; deferring it is acceptable; leaving it undocumented is not.

**The F07 correction pays a second dividend in W09.** A paragraph whose output exactly matches its
independently computed expected redaction result is a *reliable pairing*, and therefore a usable
alignment **anchor** — not only a candidate to be resolved inside an already-bounded region. §16.2
step 2 should admit exact permitted transformations as anchor candidates alongside unchanged
paragraphs. That matters most in exactly the documents where the bottleneck bites: a paragraph with a
distinctive placeholder is a unique anchor even where its neighbours repeat.

**Corpus work has a tooling prerequisite.** Moving corpus evidence to W01 means building the
comparison harness in W01. The performance benchmark utility described in §16.1 is a different tool
with a different job; do not conflate them or make one wait for the other.

**The formatting-only switch reaches further than its own detector.** `SpecifierNoteDetector` consults
`formatting_only_removal` directly, `EditorialArtifactDetector` does not, and verification reads the
same key to set `trust_formatting_only` for *all* categories. Changing the default therefore changes
verification behaviour for editorial artifacts as well. Test the four interactions separately.

### 5.3 Relative effort, risks, and delivery boundaries

These are planning comparisons. They are **not** promises of agent-hours, and no line-count or
calendar estimate in this plan should be treated as a commitment — none of this work has been
implemented.

| Package | Relative effort | Main benefit | Principal implementation risk |
|---|---|---|---|
| W00 | Small, four separable changes | Stops data loss and the two most misleading reports immediately | An over-broad F07 fix that admits extra deletions; a separator fix that discards page/column break semantics |
| W01 | Small to medium; depends on corpus availability | Turns two policy arguments into measurements | Building a harness that bakes in the implementation's assumptions instead of independently describing behavior |
| W02 | Medium | Reduces requirement deletion under default policy | Reduced editorial recall; installed users continuing to run old customized rules |
| W03 | Large, highest reasoning demand | Stops permissive verifier acceptance of unexplained loss | Incorrect interval coordinates, evidence applied at the wrong scope, accidental coupling to processor actions, or a false-alarm rate that makes the verdict useless |
| W04 | Medium to large | Removes false alarms without hiding real edits | Ambiguous duplicate text, lost multiplicity, or pairing across part/story boundaries |
| W05 | Medium | Corrects remaining XML redaction/carrier defects | Stripping a field or anchor while cleaning empty runs |
| W06 | Large; split references, revisions, and numbering into separate changes | Makes reference and structural consequences visible | Incomplete field parsing, unbounded protection scope, nested-container edge cases |
| W07 | Medium | Completes honest batch reporting | Path equivalence on Windows; incomplete propagation of new outcome types |
| W08 | Small to medium | Fixes config selection and test coverage | Unwanted seeding during engine-supplied verification |
| W09 | Medium to large, evidence dependent | Makes large manuals practical | Faster but incorrect alignment; stale container metadata after mutations |
| W10 | Medium plus corpus/Word availability | Establishes what is actually ready to trust | Treating unavailable manual evidence as a pass |

Verify each integrated stage before changing the next shared contract. A fixture may land with its fix
so mainline is not deliberately left red. **Mainline stays green throughout**, including across the
W00–W02 window: the one test that anticipates a later fix is the F06 case in §5.2, and it is carried
under `unittest.expectedFailure` precisely so it does not turn the suite red. If an agent proposes a larger redesign, require it to explain why the
smaller correction cannot satisfy the same acceptance tests; prefer the smaller correction when both work.

## 6. Working model and ownership

**Sequential implementation is the default.** The four modules that matter — `verify.py`,
`processor.py`, `docx_xml.py`, `detection.py` — are tightly coupled, and any allocation that splits
them produces more coordination than work. Revision 1's coordination rules were compensating for an
allocation that did not need to exist.

- **One implementing agent** works through the packages in order, in coherent stages.
- **An independent reviewer** examines significant boundaries: after W00, after W03, after W06, and at W10.
- **The injected-damage fixtures in §10.5 belong to the reviewer, not the implementer**, and should be
  written before or alongside the implementation rather than after it. W03's specific failure mode is
  a verifier that agrees with its own implementation; fixtures authored by the person who wrote the
  implementation are the one thing that cannot catch it.
- **CI (the workflow half of W08) is genuinely separable** and may run as a parallel task. It touches
  no application module.
- More parallelism should **follow** a demonstrated separation of work, not precede it.

Each package's handoff must contain: problem, final behavior, files touched, exact tests run and their
result, intentional behavior changes, remaining risks, and any required manual validation. A passing
test count without this context is insufficient.

## 7. W00: immediate corrections

Four changes with no dependency on the evidence model, each separately tested. They are ordered first
because their dependencies are minimal and their benefit is immediate.

### 7.1 F12 — reject conflicting output destinations before any write

Build the input/output manifest on the main thread from the snapshotted selection and destination, and
validate it completely before starting the worker or writing anything.

- Two selected inputs must not map to the same destination.
- A planned output must not equal any selected input, including one that will be processed later.
- Apply Windows case-insensitive normalized path comparison. Use resolved paths; for existing paths,
  consider filesystem identity to catch equivalent spellings. Do not assume string inequality proves
  two paths are different files.
- Retain the existing overwrite confirmation for previously existing outputs, applied only **after**
  internal collisions have been ruled out.

On collision: reject the batch before any write, naming the conflicting sources and the shared
destination, and tell the user to choose separate destinations or reorganize the sources. Do not
append arbitrary suffixes and do not let the later file win.

Pass the validated manifest to the worker; do not recompute output paths later from mutable GUI state.
This is preflight protection, not a claim of locking the directory against other processes.

W00 delivers the manifest and the rejection. The full outcome model is W07.

### 7.2 F13 — separate successful processing from successful verification

`_clean_one()` currently returns `True` after logging a failed verification, and `_run_clean()` counts
that as succeeded. Distinguish the two conditions and report them honestly. W00 does not need the
final outcome enum from W07 to do this; a minimally honest three-way report is enough:

- Processing failed → **Failed**, stating whether an output exists.
- Processing succeeded, verification failed or could not complete → **Needs review**, with the output
  path named, not hidden.
- Processing succeeded and verification passed → **Verified**.

Replace `Verifying no spec content was lost...` with a statement of what the check actually
establishes, and replace `PASS — every change matches a rule and the structure is intact` with wording
that does not imply engineering or Word-format certification. `verify.py`'s own module docstring
already concedes the check is a consistency check rather than an independent one; the GUI text should
not claim more than the module claims for itself.

Update the current tests that assert `_clean_one()` returns a Boolean. Do not keep a misleading
truthiness compatibility shim unless a real caller requires it.

### 7.3 F09 — remove supported separator elements covered by a redaction

`_redact_spans()` advances the running offset for every element `iter_text_nodes()` yields but edits
only `w:t`, so a covered separator survives as an orphan. The reproduction is exact: a placeholder
containing `w:noBreakHyphen` yields `Provide -units.`

Implement element-aware redaction:

1. Snapshot the paragraph-owned text-node stream before mutations.
2. Map each authorized span onto that stream using the existing coordinate system.
3. Partially edit `w:t` with `set_text()` as appropriate.
4. Remove a supported one-character separator element when its character is covered by a redaction.
5. Keep separators outside the interval exactly as they were.
6. Remove genuinely empty runs only after checking required structural content, range anchors, and
   already-detached state.

**Breaks need care, and are the reason this is not a one-line change.** `w:tab`, `w:cr`,
`w:noBreakHyphen` and `w:ptab` are whitespace as far as the extractor is concerned, but `w:br` carries
a `w:type` that may be `page` or `column`. Those are layout semantics, not whitespace, and the
content-only contract does not authorize discarding them. Establish and test the intended rule
explicitly — preserve such a break and report the uncertainty rather than treating all break types as
interchangeable. Do not expand text extraction to unrelated element types to make a test pass.

### 7.4 F07 — recognize an exact permitted redaction before applying similarity

For a source paragraph with eligible inline intervals, compute the expected output text from the
authorized spans without mutating XML, and accept an exact match as a pairing **before** the
similarity heuristic runs. This must work when almost all source characters were removable —
`MIN_PAIR_SIMILARITY = 0.5` rejects any pure deletion that keeps under a third of its characters, and
the reproduced case keeps 14 of 71.

Constraints, restated because the tempting shortcuts all break F06:

- Do **not** lower `MIN_PAIR_SIMILARITY`.
- Do **not** accept all short survivors.
- Do **not** drop `added` from the pass predicate.
- Do **not** suppress unexpected changes.

Use the actual source evidence, not a string supplied by `ProcessingResult`. Keep an unchanged
paragraph as an acceptable alternative outcome. Where multiple supported outcomes are permitted,
represent them explicitly or validate the deletion against the eligible intervals; do not enumerate
arbitrary subsets exponentially.

Land the F06 injected-damage test (§10.5, V04) here under `unittest.expectedFailure`, naming W03 as
the package that closes it and the reason it is expected to fail. Do **not** land it as an ordinary
failing test: that would leave the suite red across W00, W01 and W02, which is exactly when package
boundaries need a trustworthy green to validate against. See §5.2 for the verified behavior in both
directions.

**Exit gate:** no intra-batch destination collision is reachable; a failed verification cannot be
reported as success; the reproduced separator and long-redaction cases produce exact expected text;
the F06 test is present under `expectedFailure`, named and explained; the suite is green, reporting
one expected failure.

## 8. W01: baseline, fixtures, and corpus census

Corpus evidence moved here from the end of the plan. Two policy arguments — how much to narrow the
shipped rules, and whether to flip the formatting-only default — currently rest on judgment alone. Two
cheap measurements convert them into decisions.

### 8.1 Baseline procedure

Inspect the current status and applicable repository instructions. Preserve unrelated user changes.
Record the current commit and differences from the reviewed baseline. Do not assume a historical test
count still applies.

Run the existing suite before changing behavior. Use the project environment when valid, or an
available environment with the declared dependencies; do not reinstall dependencies unnecessarily.

Windows examples, from the repository root:

```powershell
& .\venv\Scripts\python.exe -B -m unittest discover -s tests -t .
& .\venv\Scripts\python.exe -B -m unittest tests.test_inline tests.test_verify
```

The `-B` flag suppresses bytecode output; the normal tests still intentionally write temporary
fixtures. Respect any write restrictions in the receiving session.

The last recorded suite state is **92 passed, 4 skipped** on Linux with Python 3.11 and `lxml` 6.0.2;
the four skips are the GUI tests. On Windows those four run, so a Windows baseline is the one that
matters for W07. Record interpreter, platform, counts, skips, and expected failures — a bare "all
green" is not a baseline. From W00 onward the expected-failure count is load-bearing: W00 adds one
(§7.4), and it must return to zero when W03 lands, so a baseline that does not distinguish expected
failures from failures cannot detect either half of that transition.

### 8.2 Regression fixtures

Create regressions using `tests/docx_builder.py` and `DocxTestCase`. Add raw XML assertions where
text-only assertions would miss damage. Do not replace these tests with mocks of the behavior being
tested.

**Copyright negatives — these must retain their requirement text under the new shipped defaults.**
All three are deleted whole by the current build.

```text
Shop Drawings submitted under this Section may not be reproduced for use on other projects.
Contractor shall verify that duplication of sprinkler coverage in adjacent zones is prohibited by the AHJ.
Unauthorized personnel shall not have access to the fire pump room; reproduction of access keys is not permitted.
```

**Additional policy negatives.** The first two are deleted by the current build; the last three
already survive and are pinned so a narrowing change does not disturb them.

```text
Select one of the listed manufacturers.
Provide two [retain or delete] spare filters per unit.
Provide pumps and revise as required.
Retain records of all tests required in Paragraph 1.6.
Select one-piece molded fittings for changes in direction.
```

For the mixed `[retain or delete]` example, the selected behavior is to retain the complete paragraph
unchanged. Do not invent a broad new inline rule merely to remove those brackets. Any later narrowly
justified inline rule must keep every surrounding requirement word and have its own positive and
negative fixtures. Note the consequence and document it: the bracketed editorial marker ships through
to downstream analysis intact, which is the accepted cost of not deleting the requirement around it.

**Representative positive fixtures that must continue cleaning.** All five behave correctly today; they
are pinned against over-narrowing in W02.

```text
© 2026 ARCOM. All rights reserved.
[Specifier: delete this note before issue]
Retain subparagraph below for wet-pipe systems.
Copy paragraphs above for each additional riser.
Provide two [Verify quantity with Owner] spare sprinklers.
```

The expected output of the last example is exactly `Provide two spare sprinklers.` — confirmed against
the current build. These strings are synthetic policy fixtures, not claims about measured corpus
frequency.

### 8.3 Census measurements

Two small utilities under `tools/`, run against representative real documents. Neither changes
application behavior, and neither requires any code change to exist first.

**Census A — formatting-only dependence.** `Detection.formatting_only` already exists, and Preview
already prints `formatting-only` beside each qualifying removal (`gui.py:85-87`), so this measurement
is available against the current build today. Report, per file and in total:

- Removals whose evidence was formatting only.
- Removals with content evidence (text pattern or editorial style).
- Inline redactions.
- Preserved paragraphs.

This is the input to the §9.2 decision. It answers what the default flip actually costs *this*
workflow — the question that decision has been resting on without an answer.

**Census B — reference overlap.** Report the share of paragraphs the cleaner would remove that lie
inside a bookmark range targeted by a supported internal reference (simple-field and complex-field
`REF`, internal hyperlink anchors). This settles, in whichever direction the data goes, whether
Word's automatic bookmarks would meaningfully suppress cleaning — which §2.6 records as a plausible
risk, not an established outcome. It is the input to the §13.1 scope decision.

Both census utilities import the existing modules. Neither requires a public CLI. Keep their output in
an explicitly chosen developer directory and do not commit proprietary content.

### 8.4 Corpus baseline

When representative real documents are available in the receiving workspace, record what the **current**
configuration removes, retains, and flags, without modifying sources and using a private temporary
output directory. A useful selection includes distinct MasterSpec, SpecLink, and firm-edited styles
when actually available; colored requirement edits; tables; references; notes; and a consolidated
manual. Do not invent coverage for a document family that was not supplied.

Build the comparison harness here — it is the prerequisite the resequencing implies (§5.2). It records,
per changed decision: source file, part and location, the original paragraph or a permitted private
preview, the action taken, and the rule or evidence behind it. W10 reruns it against the candidate and
diffs the two runs.

If no real corpus is available: complete the synthetic fixtures, record the census as outstanding, and
say so explicitly in the W02 decision rather than proceeding as though the evidence existed. Do not
fabricate precision or recall figures, and do not halt unrelated implementation solely because corpus
data is absent.

**Exit gate:** the suite's baseline is recorded with platform and skips; every fixture in §8.2 exists
and asserts current behavior; census A and B have produced numbers or are explicitly recorded as
unavailable; the corpus harness runs.

## 9. W02: scoped default policy

### 9.1 Pattern changes

Review every shipped whole-paragraph pattern, not just the copyright examples. Focus on unanchored
prose fragments, generic selection instructions, delimited note matches embedded in requirements, and
broad SpecAgent matching.

Required approach:

1. Classify each changed default rule by the scope it intends to remove and why that scope is justified.
2. Remove or narrow ambiguous standalone copyright phrases such as `may not be reproduced`,
   `duplication.*?prohibited`, and `unauthorized.*?reproduction` as independent whole-paragraph
   triggers. Note that `CopyrightDetector` scores 0.7 for a single match, so any one of these carries a
   paragraph over the threshold on its own; the scoring, not only the pattern list, is in scope.
3. Prefer patterns that recognize a complete standalone notice or a clearly editorial paragraph. A
   copyright year or publisher name somewhere inside a requirement is insufficient on its own.
4. Require explicit editorial context for ambiguous selection prose. Do not retain the broad
   `^...select one...` rule simply because one existing test expects it.
5. Keep mixed requirement/note paragraphs unless an existing, narrowly authorized inline transformation
   accounts for the removable interval.
6. Do not globally change regex flags as a substitute for a rule audit. Removing `DOTALL` fixes none of
   the three reproduced copyright false positives (§2.5). Preserve multiline notice support where
   intended and add multiline negative tests.
7. Keep named categories and the existing configuration surface understandable. Do not replace the
   detector hierarchy with a new framework.

For every edited rule, add at least one intended removal and one plausible requirement that must
remain. Test matches split across runs when matching depends on paragraph text. Test case variants and
ordinary whitespace without normalizing away meaningful content.

Use known publisher strings only as evidence for a specific notice grammar, never as universal
permission to remove any paragraph mentioning that publisher. Explicitly preserved headings still win.

### 9.2 The formatting-only default: an evidence-gated decision

Revision 1 committed to flipping `specifier_notes.formatting_only_removal` to `false`. Revision 2 keeps
that as the expected outcome but makes it conditional on census A (§8.3), because it is a different
decision from §9.1 with a different cost profile:

- The reproduced requirement false positives justify correcting unsafe default rules **on their own
  evidence**. That work proceeds regardless.
- The formatting-only flip removes a whole class of detection for firms that mark notes only by
  italic-plus-colour. Its cost depends entirely on how much of the current removal volume rests on
  formatting-only evidence — which census A measures directly.

Choosing a conservative default under asymmetric error costs does not require pretending recall was
measured. It does require stating which it is. Record the decision as: *census A reported N% of
removals were formatting-only across M documents; the default is set to X because …* — or, if no
corpus was available, *no census data; the default is set to X on the asymmetric-cost argument alone,
and this is a judgment call.* Either is acceptable. Silence is not.

If the flip proceeds, set it to `false` consistently in:

- The shipped YAML.
- `PatternConfig` and any omitted-key construction defaults.
- Verification's omitted-key behavior.
- Tests and documentation describing the default.

Explicit `true` must retain the opt-in behavior and its `formatting-only` label. Turning it off must
not disable genuinely qualifying editorial styles, hidden text, or supported low-confidence-plus-formatting
rules — those remain separate decisions, and per §5.2 the switch reaches further than its own detector,
so test the interactions separately. Update the current test that asserts removal is on by default, and
keep a separate opt-in test so changing the default does not silently delete the feature.

### 9.3 Existing installed configurations

`apppaths.resolve_config_path()` prefers an existing executable-adjacent or per-user configuration.
Changing the bundled YAML alone will not update those files.

Required transition behavior:

- Never overwrite an edited or previously seeded user configuration automatically.
- Continue logging the exact active configuration path.
- Report when formatting-only removal is enabled, including when it comes from an older explicit `true`.
- Add a focused, actionable notice for known superseded shipped patterns when they remain active.
  Compare actual rule strings against known prior defaults, not file modification times or a
  speculative version guess.
- Distinguish "new defaults are safer" from "the active configuration has been updated." Do not claim
  the latter without evidence.
- Document how an installed user can obtain the current shipped default from the distribution or
  source and compare it with the active file. A temporary frozen extraction path alone is not adequate
  documentation.
- If exact legacy-default detection is used, preserve customizations and test the known baseline, a
  customized baseline, and a current configuration.

Do not silently redefine custom regex semantics to accomplish a default-rule correction. If a proposed
detector change intentionally changes custom-rule behavior, document it with before/after examples and
have the reviewer assess the compatibility consequences.

**Exit gate:** the §8.2 negatives survive; the §8.2 positives still clean; the formatting-only decision
is recorded with its evidence or with an explicit statement that there was none; opt-in remains
available; existing user configuration is untouched; release notes explain the intentional reduction
in aggressive matching.

## 10. W03: a source-based verification contract

### 10.1 Problem to solve

`removal_patterns()` flattens policy into `(category, regex)` pairs. The verifier consequently cannot
distinguish text-only authority from a rule requiring formatting, nor a permitted interval from a whole
paragraph. `ParagraphInfo` aggregates several signals with `|=`, which loses the distinction between one
editorial run and a fully editorial paragraph. `_matches_preserve` consults preserve patterns only,
while the cleaner also honours preserve styles.

All four reproduced verifier failures follow from those three facts. Fix the distinctions directly. Do
not patch each supplied phrase with a special-case exception.

Note that `verify.py`'s module docstring already states that question 1 "shares its compiled patterns
with the detection engine … That makes it a consistency check, not an independent one." W03 is not
inventing a defect; it is closing a limitation the code documents about itself.

### 10.2 Recommended minimum data model

Use small dataclasses or a comparably explicit typed representation. Suggested information, not
mandatory class names:

| Concept | Required information |
|---|---|
| Paragraph source | Package-relative part name, relevant story/note identity, document-order position, raw owned text, display text if trimmed |
| Run source | Start/end offsets into raw paragraph text, text, relevant effective formatting/style evidence, source ownership |
| Preservation evidence | Whole-paragraph protection, reason, preserve pattern/style/reference target that established it |
| Removal evidence | Category, rule identifier/reason, source interval or whole-paragraph scope, and eligibility prerequisites already evaluated against the source |
| Revision evidence | Whether a supported tracked deletion authorizes loss under the selected option |
| Comparison outcome | Actual before/after text, accepted and unexplained intervals, preservation violations, location, and explanation |

Offsets must not be computed against stripped text and then applied to raw text. Either preserve raw
text throughout or record a reliable boundary mapping. A normalized display string is not an offset
authority.

Document-model extraction stays in `docx_xml.py`. Decisions about what qualifies as editorial stay in
`detection.py`. Comparison and classification stay in `verify.py`. A single small shared evidence type
is acceptable; a broad new abstraction layer is not required.

### 10.3 Independence boundary

It is acceptable to share immutable rule definitions and source-evidence evaluation. It is not
acceptable to trust the processor's claimed actions as proof that the resulting document is correct.

Verification must reread the original package and the actual output package and independently establish:

1. Which source characters and structures existed.
2. Which source intervals the configured policy permits removing.
3. Whether the actual output retains all protected content in order.
4. Whether structural carriers and reference targets survived as required.

Do not call the XML-mutating cleaner to manufacture verification's expected document. A source-based
pure transformation using evaluated removal intervals is appropriate; reusing processor-produced spans
without revalidation is not.

### 10.4 Classification rules

**Preservation:** compute source preserve-style and preserve-text decisions before removal evidence. A
protected paragraph cannot be excused by matching a removed fragment or an editorial pattern. Include
inherited and display-name-based styles using the existing resolver. Protect modifications as well as
complete paragraph loss. This closes F05.

**Explicit tracked deletions:** when the revision option is enabled, a supported source revision that
explicitly deletes a row/cell/run is separate authority from editorial matching. Its intended text loss
can be classified as an accepted tracked deletion even if that text would otherwise be protected from
editorial cleaning. Validate the actual source revision and its exact extent; do not use a revision
somewhere nearby as blanket permission. Broken reference consequences or unsupported structural
outcomes still require review. When the option is off, that authority does not exist. Add a
preserved-heading-in-deleted-row fixture to make this precedence explicit and avoid a new false alarm.

**Whole-paragraph removal:** accept only when a whole-paragraph rule actually qualifies at that scope,
a supported accepted revision authorizes it, or the union of eligible removal intervals covers all
substantive source text. A hidden run, colored run, or low-confidence match somewhere in the paragraph
is not sufficient. This closes F04 — replace the `|=` aggregation with scope-carrying evidence, so one
hidden run authorizes the loss of *its own text* and nothing else.

**Low-confidence rules:** evaluate formatting on the same scope as the rule. An unrelated colored run
cannot boost a plain requirement elsewhere in the paragraph. Honor category enablement and applicable
style settings. Reproduce the supported cleaner behavior deliberately; do not infer eligibility from a
regex match alone. This closes F03 — the low-confidence tier must carry its formatting prerequisite
into verification, not arrive there as a bare pattern.

**Run removal:** represent eligible run text with source intervals. Do not use substring membership in
a list of editorial run strings as location evidence: repeated text can occur in both editorial and
protected runs.

**Inline removal:** every deleted substantive character must lie within an authorized source interval.
Whitespace handling must be explicitly defined and bounded, including the existing `tidy_spans()`
adjacent-space rule. `[Verify quantity]` does not authorize deleting `spare`. This closes F06, and per
§5.2 it is one changed predicate in `_classify_modification`: require that each fragment be **covered
by** authorized spans, not merely that it **contains** a match.

**Mixed categories:** preserve violation outranks unexplained loss; unexplained loss outranks expected
changes. Report the relevant categories and reasons without labeling all fragments with the first
successful category. Avoid an unnecessary result-schema overhaul solely for presentation.

**Disabled configuration:** disabling a detector must also prevent its signals from justifying output
loss. Test hidden-text disabled, specifier-note disabled, style detection disabled, and formatting-only
disabled independently.

**Unchanged content:** verification is a loss-safety check, not a recall requirement. Retaining
eligible editorial text does not itself fail verification. Report recall separately when corpus
evaluation is available.

### 10.5 Mandatory injected-damage tests

Construct original and damaged `.docx` files independently; do not run the cleaner to produce these
damaged outputs. Per §6, these belong to the reviewer.

All twelve are now measured. "Before" is the verifier as it stood at the end of W02, on the fixtures
in `tests/test_verify.py`; "after" is the source-evidence contract. Seven of the twelve fail against
the previous verifier and pass now; five were already correct and are carried as regression guards.

| Test | Source and damage | Before | After | Required result |
|---|---|---|---|---|
| V01 | Ordinary requirement plus a hidden note run; remove the complete paragraph | **PASS**, `formatting_based / hidden text` | fails, unexplained | Unexpected removal; not PASS |
| V02 | Plain `Provide pumps and revise as required.`; remove it | **PASS**, `editorial_artifact` | fails, unexplained | Unexpected removal; not PASS |
| V03 | `ART`-styled paragraph matching a removal rule; remove it | **PASS**, `editorial_artifact` | fails, `preserve_violation` | Preserve violation |
| V04 | `Provide two [Verify quantity] spare filters per unit.` becomes `Provide two filters per unit.` | **PASS**, `inline_placeholder` | fails, unexplained | Unexpected loss of `spare`; not PASS |
| V05 | Same phrase occurs in a visible run and hidden run; delete the visible occurrence | **PASS**, `formatting_based` | fails, unexplained | Unexpected loss unless the actual surviving sequence is provably equivalent with all protected text retained |
| V06 | Hidden detector disabled; delete hidden-marked text | **PASS**, `formatting_based` | fails, unexplained | No hidden-text permission to excuse it |
| V07 | Protected heading loses a fragment that does not itself match a preserve regex | fails, unexplained | fails, `preserve_violation` | Preserve violation based on the original paragraph |
| V08 | Several permitted fragments plus one unpermitted fragment disappear | fails, unexplained | unchanged | Whole modification verdict fails |
| V09 | Paragraph with all substantive text in eligible runs disappears | PASS | unchanged | Expected, provided no preserve or structural rule forbids the loss |
| V10 | Input has duplicate requirements; output has one fewer | fails, unexplained | unchanged | Missing multiplicity must be detected |
| V11 | Body requirement disappears but identical header text remains | fails, unexplained | unchanged | Header text must not excuse body loss |
| V12 | Reordered protected clauses | fails | unchanged | Not accepted as an expected deletion |

Two corrections to the V01–V04 rows as first written, found while building the fixtures. V03's original
text — `Select one of the listed manufacturers.` — stopped matching any removal rule once W02 narrowed
`select one`, which would have made the case prove nothing about preserve *precedence*; it needs text
that still matches a rule, so the fixture uses `Retain or delete manufacturers below.` And every case
needs an anchor paragraph present on both sides: substituting a placeholder for the damaged paragraph
makes the case fail on the invented text no matter how the loss is classified, and a tripwire that
fires for the wrong reason is not a tripwire.

V05 and V06 were "not yet observed" and turned out to be live defects, not hypotheticals. V08, V10,
V11 and V12 were already handled correctly.

V04 landed early, in W00, under `unittest.expectedFailure` so the suite stayed green until W03 removed
the decorator along with the defect (§5.2 and §7.4). It reported the unexpected success that said the
decorator could go.

For ambiguous identical text, do not claim source occurrence identity that the output format cannot
establish. The safety assertion is preservation of required text, multiplicity, order, and available
structural identity. Ambiguity must not be resolved by assuming the editorial occurrence was the one
deleted.

**Partially closed in W03, on the "available structural identity" clause.** A review of the W03
branch found the concrete case: a hidden note followed by an identical visible requirement extracts
the same characters twice, so removing the note leaves a survivor consistent with either occurrence
having gone. `difflib` assigns the deletion to the first, the second interval has no authority, and a
*correct* clean is reported as damage.

No text-level rule can resolve it, and the failed candidates are worth recording so they are not
retried. Accepting when cutting every authorized interval reproduces the output wrongly accepts the
mirror case — the visible requirement lost and the hidden copy surviving — because the two outcomes
produce **byte-identical extracted text**. The difference is only which run survived, which is
structural, exactly as this paragraph anticipated.

So W03 compares the output's own run signatures against the runs a correct clean would leave, and
accepts on an exact match. Signatures are pure document fact — text plus raw `w:rPr` properties, no
style resolution and no policy. The path only ever accepts; a mismatch falls through to interval
reasoning unchanged, so it cannot manufacture a false negative in a case it does not recognise.

**The residual is W04's.** An alignment that is ambiguous *and* not an exact structural match still
rests on whichever opcode assignment `difflib` chose. §11.2 already requires that such a region return
an actionable verification limitation counted in the ambiguous-alignment category of §10.6, rather
than being certified or reported as damage; that category does not exist yet, and W03 does not create
it.

### 10.6 False-alarm acceptance criteria

A verifier that routinely raises unexplained alarms loses its practical value, and invariant 15 makes
that a safety property rather than a usability preference. W03 is therefore not complete when the
injected-damage tests fail correctly; it is complete when the verdicts are also *usable*.

**A blanket target such as "≥95% Verified" is explicitly rejected.** It rewards weakening the checks —
the cheapest way to raise a Verified rate is to make the verifier more permissive, which is precisely
the failure this package exists to eliminate — and it obscures that some documents *correctly* require
review. The denominator must be known-correct outputs, not every run regardless of its contents.

Required criteria:

1. **Every independently judged correct output in the acceptance set is reported Verified.** No
   exceptions except individually understood and documented ones.
2. **No unexplained new failure relative to the recorded W01 baseline remains at completion.**
   Baseline-relative, in the same spirit as `_compare_structure`, which already counts only the issues
   the output added.
3. **Ambiguous alignment, genuine detected damage, configuration notices, and numbering/reference
   warnings are measured separately** and never pooled into a single rate.
4. **All injected-damage tests continue to fail for their stated reason** — not merely to fail.
5. **Any permitted exception is individually understood and documented**, naming the document, the
   category, and why the verdict is correct.

**Scope the denominator to sampled decisions, not whole documents.** Independently judging a complete
400-page consolidated manual is not tractable for one reviewer. Use the per-decision records the §8.4
harness already produces: every paragraph where baseline and candidate differ, plus a random sample of
removals. Same rigor, finite work.

**One non-gating requirement, carried into W07.** The same four-way category split in criterion 3 must
be surfaced in the GUI's Needs-review reason, not only measured during acceptance. It costs nothing
once the categories exist, and it is the only way to learn whether Needs-review is dominated by
ambiguous alignment in real use — which would be the signal to revisit this package.

**Exit gate:** F03–F06 fail for the right reason under deliberately damaged outputs; ordinary supported
cleans still pass; the criteria above are met and recorded; the contract is documented clearly enough
for the XML and performance work to use.

**What W03 could and could not establish, as delivered.** The first half of the gate is met and
recorded: F03–F06 each fail for their stated reason under hand-built damaged outputs, the twelve cases
are measured before and after in §10.5, and every existing test still passes — including
`tests/test_evidence.py`, which cleans each case for real and checks the source evidence predicted the
output.

The five acceptance criteria above **cannot be evaluated**, and were not. Every one of them takes its
denominator from an acceptance set of independently judged correct outputs on real documents, and no
specification corpus was available in this workspace — the same gap §8.4 and the W02 formatting-only
decision already record. What exists instead is the property those criteria were chosen to protect,
asserted on synthetic documents: V09 requires that a *correct* clean still passes, which is what stops
the contract being satisfied by a verifier that distrusts everything, and the whole pre-existing suite
is a standing false-alarm check, since every one of its cleans must still verify.

That is weaker than the criteria ask for and should not be described as meeting them. Criteria 1, 2, 3
and 5 stay open until this runs against real files; criterion 4 is met. The measurement is the same
one W01 built the harness for, and the honest reading is that W03's *classification* is evidenced
while its *false-alarm rate on real documents* is not.

## 11. W04: pairing, part identity, and ordering

W00 fixed the reproduced long-redaction case narrowly. W04 generalises it and adds the identity and
ordering guarantees that keep pairing honest.

### 11.1 Pair permitted transformations before fuzzy heuristics

Generalise §7.4 from the single reproduced case to the full set: for any source paragraph with eligible
inline or run intervals, compute the expected text for the intended supported removal combination
without mutating XML. An exact output match is a stronger pairing signal than character similarity.

Character-level diff output can help present losses after a pair is established. It cannot be the sole
authority for whether an edit was a permitted deletion — repeated characters can make an unconstrained
diff choose misleading fragments, which is the mechanism behind F06.

The W00 constraints still bind: do not fix pairing by reducing `MIN_PAIR_SIMILARITY`, accepting all
short survivors, dropping `added` from the pass predicate, or suppressing unexpected changes.

### 11.2 Identity and ordering

Compare within matching package parts. `extract_paragraphs()` currently returns one flat list across
every content part, which is what makes invariant 6 violable today. Preserve relevant story
identities, particularly footnote and endnote IDs, so a missing note cannot borrow an identical
paragraph from another note. Use full package-relative part names; `word/document.xml` and
`word/glossary/document.xml` must not collapse to the same identity.

Preserve monotonic source/output ordering and multiplicity. Exact duplicate text does not establish a
unique source position. Do not use mutable paragraph indices or generated XPath positions as if they
were stable IDs across removals.

Begin with exact unchanged anchors and exact permitted-transformation candidates. Keep any fuzzy
fallback conservative and constrained by the neighboring established alignment. If the verifier cannot
resolve an ambiguous region safely, return an actionable verification limitation rather than
certifying it — and count it in the ambiguous-alignment category of §10.6, not as damage.

W03 closed the exactly-resolvable half of this: where the output's run signatures match the runs a
correct clean would leave, the ambiguity is settled on evidence (see §10.5). W04 owns what is left —
the region that stays ambiguous after that check, which today falls back to `difflib`'s opcode
assignment and is reported as damage when the assignment happens to land outside an authorized
interval. Creating the ambiguous-alignment category, so those are counted apart from genuine damage
rather than pooled with it, is part of this package and not of W03.

### 11.3 Required cases

- The exact long-placeholder example becomes `Provide units.` with one expected modification and no
  addition or unexpected removal.
- A paragraph reduced to a short but substantive survivor is recognized; a placeholder-only paragraph
  still follows whole-paragraph rules.
- Multiple placeholders, split runs, leading/trailing whitespace, and repeated words retain correct
  source intervals.
- Adjacent similar paragraphs cannot exchange survivors to conceal an unexplained removal.
- Duplicate unchanged headings and requirements preserve multiplicity.
- Multiple changed paragraphs between two unchanged anchors are handled deterministically.
- Empty documents and parts, insert-only output, missing parts, and unchanged documents receive
  explicit outcomes.
- The same text in different parts or notes does not mask a loss.

**Exit gate:** supported long redactions pass without weakening injected-damage tests or changing the
verdict for real additions and substitutions; part and story identity is carried through comparison.

### 11.4 What W04 found, and the one thing it did not build

Four defects, each measured against a real clean before anything was changed, and each now pinned in
`tests/test_identity.py`:

| | Defect | Kind |
|---|---|---|
| 1 | A requirement deleted from the body verified clean, because an identical *hidden* paragraph in a header carried authority the body's copy did not | **false negative** |
| 2 | Removing a hidden note beside an identical plain requirement was blamed on the plain one | false alarm |
| 3 | Two paragraphs with character-identical runs differing only by an editorial *paragraph style*: removing the styled one was blamed on the plain one | false alarm |
| 4 | A paragraph carrying both an authorized run and a placeholder had no expectation to match, so the differ chose, and blamed the wrong copy of a repeated run | false alarm (3 shapes) |

Defect 1 is invariant 6 and the reason §11.2 exists. The other three are the same mistake seen from
different sides: **a text match is not an identity**. Comparison now happens inside a location, and
within one, which occurrence disappeared is decided by paragraph signature — style and runs — rather
than by `difflib`'s alignment.

The ten §11.3 redaction shapes (repeated words, split runs, multiple and adjacent placeholders,
leading and trailing whitespace, a long redaction leaving a short survivor) were measured and all
already passed, so §11.1 needed no separate work beyond defect 4's generalisation.

**The ambiguous-alignment category was not built, and that is a deliberate finding rather than an
omission.** §11.2 assigns it to W04, on the reasoning that a region the verifier cannot resolve should
be counted apart from genuine damage. After the structural resolution above, no such region could be
constructed: every ambiguous case built for this package is now *determined* by the run and paragraph
signatures.

The attempt is worth recording because it nearly went wrong. The natural heuristic — treat a loss as
ambiguous when the lost text occurs more than once and some occurrence is authorized — would have
relabelled **V05** as ambiguous. V05 is not ambiguous: the profile mismatch positively establishes that
the *hidden* copy survived and the visible requirement was lost. It is determined damage, and a
category that softened it would have weakened an injected-damage test to add a number.

So the category stays unbuilt until a case demands it. The two places it would be reachable are known
and recorded here: a paragraph whose run offsets fail to reconstruct its text (`offsets_reliable`
false, a defensive guard with no reachable trigger today), and a partially-cleaned output where the
cleaner removed only some authorized runs — which the cleaner does not produce, so it arises only from
an output that is already damaged and correctly reported as such. If corpus evaluation later shows
Needs-review dominated by alignment rather than by findings, this is the answer, and §10.6 criterion 3
still asks for the measurement.

## 12. W05: inline XML redaction and field carriers

W00 closed the reproduced separator case. W05 closes the rest of the inline-redaction surface and the
simple-field carrier defect.

### 12.1 Remaining redaction cases

Extend the W00 element-aware redaction to the cases it did not reach:

- A placeholder crossing a hyperlink or content-control wrapper.
- The same run scheduled for both redaction and removal.
- Nested text boxes, which must remain independently processed.
- A placeholder split across multiple `w:t` nodes and wrapper boundaries.

Re-confirm the W00 break policy holds under these cases: a `w:br` carrying page or column semantics is
preserved, and any resulting uncertainty is reported rather than silently resolved.

### 12.2 Preserve simple fields

Recognize `w:fldSimple` as a structural field carrier. `EMBEDDED_CONTENT_TAGS` currently covers
`w:fldChar` and `w:instrText` but not `w:fldSimple`, and `field_chars_balanced()` counts only
`w:fldChar`, so a paragraph whose only field is simple has no carrier protection at all — reproduced as
a clean 1 → 0 loss with verification passing.

Keep the instruction attribute and required wrapper during ordinary editorial text stripping. Include
nested simple fields and fields under wrappers in focused tests.

Do not remove field instructions or update field values programmatically. Blank only text the supported
policy authorizes removing. Where field results are preserved or emptied, document that Word may
recalculate them; keeping a field wrapper is not a promise that a removed editorial field result stays
absent after refresh.

Strengthen verification enough to catch loss or alteration of a supported field carrier even when its
cached text is editorial and matches a removal pattern. Compare field evidence by part and
instruction/structure, not by a document-wide field count. Allow supported changes only where accepted
revisions legitimately remove that field and the source revision evidence explains it.

### 12.3 Required fixtures

| Test | Required behavior |
|---|---|
| X01 | Placeholder containing `w:noBreakHyphen` → output is `Provide units.` with no orphan `-` (lands in W00; pinned here) |
| X02 | Placeholder containing supported tabs and line breaks → covered separators disappear; separators outside the span remain; page/column breaks are preserved per the W00 rule |
| X03 | Placeholder split across multiple `w:t` nodes and wrapper boundaries → exact surviving requirement text and balanced XML |
| X04 | Redaction empties a run also scheduled for removal → no double-removal error or loss of neighbors |
| X05 | Editorial paragraph contains `w:fldSimple w:instr="REF Target"` → field survives ordinary removal processing |
| X06 | Deliberately damaged output removes that simple field → verification flags structural/reference loss |
| X07 | Field instruction/cached content in a nested text box → outer processing does not double-count or strip nested requirement text |

**Exit gate:** exact text assertions and XML-carrier assertions pass; no field, section, or picture
preservation regression; Word validation cases are queued for W10.

### 12.4 What W05 found

**§12.1 needed no new work.** X01–X04 were measured against real cleans first and all four already
behaved correctly — W00's element-aware redaction reaches a placeholder crossing a hyperlink wrapper,
one split across `w:t` nodes, and a run scheduled for both redaction and removal, and the page-break
policy holds under each. X01 and X02 were already pinned; X03 and X04 are pinned now.

**§12.2 was the whole package, and reproduced exactly as written.** `has_embedded_content()` returns
True for a complex field and False for the character-identical simple one, so an editorial paragraph
whose only field was simple was deleted whole, taking a live cross-reference with it — and because the
paragraph's text was exactly what the rule asked to remove, verification passed. X06 confirmed the
second half: with the text unchanged and the carrier stripped, verification reported nothing at all.

The two halves are independent and were fixed independently, which the tests check by disabling each
in turn: removing `w:fldSimple` from `EMBEDDED_CONTENT_TAGS` fails three tests, and neutering the
field comparison fails three others.

Fields are compared by instruction and **per part**, per §12.2, not by a document-wide count — the
same reasoning as W04's locations, arrived at for the same reason. Instructions are normalised because
Word splits a complex one across `w:instrText` nodes at arbitrary points, so ` REF ` + `Target ` and
` REF Target ` are one field written two ways.

No field instruction is rewritten and no field value is updated programmatically, per §12.2. The
caveat that section asks for is documented in `CLAUDE.md` and the changelog: keeping a wrapper is not
a promise about its value, because Word recalculates on refresh.

**Review round.** Two further defects, both reproduced before fixing, and both in the inventory this
package introduced:

- *Nested fields collapsed into one entry.* A field inside another field's result is its own carrier.
  Accumulating instructions into one buffer and emitting when the nesting closed merged them, so an
  output that lost the inner field's `w:fldChar` pair — keeping its instruction text and cached
  result — produced an **identical** inventory. A stack fixes it.
- *`w:moveFrom` was left out of the accepted-revision exemption.* `accept_revisions()` removes both
  `w:del` and `w:moveFrom`, but the ancestry check named only `w:del`, so a field inside an accepted
  move was reported lost from a correct run.

The second exposed a **pre-existing** false alarm one level up, unrelated to fields: `w:moveFrom` is
the only revision whose content reaches the *text* comparison, because a deleted run hides its text in
`w:delText` that no extractor reads while the source half of a move keeps real `w:t`. So accepting a
move reported its paragraph as an unexplained removal whatever it contained. Fixed here rather than
deferred, because the finding's stated symptom — a correct revision-accepting run requiring review —
persists until both halves are closed.

Both fixes are the same lesson as the first half of this package: **a second, shorter list of an
answer drifts from the first.** `in_tracked_deletion()` now reads `REVISION_DELETE_TAGS` rather than
naming a subset of it.

## 13. W06: reference integrity, revision-empty tables, and numbering notices

### 13.1 Referenced bookmark targets — detect and report first

**This is the package revision 2 rescoped most.** Revision 1 selected conservative retention of a
referenced target's content. Revision 2 defers that.

The reason bookmarks and numbering need not receive identical treatment is real and worth stating: a
surviving `REF` names a **specific** missing target, so a newly broken reference is a decidable fact,
whereas numbering consequences involve broader implicit relationships that resist precise attribution.
That argues for **more precise detection** of reference breakage. It does not, by itself, justify
automatic retention — which has an implementation cost and an effect on cleaning volume that revision 1
never bounded.

**Selected behavior for this version:**

1. Build a bounded source inventory of bookmark names and ranges, and of supported internal consumers.
   At minimum cover simple-field and complex-field `REF` references and internal hyperlink anchors.
   Evaluate `PAGEREF` and `NOTEREF` against the documented field grammar before claiming support.
2. Detect a reference that is **newly** broken by this clean — a target present in the input and absent
   from the output while a supported consumer still names it. Reproduced: a bookmark fully inside a
   removed paragraph disappears while its `REF` field survives, and verification passes.
3. Report it as **Needs review**, naming the part, the bookmark, and the consuming field.
4. **Preserve the original.** The cleaned output is still written, per §3.2.
5. Distinguish a newly broken reference from one already broken in the input. Only what the clean broke
   is the clean's fault, matching how `_compare_structure` already treats structural issues.

Leaving an empty bookmark behind does **not** count as fixing reference integrity: it suppresses one
error message while letting a `REF` field return misleading content.

Automatic retention of referenced target content is **deferred** to a later, separately justified
decision (§3.3), gated on census B (§8.3). If census B shows that a large share of removable paragraphs
sit inside referenced ranges, retention would suppress cleaning broadly and needs a narrower design; if
it shows a small share, retention becomes a cheap follow-up. Either way the measurement comes first.

Note that protecting only *referenced* ranges while not protecting *all* bookmarks is coherent, not
contradictory (§2.6) — the open question was never logical consistency but scope and cost.

For malformed or unsupported reference forms, report the limitation and identify the implicated target
where possible. Do not invent a target, retarget a field, create a replacement bookmark, or rewrite
field results. For complex fields, account for instruction text split across runs and quoted bookmark
names. Do not scan ordinary visible prose for the word `REF` and treat that as a field. Reuse a small
field-instruction reader rather than implementing a general Word field evaluator.

Tests must cover complete and half-open ranges, multi-paragraph targets, simple and complex references,
quoted names, internal hyperlinks, unreferenced bookmarks, existing broken references, and bookmarks
inside explicitly accepted tracked deletions. If accepting a revision deletes a referenced target,
report that consequence as requiring review even though the text deletion itself was requested.

### 13.2 Tables emptied by accepted revisions

Trace `_accept_structural_deletions()` when deleting a last row or last cell. It already removes a row
left with no cells, but nothing handles a `w:tbl` left with no rows — reproduced as rows 1 → 0 with the
table element still present, structural lint reporting clean, and verification passing.

Preferred behavior: when accepted revisions remove all rows of a table, remove the now-empty table and
restore only the minimum required block content in its parent where necessary. Preserve surrounding
paragraphs, required terminal cell paragraphs, section properties, and unaffected tables. Do not insert
fake empty rows as a blanket repair.

Handle nested tables from the inside outward so removing an inner structure does not invalidate the
outer traversal. Inspect container context before removing a table that is the only block in a header,
footer, note, text box, cell, or body.

A table removed by an accepted revision is an explained structural change, not an unexplained loss. A
newly introduced zero-row table or zero-cell row must be observable by the verifier — today it is not.
Validate the schema and Word assumptions for the supported representation, and **report Word behavior
separately from custom lint results**: this case is the live demonstration that a clean lint is not
evidence Word will accept a package.

Test last-row deletion, deletion of every cell in a row, multiple deleted rows with one survivor,
nested tables, sole-table containers, preserved headers and footers, and revision stripping disabled.
Keep the current policy of not merging deleted paragraph marks unless a separate change is justified.

### 13.3 Numbering-sensitive removals

Do not attempt automatic renumbering or cross-reference rewriting in this work.

Add a focused review notice when a removed paragraph participates in automatic numbering, using direct
`w:numPr` and supported inherited paragraph-style evidence. Read the relevant source metadata once
rather than parsing numbering separately for each removal. An inherited value explicitly disabling
numbering must not be treated as participation; validate the OOXML semantics before implementation.

The notice must say that displayed numbering or paragraph references may change; it must not assert
that a reference definitely broke. Include part, location, and a short source preview. This contributes
to a **Needs review** outcome in its own category (§10.6 criterion 3) without falsely claiming a
text-integrity failure.

If inherited numbering support cannot be implemented reliably within the small metadata extension, ship
accurate direct-numbering detection with an explicit limitation and a tracked follow-up. Do not claim
complete numbering awareness.

**Exit gate:** newly broken supported references are detected and reported with the original preserved;
no new empty-table residue is produced; numbering warnings are precise about what was detected and what
remains uncertain; each of the three concerns landed as a separate reviewable change.

### 13.4 What W06 found

All three concerns reproduced exactly as written, each measured against a real clean before anything
was designed, and each landed as its own commit.

| | Reproduced | Now |
|---|---|---|
| §13.1 | bookmark gone, `REF` surviving, `lint=[]`, `verify passed=True` | reported, run needs review |
| §13.2 | `tbl=1 tr=0` after accepting, `lint=[]`, `verify passed=True` | table removed; both empty shapes linted |
| §13.3 | numbered paragraph removed, no notice of any kind | its own notice category |

**Inherited numbering was implemented, not deferred.** §13.3 allowed shipping direct-only detection
with a stated limitation if the metadata extension proved unreliable. It did not: `StyleInfo` already
walks a `w:basedOn` chain for `w:vanish`, and reading `w:pPr/w:numPr/w:numId` alongside it is the same
shape of work. `w:numId` `"0"` is handled as the override it is, which is the case §13.3 specifically
warns against getting wrong.

**One deliberate narrowing, stated rather than silent.** A notice is raised only when the removed
paragraph's numbering list still has surviving members. A list whose every paragraph went renumbers
nothing, so a notice about it would assert a consequence that cannot occur — and the exit gate asks
for warnings that are *precise about what was detected*. This is narrower than "participates in
automatic numbering" read literally, and it is narrower in the direction the gate points.

**What is not measured.** Numbering notices contribute to a Needs-review outcome, per §13.3. How often
that fires on real specifications is unknown: MasterSpec numbers heavily, and if it also numbers its
specifier notes this will fire on most files. That is exactly what §10.6 criterion 3 asks to be
counted separately, and the category now exists to count it — but no corpus was available to run it
against, the same gap recorded for the W02 default and the W03 acceptance criteria.

Automatic retention of referenced target content remains deferred to census B, per §13.1 and §3.3.
Nothing here repairs a reference, renumbers a list, or rewrites a field.

**Review round.** Three findings, one against each commit, all reproduced before fixing:

- *A report named the bookmark and nothing else.* With one target referenced from the body, a header
  and a footnote it collapsed to a single line — what was wrong, not where. References are now kept
  per consumer and each is reported with its part and with the instruction or anchor naming it.
- *Every rowless table was removed, not only the ones this run emptied.* A document with **no
  revisions at all** was rewritten because the option was on. That is the cleaner changing a file for
  a reason unrelated to its task, and the most serious of the three. Only tables an acceptance took a
  row or cell from are considered now; a pre-existing one stays linted but unrepaired, which is the
  rule the structural comparison already follows.
- *Numbering ignored the default paragraph style.* `paragraph_style()` answers None for a paragraph
  naming no style, so the chain examined nothing and a default style supplying `w:numPr` was
  invisible.

The default-style lookup was deliberately confined to numbering. The same gap exists for hidden-text
and editorial-style resolution, but closing it there changes what the cleaner *removes* rather than
what it *reports* — a different kind of change, belonging in its own package with its own fixtures.

## 14. W07: full batch and outcome model

W00 delivered the manifest rejection and a minimally honest three-way report. W07 completes the model
now that W03 and W06 have defined the result types it needs.

### 14.1 Model file outcomes explicitly

Use a small enum or dataclass rather than one Boolean that conflates everything:

| Outcome | Meaning | Count in summary |
|---|---|---|
| Verified | Output written, verification completed and passed, no actionable review notice | Verified |
| Needs review | Output written, but verification failed or an actionable structural, reference, numbering, or configuration concern remains | Needs review |
| Failed | Processing failed, or verification could not complete | Failed; state whether an output exists |

**Needs review must name its category** — ambiguous alignment, detected damage, configuration notice,
or reference/numbering warning — per §10.6's non-gating requirement and invariant 15. A user who
cannot tell which of those four they are looking at cannot act on it, and a verdict that cannot be
acted on will be ignored.

If verification raises after a successful write, report that the output exists but is unverified. Do
not hide the path or claim that nothing was produced. Continue processing other independent files as
the current GUI does.

Example summary:

```text
Done: 8 verified, 2 need review, 1 failed.
```

Keep logs queued, widget updates on the main thread, controls disabled during the run, and inputs
snapshotted before worker execution. Every exception path must restore controls through the established
mechanism.

### 14.2 Tests without opening a window

The GUI tests skip where `tkinter` is unavailable, which is every Linux run — the recorded baseline is
92 passed with exactly those 4 skipped. **Skipped GUI tests establish nothing about Windows GUI
behavior.** Run this package's tests on Windows before accepting it.

- Same source basename from two folders with one output directory: manifest rejected; processor never called.
- Distinct names: both pairs accepted.
- Case-only destination collision on Windows: rejected.
- Output equals a selected source or its equivalent existing path: rejected.
- Previously existing output: original confirmation behavior still applies.
- Passing verification: verified count increases.
- Failed verification with an existing output: needs-review count increases, category named, path logged.
- Verification exception after output creation: failed/unverified outcome, path logged.
- Processing exception: batch continues, final counts accurate.
- Source/destination controls changing after start cannot alter the validated manifest.

**Exit gate:** no planned intra-batch overwrite is possible; a failed verification cannot contribute to
the verified count; every Needs-review outcome names its category; output disposition is explicit; the
package's tests have actually executed on Windows.

## 15. W08: configuration fallback, validation, and CI

### 15.1 Use the shared resolver

In `verify_clean()`, use `apppaths.resolve_config_path()` only when both an engine and an explicit
config path are absent. Supplying an engine should not trigger irrelevant fallback resolution or
first-run configuration seeding. An explicit config path takes precedence over automatic resolution
when no engine is supplied.

**This defect is latent** (§2.3, F16): the GUI always supplies `engine=`, so the current fallback line
is computed and never read, and it governs no production run today. Fix it because an engine-less
caller — a test, the §8.3 census utilities, a future harness — would silently read the wrong
`patterns.yaml` in a frozen build. Do not describe it as a live misconfiguration.

Preserve and document the existing engine-versus-path precedence when both are supplied; avoid a silent
behavioral change. Test it explicitly.

Test source mode, executable-adjacent configuration, existing per-user configuration, first-run
seeding, read-only fallback, explicit path, and engine-supplied paths. Reuse the frozen-build
simulation already in `tests/test_apppaths.py`; do not require an executable build for every path test.

### 15.2 Focused config checks

Validate fields touched by the change: Boolean options must be Booleans, relevant lists must contain
strings, and editorial colors must have the supported hexadecimal shape. Decide whether to normalize a
leading `#` or reject it with a useful message; document and test one consistent behavior. Do not
accept the string `"false"` as a truthy switch.

Do not reject style names simply because they are absent from one input document; styles differ across
templates. Avoid broad unrelated validation changes.

### 15.3 CI coverage

`release.yml` is the only workflow, and its pull-request filter lists `*.py`, which matches root files
only — so a pull request touching just `tests/**` triggers nothing at all. The workflow does run the
suite when it fires; the gap is that it does not fire.

Add a small test workflow on pull requests and normal branch pushes, with read-only repository
permissions and no release publishing side effects. At minimum test on Windows, which exercises Tkinter
and relevant path behavior — and which is the only way the four skipped GUI tests ever run. A Python
3.10 compatibility lane and the current supported Windows Python version are useful; choose a small
matrix rather than an expensive exhaustive one.

Keep the release build workflow for packaging validation. Add `tests/**` to its pull-request filters, or
deliberately document why test-only changes use the dedicated test workflow without packaging. Do not
leave test-only changes with no CI.

No new test framework is needed. The test workflow runs `python -B -m unittest discover -s tests -t .`.
Keep large timing benchmarks out of ordinary CI; small deterministic correctness and work-count
regressions belong there.

This is the one package that may run in parallel with application work (§6): it touches no application
module.

**Exit gate:** the verifier uses the intended active config under every call mode; malformed touched
options fail clearly; test-only changes receive test coverage without publishing anything.

## 16. W09: measured performance improvements

### 16.1 Establish a corrected baseline

Benchmark after W02–W06 so timing improvements are not confused with changed cleaning policy. Measure
baseline and candidate on the same machine, Python version, fixtures, and active configuration. Profile
separately from wall-clock timing.

Create a developer utility, for example `tools/benchmark_pipeline.py`, that uses synthetic fixtures and
the existing APIs. This is a different tool from the §8.4 corpus harness; do not conflate them. Keep
outputs in an explicitly chosen developer directory. Record paragraph counts, relevant pattern
distribution, cleaning time, verification time, output text digest, outcome counts, and
interpreter/version information.

Do not compare ZIP bytes as the behavior oracle. Compare extracted semantic content, relevant XML
carrier and reference evidence, and expected verification results.

Fixture families:

1. The original repeated-heading distribution at 500, 2,000, 6,000, 12,000, and 24,000 paragraphs where
   practical. **This is the adversarial case, not the representative one** (§2.4).
2. Mostly unique requirements with sparse editorial removals — the controlled contrast that isolated
   the duplicate effect, and the closest thing here to a realistic distribution.
3. Repeated identical headings and repeated identical requirement paragraphs.
4. **Long sequences of changed paragraphs without unchanged anchors.** This family is now load-bearing:
   it is precisely where a unique-anchor approach has no anchors to work with, and it is the reason
   duplicate-bucket evidence does not by itself establish patience/LIS as the answer.
5. Many inline placeholders, including short survivors.
6. Very long individual paragraphs, kept separate from paragraph-count scaling.
7. Tables and multiple content parts.
8. Deliberately damaged output with unexplained losses; faster verification must still fail.

Use repeated unprofiled runs, reporting a median and input/setup boundaries. **Do not extrapolate a
universal complexity law from a few timings.** The measured growth is super-quadratic and accelerating
over the duplicate-heavy range; that is the claim the data supports and the claim to make.

### 16.2 Main paragraph alignment

Optimize the paragraph-level alignment at the current `verify.py` outer matcher first — it accounted
for roughly 97% of profiled verification time at 6,000 paragraphs, against 0.97% for survivor pairing.

Candidate approach:

1. Partition by part and story identity — which W04 established, and which is a correctness requirement
   before it is a performance one.
2. Establish uniquely matching exact unchanged anchors in monotonic order, using a patience/LIS-style
   approach or another justified deterministic method. **Admit exact permitted transformations as
   anchor candidates here as well**, not only inside bounded regions: a paragraph whose output exactly
   matches its independently computed expected redaction result is a reliable pairing, and it is unique
   in exactly the documents where duplicate text otherwise destroys the anchor set (§5.2).
3. Align the bounded regions between anchors, preferring exact permitted transformations from W04.
4. Handle duplicate occurrences explicitly, preserving multiplicity and order.
5. Apply conservative fallback only to unresolved regions and record any inability to verify fully.

This is a suggested algorithm, not permission to weaken the comparison, and the duplicate-bucket
evidence does not make it the proven answer — fixture family 4 exists to test the case it handles
worst. A simpler approach is preferable if it passes the same adversarial tests and measured workloads.

Do not turn `autojunk` on blindly, use a bag-of-paragraphs comparison, limit candidate windows without
accounting for excluded possibilities, or skip large ambiguous regions while returning PASS. If a
resource bound is needed, surface an incomplete-verification outcome with the unresolved location, and
count it in the ambiguous-alignment category — that is a fallback limitation, not completion of the
performance objective.

Remove the redundant character-level matcher only after correctness is settled, and note it is worth
under 1% of the profile. For an already established pure deletion, the similarity ratio can be derived
from lengths. Length prefilters are useful where mathematically justified, but they no longer apply as
a universal rejection rule to exact permitted long redactions — that rule is what produced F07.

### 16.3 Container scan optimization

`can_delete_paragraph()` materialises the parent's full block-child list per candidate removal. Replace
that only if the replacement remains easy to audit. Consider an early-exit scan, cached per-container
block metadata updated during the removal phase, or targeted sibling traversal. Choose based on
profiling and clarity — and note this is secondary to §16.2, not a substitute for it.

The replacement must preserve behavior with:

- A sole paragraph.
- Several removable paragraphs processed consecutively.
- Tables and paragraphs interleaved.
- Non-block markers before, between, and after blocks.
- A table cell that must end in a paragraph.
- Section properties at the end of a body.
- Nested containers and wrappers.
- Earlier mutations that change sibling relationships.

Do not cache counts across mutations without updating them. Do not claim constant-time behavior if the
implementation still walks an unbounded sibling chain.

### 16.4 Acceptance and stopping rule

- All corrected correctness tests and injected-damage tests pass.
- Report before/after medians at several sizes and explain where the gain comes from.
- For the repeated-heading 12,000/24,000-paragraph family, aim for a substantial multi-fold improvement
  and growth well below the prior super-quadratic behavior. Treat these as review targets, not
  machine-independent hard seconds in CI, and **do not gate solely on the adversarial family** —
  fixture family 2 is the closest thing to a realistic distribution and must not regress.
- A material improvement only to character pairing does not satisfy the main bottleneck objective; it
  is worth under 1% of the measured profile.
- No meaningful small-document regression should be introduced; distinguish millisecond noise from an
  actual regression.
- Add a deterministic small regression for the algorithm's problematic pattern, using work counts or a
  bounded fixture rather than fragile wall-clock assertions.
- Stop after the measured hotspots are resolved. Do not then rewrite all ZIP I/O, add concurrency, or
  optimize unrelated allocations without new evidence.

**Exit gate:** large-input verification is measurably better on the corrected behavior across at least
the adversarial and realistic families, and no safety property was traded away to obtain it.

## 17. W10: integrated validation and documentation

### 17.1 Automated validation layers

Use three distinct layers:

1. **Focused unit/algorithm tests:** interval eligibility, path manifests, source metadata, and
   alignment edge cases.
2. **Synthetic package integration:** build actual DOCX ZIPs, run the public processor and verifier
   path, reread output XML, and assert text and structural behavior. The second review confirmed this
   layer is achievable for every reproduced finding — the reproductions in §2.3 were performed exactly
   this way — so there is no excuse for leaving a finding tested only in memory.
3. **Injected-damage verification:** independently construct damaged outputs so agreement with the
   cleaner cannot make a broken verifier appear correct. Owned by the reviewer (§6).

Do not mock the parser, writer, or structural inspection in the final package integration tests.

Extend `tests/docx_builder.py` narrowly. Check that any fixture used for Word validation has complete
enough namespaces, relationships, and field/bookmark definitions to be a valid positive control.
Existing minimalist fixtures are not automatically valid Word compatibility fixtures — for example,
verify that every namespace named by `mc:Ignorable` is actually declared.

Run the full suite after integration, and once more only if subsequent edits justify it. Record
interpreter, platform, test count, skips, expected failures, and failures against the W01 baseline.
An expected failure is not a failure and an unexpected success is; record them as distinct counts.
GUI tests skipped on
Linux do not establish Windows GUI behavior.

### 17.2 Corpus comparison

Rerun the §8.4 harness against the candidate and diff it with the baseline recorded in W01. For each
changed decision record:

- Source file and part/location.
- Original paragraph or a permitted private preview.
- Baseline action and candidate action.
- Rule/evidence explaining the difference.
- Whether the change is an intended conservative retention, a confirmed precision improvement, or unresolved.

Review all newly removed substantive content. Narrowing rules may intentionally increase retained
editorial text; record this as a recall tradeoff rather than assuming it is a regression. Use a private
temporary output directory; do not commit proprietary specifications or excerpts without the
maintainer's authorization.

This is also where §10.6's acceptance criteria are evaluated on their proper denominator: sampled
decisions from this diff, plus a random sample of removals — not whole documents.

If no real corpus is available, finish all synthetic and code work and label corpus validation
outstanding. Do not fabricate precision/recall percentages or halt unrelated implementation solely
because corpus data is absent.

### 17.3 Word validation

Open original and cleaned copies of representative affected documents in Windows Word. Validate:

- No new repair or unreadable-content prompt.
- Expected requirement text, headings, and section boundaries remain.
- Headers, footers, note anchors, pictures, and fields still behave as intended.
- Simple fields survive; refresh references in a disposable copy and inspect the result.
- Referenced bookmark targets: newly broken references were reported, and no *unreported* new broken
  reference appears.
- Last-row and last-cell revision cases produce an acceptable layout with no residual broken table.
  **This is the case that most needs Word rather than lint** — the zero-row table passes the custom
  structural check today, so lint agreement proves nothing here.
- Automatic numbering changes are either absent or clearly reported for review.
- Saving and reopening a disposable cleaned copy does not introduce new problems.

Distinguish pre-existing issues from new ones by opening the original control. Do not save over the
original. If Word is unavailable, report manual validation as outstanding and do not claim full Word
compatibility. This is a validation limitation; do not invent a passing result.

### 17.4 Documentation edits

Update `README.md` and `CLAUDE.md` for every intentional behavior change. Preserve useful existing
explanations rather than trimming them wholesale; length is acceptable, invention is not.

Required topics:

- Safer pattern scope and examples of ambiguous content now retained.
- The formatting-only decision, its evidence, the opt-in path, and the installed-configuration transition.
- Source-evidence verification and its remaining dependence on configured policy. `CLAUDE.md`'s
  description of confidence scoring and `verify.py`'s module docstring both need to match the new
  contract.
- Exact inline transformation pairing and part-aware comparisons.
- New reference, field, and table safeguards, and the limits of structural checks — specifically that a
  clean structural lint is not evidence Word will open a file.
- Numbering notices without a promise of automatic cross-reference repair.
- Verified/needs-review/failed outcomes, the four Needs-review categories, and the disposition of
  written outputs.
- Batch collision rejection.
- Correct test instructions: synthetic fixtures are generated programmatically, but package tests write
  temporary files, and GUI tests skip without `tkinter`.
- Updated CI coverage, the census utilities, and developer benchmark instructions.

Correct contradictory statements encountered in touched sections, including claims of guaranteed intact
structure, tail handling, configuration location in frozen builds, and the breadth of what a PASS
proves. Do not rewrite unrelated documentation for style alone.

Add changes under an Unreleased section in `CHANGELOG.md`, following its existing Keep a Changelog
structure. Do not rewrite historical release entries to describe new behavior and do not invent a new
release version or date. **Update `requirements.txt` only if dependencies actually change; the default
expectation is no dependency change**, and nothing in this plan requires one.

### 17.5 Final implementation report

The implementer and reviewer must jointly provide:

1. Completed work-package IDs and any explicitly deferred items.
2. A concise description of final behavior and intentional compatibility/default changes.
3. Automated test results and the precise coverage of injected-damage and package integration tests,
   compared against the W01 baseline including platform and skips.
4. Before/after performance results with fixture, environment, and timing boundaries, across at least
   the adversarial and realistic fixture families.
5. Census A and B results, and corpus findings, each separated from synthetic results.
6. The §10.6 acceptance criteria, evaluated, with every permitted exception documented individually.
7. Word validation performed, or the exact outstanding manual checks.
8. Active-configuration transition behavior for installed users.
9. Remaining risks and known unsupported reference/numbering/OOXML cases.
10. Files changed and how each group relates to the implementation objective.

No release tag, push, publication, or deployment is part of this completion report.

## 18. Consolidated acceptance checklist

### W00 — immediate corrections

- [ ] No intra-batch destination collision is reachable; the batch is rejected before any write.
- [ ] A planned output can never equal a selected input, under Windows path equivalence.
- [ ] A failed verification is never reported as success.
- [ ] GUI wording no longer claims more than the checks establish.
- [ ] A separator covered by a redaction is removed; `Provide -units.` cannot recur.
- [ ] Page and column breaks are preserved, with the rule tested explicitly.
- [ ] The reproduced long redaction pairs correctly, without lowering the similarity threshold.
- [ ] The F06 injected-damage test is present under `unittest.expectedFailure`, named, and attributed
      to W03; the suite is green with one expected failure, not red.

### Content and verification

- [ ] F01/F02 requirement examples survive under the shipped defaults.
- [ ] Unambiguous copyright and editorial positives still clean.
- [ ] The formatting-only default decision is recorded with its census evidence, or with an explicit
      statement that none was available.
- [ ] Explicit opt-in for formatting-only removal still works, and its four interactions are tested.
- [ ] Preserve styles protect complete loss and partial modification.
- [ ] Low-confidence prerequisites and detector switches govern verification.
- [ ] Partial hidden/editorial evidence cannot excuse losing unrelated requirement text.
- [ ] Inline patterns cannot excuse extra deleted requirement words.
- [ ] Long valid redactions pair correctly without a similarity loophole.
- [ ] Duplicate text, part identity, and output order do not conceal unexplained loss.
- [ ] Source-based evidence is independently obtained from actual input and output files.
- [ ] The §10.6 false-alarm criteria are met on sampled decisions, with no blanket percentage target used.

### XML and reference behavior

- [ ] Covered separator elements are handled according to the documented redaction policy.
- [ ] Simple-field instructions and wrappers survive ordinary editorial removal.
- [ ] Loss of a supported field carrier is visible to verification.
- [ ] Newly broken supported internal references are detected and reported as Needs review.
- [ ] Existing broken references are distinguished from newly broken references.
- [ ] Census B has been recorded, and any decision to defer or pursue automatic retention cites it.
- [ ] Revision-empty tables do not leave unsupported zero-row residue, and the residue is observable
      by the verifier.
- [ ] Required parent block content and terminal cell paragraphs remain valid.
- [ ] Numbering-sensitive changes receive accurate review notices; no automatic renumbering occurs.
- [ ] Existing section, field, image, note, wrapper, and text-box tests still pass.

### GUI, configuration, and delivery

- [ ] Every Needs-review outcome names which of the four categories it belongs to.
- [ ] Written-but-unverified outputs are explicitly identified by path.
- [ ] Worker exception paths restore controls and continue independent files where appropriate.
- [ ] GUI tests have executed on Windows, not merely skipped on Linux.
- [ ] Verifier configuration fallback uses `apppaths` only when needed; F16 is described as latent.
- [ ] Existing user configuration is never silently overwritten.
- [ ] Superseded active defaults and the formatting-only opt-in are surfaced accurately.
- [ ] Test-only changes run CI.
- [ ] Performance improvement targets the measured outer paragraph alignment, and both the adversarial
      and realistic fixture families are reported.
- [ ] Documentation matches final behavior; historical changelog entries remain historical;
      `requirements.txt` is unchanged unless a dependency actually changed.
- [ ] Full tests, census, corpus coverage, and Word validation are reported honestly and separately.

## 19. Ready-to-use assignment text

### Implementing agent

Implement this plan against the current SpecCleanse checkout, sequentially, in package order. Start
with **W00** — the four corrections that depend on nothing else — and land each as a separately tested
change. Then take W01's baseline and census measurements before making the policy decisions in W02 that
depend on them. Preserve the flat architecture and current runtime dependencies.

Treat every finding in §2.3 as reproducible: most were confirmed end to end against real `.docx`
packages, and the reproductions are described precisely enough to rebuild. Add actual package
integration tests; do not leave a finding covered only in memory.

Do not claim a verifier defect is an observed cleaner deletion — F03 through F06 are injected-damage
tests, and the cleaner behaves correctly in all four. Do not treat faster runtime as permission to
weaken verification. Do not lower `MIN_PAIR_SIMILARITY` to fix F07. Do not fix F06 and F07 in a way
that entangles them without reading §5.2 first.

Where this plan states an estimate of size or effort, treat it as a planning comparison, not a
commitment: no part of this work has been implemented, and no line count or calendar figure here is
evidence of anything.

Complete W10 and report manual Word, corpus, and Windows-GUI limitations explicitly. Publishing a
release is outside scope.

### Independent reviewer

Examine the boundaries after W00, W03, W06, and at W10. **Own the injected-damage fixtures in §10.5**
and write them before or alongside the implementation rather than after it: W03's specific failure mode
is a verifier that agrees with its own implementation, and fixtures authored by the implementer are the
one thing that cannot catch it.

At each boundary, confirm the handoff contains problem, final behavior, files touched, exact tests run
and their result, intentional behavior changes, remaining risks, and required manual validation. A
passing test count is not a handoff.

At W10, evaluate the §10.6 acceptance criteria on sampled decisions from the corpus diff, and reject
any attempt to substitute a blanket Verified percentage for them.

### CI task (may run in parallel)

Own §15.3. Add a small test workflow on pull requests and normal branch pushes, read-only permissions,
no publishing side effects, testing on Windows at minimum — which is the only way the four skipped GUI
tests ever run. Either add `tests/**` to the release workflow's pull-request filters or document why
test-only changes use the dedicated workflow instead. Touch no application module.

## 20. Primary references for standards-dependent checks

These sources support specific implementation questions; they do not replace testing the actual
supported Word version and package representation.

- [GitHub Actions workflow syntax and path filters](https://docs.github.com/en/actions/reference/workflows-and-actions/workflow-syntax#patterns-to-match-file-paths):
  root-level `*` versus recursive matching — the mechanism behind F17.
- [Microsoft Open XML simple fields](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.simplefield?view=openxml-3.0.1):
  `w:fldSimple` and its field instruction — the carrier missing from `EMBEDDED_CONTENT_TAGS`.
- [Microsoft: working with WordprocessingML tables](https://learn.microsoft.com/en-us/office/open-xml/word/working-with-wordprocessingml-tables):
  table, row, and cell structure; verify edge-case acceptance in Word before claiming compatibility.

When implementing bookmark field parsing, numbering inheritance, or break semantics, consult the
applicable Microsoft/OOXML primary documentation and record the supported subset. Avoid inferring a
general schema rule from one synthetic fixture — and remember that the zero-row table in F11 passes
this project's own structural lint, which is exactly why lint agreement is not evidence.
