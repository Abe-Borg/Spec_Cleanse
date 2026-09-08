# SpecCleanse: detailed implementation and agent handoff plan

**Prepared:** 2026-09-08  
**Repository:** `C:\Github-Repos\Spec_Cleanse`  
**Reviewed baseline:** `dfb9c8a44d48b79877a097d14ba6d17ec7dcc27d`  
**Status:** proposed implementation work; no application changes have been made as part of preparing this plan.  
**Audience:** capable coding agents and the maintainer who will review their work.

## 1. Objective and scope

Make SpecCleanse safer at preserving specification requirements, make its verification verdicts more reliable, prevent batch output loss, and address the measured large-document bottleneck without replacing the existing architecture.

This plan is the deliverable requested after an analysis-only review. Creating this Markdown file is authorized. It is not a record that the implementation, full test suite, Word validation, or release has already happened. A receiving coding agent should implement the work when assigned this plan, subject to the permissions and instructions in that receiving session. Publishing a release is outside this implementation plan.

The highest-value outcome is fewer silent false positives. Retaining some uncertain editorial text is preferable to deleting a requirement. Verification must not interpret a substring match or one editorial run as permission to discard unrelated text.

Retain the flat Python layout, direct `lxml` manipulation, stdlib `unittest`, Tkinter GUI, and the two runtime dependencies. Do not bring back the archived structural or style optimization stages. Do not introduce a service, database, LLM call, dependency-heavy rule framework, or broad package reorganization.

The plan deliberately separates:

1. Confirmed cleaner behavior that needs correction.
2. Confirmed verifier weaknesses exposed by deliberately damaged outputs.
3. Structural concerns requiring standards and Word validation.
4. Measured performance work.
5. Optional product ideas that are deferred.

## 2. Evidence, confidence, and corrections to the original brief

### 2.1 What the review actually established

The original review brief was supplied at `C:\Users\AbrahamBorg\Downloads\spec-cleanse_REVIEW_BRIEF.md`. It contains useful leads, but its claims and instructions are not a substitute for independent investigation. The implementation should use the evidence below and recheck the current checkout.

The follow-up review read the active cleaner, detection, XML, verification, and GUI paths and performed synthetic checks on Windows using Python 3.14.6. Python bytecode caching was disabled. The broader probe explicitly blocked filesystem mutations. XML trees, source metadata, and output comparisons stayed in memory; file-operation boundaries were replaced for these probes.

Consequently:

- The tests established behavior of the relevant application functions, including the actual XML redactor and classification logic.
- They did not exercise ZIP unpacking/repacking end to end.
- They did not rerun the normal test suite, which creates temporary files.
- They did not measure precision or recall on a real specification corpus.
- They did not open or round-trip a generated document in Word.
- Existing line numbers below are navigation hints against the reviewed baseline, not stable identifiers.

### 2.2 Confirmed findings

| ID | Finding | Evidence and consequence |
|---|---|---|
| F01 | Copyright patterns classify plausible requirements as removable paragraphs | All three examples in section 6.1 below triggered removal; verification classified their disappearance as copyright. |
| F02 | Some high-confidence editorial patterns also overreach | `Select one of the listed manufacturers.` and a requirement containing `[retain or delete]` both triggered whole-paragraph removal. |
| F03 | Verification drops the conditions attached to low-confidence patterns | A plain paragraph containing `revise as required` was retained by the cleaner, but a deliberately missing copy was accepted by verification. |
| F04 | Verification promotes partial formatting evidence into whole-paragraph permission | Removing an entire paragraph containing ordinary requirement text plus one hidden run produced PASS in the injected-damage test. |
| F05 | Verification does not carry preserve-style evidence | Deliberately removing an `ART`-styled heading whose text matched an editorial rule produced PASS. |
| F06 | Verification allows an inline match to excuse extra deleted words | Removing `[Verify quantity]` together with `spare` from a requirement was classified as expected inline removal. |
| F07 | A correct long inline redaction produces a false verification failure | `Provide [Verify quantity with the Owner and the AHJ prior to bid] units.` became `Provide units.` and was reported as one unexpected removal plus one addition. |
| F08 | Simple fields are missing from the embedded-content safeguard | An editorial paragraph containing `w:fldSimple` lost the field; text verification and current structure inspection passed. |
| F09 | Non-`w:t` characters within redactions can remain | A placeholder containing `w:noBreakHyphen` became `Provide -units.` because the redactor counted but did not remove the separator element. |
| F10 | Complete bookmark targets can be lost | A bookmark fully inside a removed paragraph disappeared while an external `REF` field still referred to its name; verification passed. |
| F11 | Accepting the last deleted table row leaves a zero-row table | The table remained with no `w:tr`; current structural inspection reported no issue. Whether Word repairs or rejects the package was not tested. |
| F12 | Batch destinations can collide | `_output_for()` maps same-named inputs from different folders to one output when a common destination is selected; preflight only checks files already present on disk. |
| F13 | Verification failure is counted as GUI success | `_clean_one()` returns `True` after logging a failed verification, and `_run_clean()` counts it as succeeded. |
| F14 | Verification's main paragraph alignment dominates the reproduced large-input cost | At 6,000 paragraphs, survivor pairing consumed about 1.3% of profiled verification time. |
| F15 | Paragraph deletion repeatedly scans the entire parent container | `can_delete_paragraph()` builds complete block-child lists for each candidate removal. |
| F16 | Verifier fallback configuration bypasses the existing path resolver | `verify_clean()` resolves `patterns.yaml` beside `__file__` when no engine/path is supplied. |
| F17 | Test-only pull requests can miss CI | The release workflow's `*.py` path filter matches root files, not `tests/...`. |

F03–F06 are explicitly verifier tests with injected damage. Do not describe them as observed extra deletions by the current cleaner.

### 2.3 Performance evidence

The following measurements used the original brief's synthetic paragraph distribution, excluding ZIP and disk operations. They are not end-to-end timings or performance promises for another machine.

| Body paragraphs | Cleaning logic | Verification logic |
|---:|---:|---:|
| 500 | 0.0282 s | 0.0073 s |
| 2,000 | 0.1536 s | 0.1445 s |
| 6,000 | 0.7957 s | 2.9939 s |

A separate profiled 6,000-paragraph verification took 8.8903 s. `_pair_with_survivor()` took 0.1114 s cumulatively over 1,000 calls. `_removed_fragments()` took 0.0541 s cumulatively over 500 calls. The main paragraph-level `SequenceMatcher` dominated.

The brief's generator also repeats the same part heading: when `i % 12 == 1`, `i % 3 + 1` is always 2. Preserve that distribution as a regression case, but add other distributions rather than treating it as representative of all documents.

### 2.4 Corrections that must inform implementation

- Removing global `DOTALL` does not fix the supplied copyright false positives. They match on one line without it.
- Shorter wildcard gaps reduce some false positives but do not prove a paragraph is boilerplate.
- For a pure deletion, similarity `2 * len(after) / (len(before) + len(after)) < 0.5` means more than two-thirds of the original characters were lost, not roughly half.
- Eliminating a duplicate character comparison cannot halve total verification time in the reproduced workload. It is a secondary optimization.
- Immediate sibling checks are not automatically equivalent to the current block-container rules: markers and other non-block children can intervene.
- A CLI is not a prerequisite for a corpus harness; the harness can import the existing modules.
- Flattening the verifier's paragraph list is not a complete specification text export. Generated numbering, table relationships, notes, fields, and part identity need an extraction design.
- Sorting ZIP member names alone does not make archives byte-deterministic; timestamps and metadata also matter.
- Automatic numbering consequences depend on list membership, restarts, and inheritance. Do not claim every later number necessarily changes.
- Valid-looking XML and a passing custom lint are not proof that Word will open a document without repair.

## 3. Decisions selected for this implementation

### 3.1 Required work

Implement these in reviewable stages:

1. Add reproductions and independent expected outcomes for confirmed safety findings.
2. Narrow shipped whole-paragraph rules and make formatting-only deletion opt-in for new/default configurations.
3. Preserve the scope, prerequisites, and source location of removal evidence during verification.
4. Recognize exact permitted inline transformations regardless of retained-text percentage.
5. Fix separator redaction and simple-field preservation.
6. Protect referenced bookmark targets and safely handle tables emptied by accepted revisions.
7. Detect batch destination conflicts before writing any file and report verification failures as requiring review.
8. Use the shared configuration resolver and cover nested test changes in CI.
9. Optimize paragraph alignment after its correctness contract is established; optimize container scanning separately.
10. Update documentation, run representative corpus comparisons where files are available, and validate changed XML cases in Word.

### 3.2 Explicit product choices

- Keep producing a cleaned output when verification finds a problem, provided processing completed successfully. Mark it **Needs review**, distinguish it from a verified output, and identify its path. Do not silently delete it or rename it after the fact.
- A successful write and a passing verification are separate outcomes.
- Change `specifier_notes.formatting_only_removal` to `false` in shipped defaults and omitted-key defaults. Continue honoring an explicit `true` as an opt-in; show that choice in Preview/run reporting. Do not add a second GUI override for this switch in this work.
- Ambiguous prose such as a contractor being told to select a manufacturer is retained by default. Additional editorial text retained by narrower rules is an accepted tradeoff and must be documented.
- Preserve a referenced bookmark's target text rather than merely leaving an empty bookmark that makes a field return misleading content.
- Do not renumber documents or rewrite cross-references automatically. Initially report numbering-sensitive deletions as requiring review when reliably identified.
- Existing user-owned configuration files must not be silently overwritten or migrated. Safer defaults do not automatically repair an installed user's old patterns.

### 3.3 Deferred work

Do not include these unless later evidence and a separate task justify them:

- A public cleaning CLI, text/Markdown output sidecar, section chunking, deduplication, token estimation, caching, or multiprocessing.
- Converting to `python-docx`, adding pytest, reorganizing into a package, or restoring `legacy/` stages.
- General-purpose OOXML repair or complete schema validation of arbitrary input packages.
- Broad style-resolution redesign, speculative security hardening, or formatting changes unrelated to the confirmed defects.
- ZIP byte determinism, per-pattern telemetry, a new persistent settings system, or a general rules DSL.

A small developer benchmark/corpus runner is in scope. It is not a new public CLI product.

## 4. Non-negotiable invariants

1. **Preserve decisions dominate editorial cleaning.** Source text protected by a preserve pattern, preserve style, or explicit reference-target protection cannot be deleted or redacted through another editorial path. Explicit supported tracked deletions have the separate authority and consequence reporting defined in section 7.4.
2. **Evidence has a scope.** An inline match permits an interval removal. A qualifying run permits removal of its own text. Whole-paragraph removal needs whole-paragraph authority or complete coverage by eligible intervals.
3. **Prerequisites travel with rules.** Disabled categories, formatting-only settings, and low-confidence formatting requirements apply equally during cleaning and verification.
4. **Verification rereads the actual input and output.** The processor's target list, claimed removals, or `ProcessingResult` cannot be the sole authority for acceptance.
5. **Protected text remains in order.** No unexplained insertion, substitution, reordering, or loss of multiplicity becomes PASS because a fuzzy match found similar text elsewhere.
6. **Part identity is retained.** Text in a header cannot account for body text that disappeared. Notes and distinct content parts must not be flattened into one interchangeable pool.
7. **Offsets have one definition.** Use the same paragraph-owned text stream to interpret source intervals. Boundary trimming and any permitted whitespace cleanup must have an explicit mapping.
8. **Nested paragraphs are independent.** Do not cross into text-box paragraphs while handling an outer paragraph's runs or text.
9. **Structural carriers survive ordinary editorial removal.** Keep required blocks, section properties, fields, pictures, and note anchors. Accepted revisions may remove structures only under the explicit revision option and its documented rules.
10. **Mutations remain staged.** Collect targets, apply redactions, then run removals, then paragraph removals, or prove a revised order preserves all offsets and detached-element behavior.
11. **No silent batch overwrite.** A batch must have distinct output destinations, and no destination may overwrite a selected source document.
12. **Unknown is not PASS.** If alignment or structure cannot be verified, report the limitation as requiring review. Do not hide unresolved differences for speed.
13. **Do not improve runtime by removing safety checks.** Preserve-classification, multiplicity, part boundaries, and structure checks remain mandatory.
14. **Report what was tested.** Do not claim real-corpus precision, recall, or Word compatibility from synthetic tests alone.

## 5. Execution sequence and ownership

### 5.1 Work packages

| Package | Purpose | Dependency | Primary files |
|---|---|---|---|
| W00 | Establish baseline and regression ledger | None | Existing tests; developer validation notes |
| W01 | Tighten default policy and document configuration transition | W00 | `patterns.yaml`, `detection.py`, config/inline tests |
| W02 | Define source evidence and strengthen verification classification | W00; coordinate policy contract with W01 | `detection.py`, `verify.py`, metadata helpers in `docx_xml.py`, verifier tests |
| W03 | Fix long-redaction pairing and preserve part/order identity | W02 | `verify.py`, verifier tests |
| W04 | Fix separator redaction and simple-field handling | W00; W02 contract coordination | `processor.py`, `docx_xml.py`, inline/processor tests |
| W05 | Protect reference targets; handle revision-empty tables and numbering warnings | W02, W04 | `docx_xml.py`, `processor.py`, `verify.py`, structural/revision tests |
| W06 | Batch preflight and honest file/batch outcomes | W00; use agreed W02/W05 result types | `gui.py`, GUI tests |
| W07 | Configuration fallback and focused CI | W00; W01 for default behavior | `verify.py`, `apppaths.py` only if necessary, workflow files, config/path tests |
| W08 | Optimize measured alignment and container scans | W03, W04, W05 correctness settled | `verify.py`, `docx_xml.py`, benchmark utility, targeted tests |
| W09 | Integrated regression, corpus/Word validation, documentation | All implemented packages | Tests, `README.md`, `CLAUDE.md`, `CHANGELOG.md`, final report |

Safety corrections should be deliverable even if performance work needs another iteration. Do not make a faster diff algorithm a prerequisite for fixing false-positive deletion.

### 5.2 Agent coordination

An orchestrating agent should own the evidence/result contract and integration. Coding agents may work independently where files and interfaces do not conflict, but should use isolated branches/worktrees if the execution environment supports them.

Suggested allocation:

- **Policy and verification agent:** W01–W03, coordinating the metadata contract before other agents use it.
- **XML safety agent:** W04–W05 after the evidence/protection interfaces are agreed.
- **GUI and CI agent:** W06 and the workflow portion of W07. Coordinate verifier fallback changes with the verification owner.
- **Performance agent:** W08 only after the corrected verifier is available. This can be the verification agent in a later wave.
- **Integrator/reviewer:** W00, integration, and W09; owns final shared-document edits.

Avoid simultaneous uncoordinated edits to `verify.py`, `processor.py`, `docx_xml.py`, and common fixtures. Each agent must name its owned files, proposed interface changes, tests, and unresolved questions before substantial work. Do not ask each agent to rewrite all documentation independently; send accurate documentation notes to the integrator.

Each package's handoff must contain: problem, final behavior, files touched, exact tests run and their result, intentional behavior changes, remaining risks, and any required manual validation. A passing test count without this context is insufficient.

### 5.3 Relative effort, risks, and delivery boundaries

These are planning comparisons, not promises of agent-hours. Do not make every package one giant change merely because the table groups related work.

| Package | Relative effort | Main benefit | Principal implementation risk |
|---|---|---|---|
| W00 | Small | Establishes reproducible failures and a trustworthy baseline | Fixtures accidentally duplicate the implementation's assumption rather than independently specifying correct behavior |
| W01 | Medium | Reduces requirement deletion under default policy | Reduced editorial recall; installed users continuing to run old customized rules |
| W02 | Large, highest reasoning demand | Stops permissive verifier acceptance of unexplained loss | Incorrect interval coordinates, evidence applied at the wrong scope, or accidental coupling to processor actions |
| W03 | Medium to large | Removes false alarms without hiding real edits | Ambiguous duplicate text, lost multiplicity, or pairing across part/story boundaries |
| W04 | Medium | Corrects concrete XML redaction/carrier defects | Removing a semantic break or accidentally stripping a field/anchor while cleaning empty runs |
| W05 | Large; split references and revisions into separate changes | Preserves reference meaning and prevents malformed table residue | Incomplete field parsing, overprotecting large bookmark ranges, and nested-container edge cases |
| W06 | Medium | Prevents batch output loss and misleading success summaries | Path equivalence on Windows and incomplete propagation of new outcome types |
| W07 | Small to medium | Fixes config selection and test coverage | Unwanted seeding during engine-supplied verification or surprising validation of existing settings |
| W08 | Medium to large, evidence dependent | Makes large manuals practical | Faster but incorrect alignment or stale container metadata after mutations |
| W09 | Medium plus corpus/Word availability | Establishes what is actually ready to trust | Treating unavailable manual evidence as a pass or missing an installed-config transition |

Suggested coherent change sequence: baseline fixtures; default policy; verification evidence; pairing; separator/field handling; bookmark protection; revision-table handling and numbering notices; batch/outcome reporting; resolver/CI; measured performance; final integrated documentation. A fixture may land with its fix so mainline is not deliberately left red. Verify each integrated stage before changing the next shared contract. If an agent proposes a larger redesign, require it to explain why the smaller correction cannot satisfy the same acceptance tests; prefer the smaller correction when both work.

## 6. W00–W01: baseline, fixtures, and safer detection policy

### 6.1 Baseline procedure and exact reproductions

Inspect the current status and applicable repository instructions. Preserve unrelated user changes. Record the current commit and differences from the reviewed baseline. Do not assume the historical test count still applies.

Run the existing suite before changing behavior. Use the project environment when valid, or an available environment with the declared dependencies; do not reinstall dependencies unnecessarily.

Windows examples, from the repository root:

```powershell
& .\venv\Scripts\python.exe -B -m unittest discover -s tests -t .
& .\venv\Scripts\python.exe -B -m unittest tests.test_inline tests.test_verify
```

The `-B` flag suppresses bytecode output; the normal tests still intentionally write temporary fixtures. Respect any write restrictions in the receiving session.

Create regressions using `tests/docx_builder.py` and `DocxTestCase`. Add raw XML assertions where text-only assertions would miss damage. Do not replace these tests with mocks of the behavior being tested.

**Copyright negatives: these must retain their requirement text under the new shipped defaults.**

```text
Shop Drawings submitted under this Section may not be reproduced for use on other projects.
Contractor shall verify that duplication of sprinkler coverage in adjacent zones is prohibited by the AHJ.
Unauthorized personnel shall not have access to the fire pump room; reproduction of access keys is not permitted.
```

**Additional policy negatives:**

```text
Select one of the listed manufacturers.
Provide two [retain or delete] spare filters per unit.
Provide pumps and revise as required.
Retain records of all tests required in Paragraph 1.6.
Select one-piece molded fittings for changes in direction.
```

For the mixed `[retain or delete]` example, the initial selected behavior is to retain the complete paragraph unchanged. Do not invent a broad new inline rule merely to remove those brackets. Any later narrowly justified inline rule must keep every surrounding requirement word and have its own positive/negative fixtures.

**Representative positive fixtures that should continue cleaning:**

```text
© 2026 ARCOM. All rights reserved.
[Specifier: delete this note before issue]
Retain subparagraph below for wet-pipe systems.
Copy paragraphs above for each additional riser.
Provide two [Verify quantity with Owner] spare sprinklers.
```

The expected output of the last example is exactly `Provide two spare sprinklers.` These strings are synthetic policy fixtures, not claims about measured corpus frequency.

### 6.2 Pattern changes

Review every shipped whole-paragraph pattern, not just the three copyright examples. Focus on unanchored prose fragments, generic selection instructions, delimited note matches embedded in requirements, and broad SpecAgent matching.

Required approach:

1. Classify each changed default rule by the scope it intends to remove and why that scope is justified.
2. Remove or narrow ambiguous standalone copyright phrases such as `may not be reproduced`, `duplication.*?prohibited`, and `unauthorized.*?reproduction` as independent whole-paragraph triggers.
3. Prefer patterns that recognize a complete standalone notice or a clearly editorial paragraph. A copyright year or publisher name somewhere inside a requirement is insufficient on its own.
4. Require explicit editorial context for ambiguous selection prose. Do not retain the broad `^...select one...` rule simply because one existing test expects it.
5. Keep mixed requirement/note paragraphs unless an existing, narrowly authorized inline transformation accounts for the removable interval.
6. Do not globally change regex flags as a substitute for a rule audit. Preserve multiline notice support where intended and add multiline negative tests.
7. Keep named categories and the existing configuration surface understandable. Do not replace the detector hierarchy with a new framework.

For every edited rule, add at least one intended removal and one plausible requirement that must remain. Test matches split across runs when matching depends on paragraph text. Test case variants and ordinary whitespace without normalizing away meaningful content.

Use known publisher strings only as evidence for a specific notice grammar, never as a universal permission to remove any paragraph mentioning that publisher. Explicitly preserved headings still win.

### 6.3 Formatting-only default

Set `formatting_only_removal` to `false` consistently in:

- The shipped YAML.
- `PatternConfig` and any omitted-key construction defaults.
- Verification's omitted-key behavior.
- Tests and documentation describing the default.

Explicit `true` must retain the opt-in behavior and its `formatting-only` label. Turning it off must not disable genuinely qualifying editorial styles, hidden text, or supported low-confidence-plus-formatting rules. Those remain separate decisions.

Update the current test that explicitly asserts removal is on by default. Preserve a separate opt-in test so changing the default does not accidentally remove the feature.

### 6.4 Existing installed configurations

`apppaths.resolve_config_path()` prefers an existing executable-adjacent or per-user configuration. Changing the bundled YAML alone will not update those files.

Required transition behavior:

- Never overwrite an edited or previously seeded user configuration automatically.
- Continue logging the exact active configuration path.
- Report when formatting-only removal is enabled, including when it comes from an older explicit `true`.
- Add a focused, actionable notice for known superseded shipped patterns when they remain active. Compare actual rule strings/known prior defaults, not file modification times or a speculative version guess.
- Distinguish “new defaults are safer” from “the active configuration has been updated.” Do not claim the latter without evidence.
- Document how an installed user can obtain the current shipped default from the distribution/source and compare it with the active file. The revised default must be practically accessible; a temporary frozen extraction path alone is not adequate documentation.
- If exact legacy-default detection is used, preserve customizations and test the known baseline, a customized baseline, and a current configuration.

Do not silently redefine all custom regex semantics to accomplish a default-rule correction. If a proposed detector change intentionally changes custom-rule behavior, document that change, provide before/after examples, and have the integrator review its compatibility consequences.

**Exit gate:** the new default negative fixtures survive; established unambiguous positives still clean; opt-in formatting remains available; existing user configuration is untouched; release notes explain the intentional reduction in aggressive matching.

## 7. W02: a source-based verification contract

### 7.1 Problem to solve

Current `removal_patterns()` flattens policy into `(category, regex)` pairs. The verifier consequently cannot distinguish text-only authority from a rule requiring formatting, nor a permitted interval from a whole paragraph. Current `ParagraphInfo` aggregates some signals with `any`, which loses the distinction between one editorial run and a fully editorial paragraph.

Fix these distinctions directly. Do not patch each supplied phrase with a special-case exception.

### 7.2 Recommended minimum data model

Use small dataclasses or a comparably explicit typed representation. Suggested information, not mandatory class names:

| Concept | Required information |
|---|---|
| Paragraph source | Package-relative part name, relevant story/note identity, document-order position, raw owned text, display text if trimmed |
| Run source | Start/end offsets into raw paragraph text, text, relevant effective formatting/style evidence, source ownership |
| Preservation evidence | Whole-paragraph protection, reason, preserve pattern/style/reference target that established it |
| Removal evidence | Category, rule identifier/reason, source interval or whole-paragraph scope, and eligibility prerequisites already evaluated against the source |
| Revision evidence | Whether a supported tracked deletion authorizes loss under the selected option |
| Comparison outcome | Actual before/after text, accepted and unexplained intervals, preservation violations, location, and explanation |

Offsets must not be computed against stripped text and then applied to raw text. Either preserve raw text throughout or record a reliable boundary mapping. A normalized display string is not an offset authority.

Document-model extraction stays in `docx_xml.py`. Decisions about what qualifies as editorial stay in `detection.py`. Comparison and classification stay in `verify.py`. A single small shared evidence type is acceptable; a broad new abstraction layer is not required.

### 7.3 Independence boundary

It is acceptable to share immutable rule definitions and source-evidence evaluation. It is not acceptable to trust the processor's claimed actions as proof that the resulting document is correct.

Verification must reread the original package and the actual output package and independently establish:

1. Which source characters and structures existed.
2. Which source intervals the configured policy permits removing.
3. Whether the actual output retains all protected content in order.
4. Whether structural carriers and reference targets survived as required.

Do not call the XML-mutating cleaner to manufacture verification's expected document. A source-based pure transformation using evaluated removal intervals is appropriate; reusing processor-produced spans without revalidation is not.

### 7.4 Classification rules

**Preservation:** compute source preserve-style and preserve-text decisions before removal evidence. A protected paragraph cannot be excused by matching a removed fragment or an editorial pattern. Include inherited and display-name-based styles using the existing resolver. Protect modifications as well as complete paragraph loss.

**Explicit tracked deletions:** when the revision option is enabled, a supported source revision that explicitly deletes a row/cell/run is separate authority from editorial matching. Its intended text loss can be classified as an accepted tracked deletion even if that text would otherwise be protected from editorial cleaning. Validate the actual source revision and its exact extent; do not use a revision somewhere nearby as blanket permission. Broken reference consequences or unsupported structural outcomes still require review. When the option is off, that authority does not exist. Add a preserved-heading-in-deleted-row fixture to make this precedence explicit and avoid an accidental new false alarm.

**Whole-paragraph removal:** accept only when a whole-paragraph rule actually qualifies at that scope, a supported accepted revision authorizes it, or the union of eligible removal intervals covers all substantive source text. A hidden run, colored run, or low-confidence match somewhere in the paragraph is not sufficient.

**Low-confidence rules:** evaluate formatting on the same scope as the rule. An unrelated colored run cannot boost a plain requirement elsewhere in the paragraph. Honor category enablement and applicable style settings. Reproduce the supported cleaner behavior deliberately; do not infer eligibility from a regex match alone.

**Run removal:** represent eligible run text with source intervals. Do not use substring membership in a list of editorial run strings as location evidence: repeated text can occur in both editorial and protected runs.

**Inline removal:** every deleted substantive character must lie within an authorized source interval. Whitespace handling must be explicitly defined and bounded, including the existing `tidy_spans()` adjacent-space rule. `[Verify quantity]` does not authorize deleting `spare` immediately after it.

**Mixed categories:** preserve violation outranks unexpected loss; unexpected loss outranks expected changes. Report the relevant categories/reasons without labeling all fragments with the first successful category. Avoid an unnecessary result-schema overhaul solely for presentation.

**Disabled configuration:** disabling a detector must also prevent its signals from justifying output loss. Test hidden-text disabled, specifier-note disabled, style detection disabled, and formatting-only disabled independently.

**Unchanged content:** verification is a loss-safety check, not a recall requirement. Retaining eligible editorial text does not itself fail verification. Report recall separately when corpus evaluation is available.

### 7.5 Mandatory injected-damage tests

Construct original and damaged `.docx` files independently; do not run the cleaner to produce these damaged outputs.

| Test | Source and damage | Required result |
|---|---|---|
| V01 | Ordinary requirement plus a hidden note run; remove the complete paragraph | Unexpected removal; not PASS |
| V02 | Plain `Provide pumps and revise as required.`; remove it | Unexpected removal; not PASS |
| V03 | `ART`-styled `Select one of the listed manufacturers.`; remove it | Preserve violation |
| V04 | `Provide two [Verify quantity] spare filters per unit.` becomes `Provide two filters per unit.` | Unexpected loss of `spare`; not PASS |
| V05 | Same phrase occurs in a visible run and hidden run; delete the visible occurrence | Unexpected loss unless the actual surviving sequence is provably equivalent with all protected text retained |
| V06 | Hidden detector disabled; delete hidden-marked text | No hidden-text permission to excuse it |
| V07 | Protected heading loses a fragment that does not itself match a preserve regex | Preserve violation based on the original paragraph |
| V08 | Several permitted fragments plus one unpermitted fragment disappear | Whole modification verdict fails |
| V09 | Paragraph with all substantive text in eligible runs disappears | Expected, provided no preserve or structural rule forbids the loss |
| V10 | Input has duplicate requirements; output has one fewer | Missing multiplicity must be detected |
| V11 | Body requirement disappears but identical header text remains | Header text must not excuse body loss |
| V12 | Insert, substitution, or reordered protected clauses | Not accepted as an expected deletion |

For ambiguous identical text, do not claim source occurrence identity that the output format cannot establish. The safety assertion is preservation of required text, multiplicity, order, and available structural identity. Ambiguity must not be resolved by assuming the editorial occurrence was the one deleted.

**Exit gate:** F03–F06 fail for the right reason under deliberately damaged outputs; ordinary supported cleans still pass; the contract is documented clearly enough for the XML and performance agents to use.

## 8. W03: exact transformations, pairing, and part identity

### 8.1 Pair permitted transformations before fuzzy heuristics

For a source paragraph with eligible inline/run intervals, compute the expected text for the intended supported removal combination without mutating XML. An exact output match to that text is a stronger pairing signal than character similarity and must work even if almost all source characters were removable.

Use the actual original source evidence, not a string supplied by `ProcessingResult`. Keep an unchanged paragraph as an acceptable alternative. If multiple supported outcomes are permitted, represent them explicitly or validate deletion against the eligible intervals; do not enumerate arbitrary subsets exponentially.

Character-level diff output can help present losses after a pair is established. It cannot be the sole authority for whether an edit was a permitted deletion. Repeated characters can make an unconstrained diff choose misleading fragments.

Do not fix F07 by merely reducing `MIN_PAIR_SIMILARITY`, accepting all short survivors, dropping `added` from the pass predicate, or suppressing unexpected changes.

### 8.2 Identity and ordering

Compare within matching package parts. Preserve relevant story identities, particularly footnote/endnote IDs, so a missing note cannot borrow an identical paragraph from another note. Use full package-relative part names; `word/document.xml` and `word/glossary/document.xml` must not collapse to the same identity.

Preserve monotonic source/output ordering and multiplicity. Exact duplicate text does not establish a unique source position. Do not use mutable paragraph indices or generated XPath positions as if they were stable IDs across removals.

Begin with exact unchanged anchors and exact permitted-transformation candidates. Keep any fuzzy fallback conservative and constrained by the neighboring established alignment. If the verifier cannot resolve an ambiguous region safely, return an actionable verification limitation rather than certifying it.

### 8.3 Required cases

- The exact long-placeholder example becomes `Provide units.` with one expected modification and no addition/unexpected removal.
- A paragraph reduced to a short but substantive survivor is recognized; a placeholder-only paragraph still follows whole-paragraph rules.
- Multiple placeholders, split runs, leading/trailing whitespace, and repeated words retain correct source intervals.
- Adjacent similar paragraphs cannot exchange survivors to conceal an unexplained removal.
- Duplicate unchanged headings and requirements preserve multiplicity.
- Multiple changed paragraphs between two unchanged anchors are handled deterministically.
- Empty documents and parts, insert-only output, missing parts, and unchanged documents receive explicit outcomes.
- The same text in different parts or notes does not mask a loss.

**Exit gate:** supported long redactions pass without weakening injected-damage tests or changing the verdict for real additions/substitutions.

## 9. W04: inline XML redaction and field carriers

### 9.1 Remove covered separator elements

`iter_text_nodes()` yields `w:t` and supported separator elements. `_redact_spans()` currently advances offsets for all of them but only edits `w:t`.

Implement element-aware redaction:

1. Snapshot the paragraph-owned text-node stream before mutations.
2. Map each authorized span onto that stream using the existing coordinate system.
3. Partially edit `w:t` with `set_text()` as appropriate.
4. Remove a supported one-character separator element when its character is covered by a redaction.
5. Keep separators outside the interval exactly as they were.
6. Remove genuinely empty runs only after checking required structural content, range anchors, and already-detached state.

Cover `w:tab`, `w:br`, `w:cr`, `w:ptab`, and `w:noBreakHyphen` as represented by the extractor. Do not expand text extraction to unrelated element types just to make the test pass. For a break that carries page/column semantics, preserve it and report uncertainty if removing it would violate the content-only contract; explicitly establish the intended rule and test it rather than treating all break types as interchangeable whitespace.

Add cases where a placeholder crosses a hyperlink or content-control wrapper and where the same run is scheduled for both redaction and removal. Nested text boxes must remain independently processed.

### 9.2 Preserve simple fields

Recognize `w:fldSimple` as a structural field carrier. Keep its instruction attribute and required wrapper during ordinary editorial text stripping. Include nested simple fields and fields under wrappers in focused tests.

Do not remove field instructions or update field values programmatically. Blank only text the supported policy authorizes removing. Where field results are preserved or emptied, document that Word may recalculate them; keeping a field wrapper is not a promise that a removed editorial field result will stay absent after refresh.

Strengthen verification enough to catch loss or alteration of a supported field carrier even when its cached text is editorial and matches a removal pattern. Compare field evidence by part and instruction/structure, not just a document-wide field count. Allow supported changes only where accepted revisions legitimately remove that field and the source revision evidence explains it.

### 9.3 Required fixtures

| Test | Required behavior |
|---|---|
| X01 | Placeholder containing `w:noBreakHyphen` | Output is `Provide units.`, without an orphan `-` |
| X02 | Placeholder containing supported tabs and line breaks | Covered separators disappear; separators outside the span remain |
| X03 | Placeholder split across multiple `w:t` nodes and wrapper boundaries | Exact surviving requirement text and balanced XML |
| X04 | Redaction empties a run also scheduled for removal | No double-removal error or loss of neighbors |
| X05 | Editorial paragraph contains `w:fldSimple w:instr="REF Target"` | Field survives ordinary removal processing |
| X06 | Deliberately damaged output removes that simple field | Verification flags structural/reference loss |
| X07 | Field instruction/cached content in a nested text box | Outer processing does not double-count or strip nested requirement text |

**Exit gate:** exact text assertions and XML-carrier assertions pass; no field/section/picture preservation regression; Word validation cases are queued for W09.

## 10. W05: referenced bookmarks, revision-empty tables, and numbering notices

### 10.1 Referenced bookmark targets

Relocating half-open markers is not sufficient when the complete target disappears. Keeping only an empty bookmark avoids one error message but can still change the meaning of a `REF` field. The selected behavior is conservative retention of the referenced target's content.

Add a bounded source inventory of bookmark names/ranges and supported internal consumers. At minimum cover simple-field and complex-field `REF` references and internal hyperlink anchors. Evaluate `PAGEREF` and `NOTEREF` using the documented field grammar before claiming support.

For complex fields, account for instruction text split across runs and quoted bookmark names. Do not scan ordinary visible prose for the word `REF` and treat that as a field. Reuse a small field-instruction reader rather than implementing a general Word field evaluator.

Before target selection, mark paragraphs/ranges whose removal or redaction would alter an internally referenced target as protected. The same decision must be reflected in Preview and in verification. Prefer keeping an affected paragraph in full in this version instead of partially rewriting target ranges.

For malformed or unsupported reference forms, report the limitation and preserve the implicated target when it can be identified. Do not invent a target, retarget a field, create a replacement bookmark, or rewrite field results.

References spanning multiple paragraphs must protect the affected target content across the range, not just the start/end paragraphs. Maintain part/story identity and distinguish a newly missing target from one already missing in the input. Avoid a document-wide unconditional “all bookmarks protect all text” rule; that could disable cleaning on common automatically bookmarked documents.

Tests must cover complete and half-open ranges, multi-paragraph targets, simple/complex references, quoted names, internal hyperlinks, unreferenced bookmarks, existing broken references, and bookmarks inside explicitly accepted tracked deletions. If accepting a revision deletes a referenced target, report that consequence as requiring review even though the text deletion itself was requested.

### 10.2 Tables emptied by accepted revisions

Trace `_accept_structural_deletions()` when deleting a last row or last cell. Do not create invalid residue merely to keep a `w:tbl` element present.

Preferred behavior: when accepted revisions remove all rows of a table, remove the now-empty table and restore only the minimum required block content in its parent where necessary. Preserve surrounding paragraphs, required terminal cell paragraphs, section properties, and unaffected tables. Do not insert fake empty rows as a blanket repair.

Handle nested tables from the inside outward so removing an inner structure does not invalidate the outer traversal. Inspect container context before removing a table that is the only block in a header/footer, note, text box, cell, or body.

A table removed by an accepted revision is an explained structural change, not an unexplained loss. A newly introduced zero-row table or zero-cell row must be observable by the verifier. Validate the schema/Word assumptions for the supported representation; report Word behavior separately from custom lint results.

Test last-row deletion, deletion of every cell in a row, multiple deleted rows with one survivor, nested tables, sole-table containers, preserved headers/footers, and revision stripping disabled. Keep the current policy of not merging deleted paragraph marks unless a separate change is justified.

### 10.3 Numbering-sensitive removals

Do not attempt automatic renumbering or cross-reference rewriting in this work.

Add a focused review notice when a removed paragraph participates in automatic numbering, using direct `w:numPr` and supported inherited paragraph-style evidence. Read the relevant source metadata once rather than parsing numbering separately for each removal. An inherited value explicitly disabling numbering must not be treated as participation; validate the OOXML semantics before implementation.

The notice should say that displayed numbering or paragraph references may change; it must not assert that a reference definitely broke. Include part/location and a short source preview. This contributes to a **Needs review** file outcome without falsely claiming a text-integrity failure.

If inherited numbering support cannot be implemented reliably within the small metadata extension, ship accurate direct-numbering detection with an explicit limitation and a tracked follow-up. Do not claim complete numbering awareness.

**Exit gate:** referenced target content is preserved or an explicit accepted-revision consequence is reported; no new empty-table residue is produced; numbering warnings are precise about what was detected and what remains uncertain.

## 11. W06: batch preflight and truthful GUI outcomes

### 11.1 Build the batch manifest before processing

Build a list of input/output pairs from the snapshotted selection and destination on the main thread. Validate the complete manifest before starting the worker or writing any output.

Required checks:

- Two selected inputs must not map to the same destination.
- A planned output must not equal any selected input, including an input that will be processed later.
- Apply appropriate Windows case-insensitive normalized path comparison. Use resolved/normalized paths where possible; for existing paths, consider filesystem identity to catch equivalent spellings. Do not assume string equality proves different files.
- Retain existing overwrite confirmation for previously existing outputs only after internal collisions have been ruled out.

Selected behavior on internal collision: reject the batch before any write and show which sources conflict and the destination. Tell the user to choose separate destinations or change the source/output organization. Do not silently append arbitrary suffixes or let the later file win.

Pass the validated manifest to the worker. Do not recompute output paths later from mutable GUI state. This is a preflight protection, not a claim of locking the directory against all external processes.

### 11.2 Model file outcomes explicitly

Use a small enum/dataclass or similarly explicit result instead of one Boolean that conflates everything. Suggested outcomes:

| Outcome | Meaning | Count in summary |
|---|---|---|
| Verified | Output written, verification completed and passed, no actionable review notice | Verified |
| Needs review | Output written, but verification failed or an actionable structural/numbering/configuration concern remains | Needs review |
| Failed | Processing failed, or verification could not complete | Failed; state whether an output exists |

If verification raises after a successful write, report that the output exists but is unverified. Do not hide the path or claim that nothing was produced. Continue processing other independent files as the current GUI does.

Example summary:

```text
Done: 8 verified, 2 need review, 1 failed.
```

Replace wording such as `Verifying no spec content was lost...` with a statement of what the checks actually establish. A suitable passing message is `PASS — no unexplained text changes or new checked structural problems were found.` It must not imply complete engineering or Word-format certification.

Keep logs queued, widget updates on the main thread, controls disabled during the run, and inputs snapshotted before worker execution. Every exception path must restore controls through the established mechanism.

### 11.3 Tests without opening a window

- Same source basename from two folders with one output directory: manifest rejected; processor never called.
- Distinct names: both pairs accepted.
- Case-only destination collision on Windows: rejected.
- Output equals a selected source or its equivalent existing path: rejected.
- Previously existing output: original confirmation behavior still applies.
- Passing verification: verified count increases.
- Failed verification with an existing output: needs-review count increases, path logged.
- Verification exception after output creation: failed/unverified outcome, path logged.
- Processing exception: batch continues, final counts accurate.
- Source/destination controls changing after start cannot alter the validated manifest.

Update current tests that assert `_clean_one()` returns a Boolean. Do not keep a misleading truthiness compatibility behavior unless a real caller requires it.

**Exit gate:** no planned intra-batch overwrite is possible; a failed verification cannot contribute to the verified count; output disposition is explicit.

## 12. W07: configuration fallback, validation, and CI

### 12.1 Use the shared resolver

In `verify_clean()`, use `apppaths.resolve_config_path()` only when both an engine and explicit config path are absent. Supplying an engine should not trigger irrelevant fallback resolution or first-run configuration seeding. An explicit config path takes precedence over automatic resolution when no engine is supplied.

Preserve and document the existing engine-versus-path precedence when both are supplied; avoid a silent behavioral change. Test it explicitly.

Test source mode, executable-adjacent configuration, existing per-user configuration, first-run seeding, read-only fallback, explicit path, and engine-supplied paths. Reuse the frozen-build simulation already present in `tests/test_apppaths.py`; do not require an executable build for every path test.

### 12.2 Focused config checks

Validate fields touched by the change: Boolean options must be Booleans, relevant lists must contain strings, and editorial colors must have the supported hexadecimal shape. Decide whether to normalize a leading `#` or reject it with a useful message; document and test one consistent behavior. Do not accept the string `"false"` as a truthy switch.

Do not reject style names simply because they are absent from one input document; styles differ across templates. Avoid broad unrelated validation changes.

### 12.3 CI coverage

Add a small test workflow on pull requests and normal branch pushes, with read-only repository permissions and no release publishing side effects. At minimum test on Windows, which exercises Tkinter and relevant path behavior. A Python 3.10 compatibility lane and the existing supported Windows Python version are useful; choose a small matrix rather than an expensive exhaustive one.

Keep the release build workflow for packaging validation. Add `tests/**` to its pull-request filters, or deliberately document why test-only changes use the dedicated test workflow without packaging. Do not leave test-only changes with no CI.

No new test framework is needed. The test workflow runs `python -B -m unittest discover -s tests -t .`. Keep large timing benchmarks out of ordinary CI; small deterministic correctness and work-count regressions belong there.

**Exit gate:** verifier uses the intended active config under every call mode; malformed touched options fail clearly; test-only changes receive test coverage without publishing anything.

## 13. W08: measured performance improvements

### 13.1 Establish a corrected baseline

Benchmark after W01–W05 so timing improvements are not confused with changed cleaning policy. Measure baseline and candidate on the same machine, Python version, fixtures, and active configuration. Profile separately from wall-clock timing.

Create a developer utility, for example `tools/benchmark_pipeline.py`, that uses synthetic fixtures and the existing APIs. Keep its outputs in an explicitly chosen temporary/developer directory. It should record paragraph counts, relevant pattern distribution, cleaning time, verification time, output text digest, outcome counts, and interpreter/version information.

Do not compare ZIP bytes as the behavior oracle. Compare extracted semantic content, relevant XML carrier/reference evidence, and expected verification results.

Fixture families:

1. The original repeated-heading distribution at 500, 2,000, 6,000, 12,000, and 24,000 paragraphs where practical.
2. Mostly unique requirements with sparse editorial removals.
3. Repeated identical headings and repeated identical requirement paragraphs.
4. Long sequences of changed paragraphs without unchanged anchors.
5. Many inline placeholders, including short survivors.
6. Very long individual paragraphs, kept separate from paragraph-count scaling.
7. Tables and multiple content parts.
8. Deliberately damaged output with unexplained losses; faster verification must still fail.

Use repeated unprofiled runs, reporting a median and input/setup boundaries. Do not extrapolate a universal complexity law from a few timings.

### 13.2 Main paragraph alignment

Optimize the paragraph-level alignment at the current `verify.py` outer matcher first. Candidate approach:

1. Partition by part/story identity.
2. Establish uniquely matching exact unchanged anchors in monotonic order, using a patience/LIS-style approach or another justified deterministic method.
3. Align the bounded regions between anchors, preferring exact permitted transformations from W03.
4. Handle duplicate occurrences explicitly, preserving multiplicity and order.
5. Apply conservative fallback only to unresolved regions and record any inability to verify fully.

This is a suggested algorithm, not permission to weaken the comparison. A simpler approach is preferable if it passes the same adversarial tests and measured workloads.

Do not turn `autojunk` on blindly, use a bag-of-paragraphs comparison, limit candidate windows without accounting for excluded possibilities, or skip large ambiguous regions while returning PASS. If a resource bound is needed, surface an incomplete-verification outcome with the unresolved location; that is a fallback limitation, not completion of the performance objective.

Remove the redundant character-level matcher only after correctness is settled. For an already established pure deletion, the similarity ratio can be derived from lengths. Length prefilters are useful where mathematically justified, but no longer apply as a universal rejection rule to exact permitted long redactions.

### 13.3 Container scan optimization

Replace repeated materialization of all block children only if the replacement remains easy to audit. Consider an early-exit scan, cached per-container block metadata updated during the removal phase, or targeted sibling traversal. Choose based on profiling and clarity.

The replacement must preserve behavior with:

- A sole paragraph.
- Several removable paragraphs processed consecutively.
- Tables and paragraphs interleaved.
- Non-block markers before, between, and after blocks.
- A table cell that must end in a paragraph.
- Section properties at the end of a body.
- Nested containers and wrappers.
- Earlier mutations that change sibling relationships.

Do not cache counts across mutations without updating them. Do not claim constant-time behavior if the implementation still walks an unbounded sibling chain.

### 13.4 Acceptance and stopping rule

- All corrected correctness tests and injected-damage tests pass.
- Report before/after medians at several sizes and explain where the gain comes from.
- For the repeated-heading 12,000/24,000-paragraph family, aim for a substantial multi-fold improvement and growth well below the prior approximately cubic behavior. Treat these as review targets, not machine-independent hard seconds in CI.
- A material improvement only to character pairing does not satisfy the main bottleneck objective.
- No meaningful small-document regression should be introduced; distinguish millisecond noise from an actual regression.
- Add a deterministic small regression for the algorithm's problematic pattern, using work counts or a bounded fixture where appropriate rather than fragile wall-clock assertions.
- Stop after the measured hotspots are resolved. Do not then rewrite all ZIP I/O, add concurrency, or optimize unrelated allocations without new evidence.

**Exit gate:** large-input verification is measurably better on the corrected behavior, and no safety property was traded away to obtain it.

## 14. W09: integrated validation and documentation

### 14.1 Automated validation layers

Use three distinct layers:

1. **Focused unit/algorithm tests:** interval eligibility, path manifests, source metadata, and alignment edge cases.
2. **Synthetic package integration:** build actual DOCX ZIPs, run the public processor/verifier path, reread output XML, and assert text and structural behavior.
3. **Injected-damage verification:** independently construct damaged outputs so agreement with the cleaner cannot make a broken verifier appear correct.

Do not mock the parser, writer, or structural inspection in the final package integration tests. The earlier review's in-memory probes are leads, not a substitute for this layer.

Extend `tests/docx_builder.py` narrowly. Check that any fixture used for Word validation has complete enough namespaces, relationships, and field/bookmark definitions to be a valid positive control. Existing minimalist fixtures are not automatically valid Word compatibility fixtures. For example, verify that every namespace named by `mc:Ignorable` is actually declared.

Run the full suite after integration, and once more only if subsequent edits justify it. Record interpreter, platform, test count, skips, and failures. GUI tests skipped on Linux do not establish Windows GUI behavior.

### 14.2 Corpus comparison

When representative real documents are available in the receiving workspace, compare baseline and candidate outputs without modifying sources. Use a private temporary output directory; do not commit proprietary specifications or excerpts without the maintainer's authorization.

A useful initial selection includes distinct MasterSpec, SpecLink, and firm-edited styles when actually available; colored requirement edits; tables; references; notes; and a consolidated manual. Do not invent coverage for a document family that was not supplied.

For each changed decision record:

- Source file and part/location.
- Original paragraph or a permitted private preview.
- Baseline action and candidate action.
- Rule/evidence explaining the difference.
- Whether the change is an intended conservative retention, a confirmed precision improvement, or unresolved.

Review all newly removed substantive content. Narrowing rules may intentionally increase retained editorial text; record this as a recall tradeoff rather than assuming it is a regression. A harness can call the existing importable API and does not require a public CLI.

If no real corpus is available, finish all synthetic and code work and label corpus validation outstanding. Do not fabricate precision/recall percentages or halt unrelated implementation solely because corpus data is absent.

### 14.3 Word validation

Open original and cleaned copies of representative affected documents in Windows Word. Validate:

- No new repair/unreadable-content prompt.
- Expected requirement text, headings, and section boundaries remain.
- Headers, footers, note anchors, pictures, and fields still behave as intended.
- Simple fields survive; refresh references in a disposable copy and inspect their result.
- Referenced bookmark targets resolve without new broken-reference errors or misleading empty target values.
- Last-row and last-cell revision cases produce an acceptable layout with no residual broken table.
- Automatic numbering changes are either absent or clearly reported for review.
- Saving and reopening a disposable cleaned copy does not introduce new problems.

Distinguish pre-existing issues from new ones by opening the original control. Do not save over the original. If Word is unavailable, report manual validation as outstanding and do not claim full Word compatibility. This is a validation limitation; do not invent a passing result.

### 14.4 Documentation edits

Update `README.md` and `CLAUDE.md` for every intentional behavior change. Preserve useful existing explanations rather than trimming them wholesale.

Required topics:

- Safer pattern scope and examples of ambiguous content now retained.
- Formatting-only removal now off by default, explicit opt-in, and existing configuration transition.
- Source-evidence verification and its remaining dependence on configured policy.
- Exact inline transformation pairing and part-aware comparisons.
- New reference/field/table safeguards and the limits of structural checks.
- Numbering notices without a promise of automatic cross-reference repair.
- Verified/needs-review/failed outcomes and the disposition of written outputs.
- Batch collision rejection.
- Correct test instructions: synthetic fixtures are generated programmatically, but package tests write temporary files.
- Updated CI coverage and developer benchmark instructions.

Correct contradictory statements encountered in touched sections, including claims of guaranteed intact structure, tail handling, configuration location in frozen builds, and the breadth of what a PASS proves. Do not rewrite unrelated documentation for style alone.

Add changes under an Unreleased section in `CHANGELOG.md`, following its existing Keep a Changelog structure. Do not rewrite historical release entries to describe new behavior and do not invent a new release version/date. Update `requirements.txt` only if dependencies actually change; the default expectation is no dependency change.

### 14.5 Final implementation report

The integrator must provide:

1. Completed work-package IDs and any explicitly deferred items.
2. A concise description of final behavior and intentional compatibility/default changes.
3. Automated test results and the precise coverage of injected-damage and package integration tests.
4. Before/after performance results with fixture, environment, and timing boundaries.
5. Corpus findings, if any, separated from synthetic results.
6. Word validation performed, or the exact outstanding manual checks.
7. Active-configuration transition behavior for installed users.
8. Remaining risks and known unsupported reference/numbering/OOXML cases.
9. Files changed and how each group relates to the implementation objective.

No release tag, push, publication, or deployment is part of this completion report.

## 15. Consolidated acceptance checklist

### Content and verification

- [ ] F01/F02 requirement examples survive under the shipped defaults.
- [ ] Unambiguous copyright and editorial positives still clean.
- [ ] Formatting-only removal defaults to off everywhere and explicit opt-in still works.
- [ ] Preserve styles protect complete loss and partial modification.
- [ ] Low-confidence prerequisites and detector switches govern verification.
- [ ] Partial hidden/editorial evidence cannot excuse losing unrelated requirement text.
- [ ] Inline patterns cannot excuse extra deleted requirement words.
- [ ] Long valid redactions pair correctly without a similarity loophole.
- [ ] Duplicate text, part identity, and output order do not conceal unexplained loss.
- [ ] Source-based evidence is independently obtained from actual input/output files.

### XML and reference behavior

- [ ] Covered separator elements are handled according to the documented redaction policy.
- [ ] Simple-field instructions/wrappers survive ordinary editorial removal.
- [ ] Loss of a supported field carrier is visible to verification.
- [ ] Internally referenced bookmark target content is protected.
- [ ] Existing broken references are distinguished from newly broken references.
- [ ] Revision-empty tables do not leave unsupported zero-row residue.
- [ ] Required parent block content and terminal cell paragraphs remain valid.
- [ ] Numbering-sensitive changes receive accurate review notices; no automatic renumbering occurs.
- [ ] Existing section, field, image, note, wrapper, and text-box tests still pass.

### GUI, configuration, and delivery

- [ ] Batch outputs are unique and cannot overwrite selected sources.
- [ ] Preflight rejects conflicts before any processing write.
- [ ] Written-but-unverified outputs are explicitly identified.
- [ ] Verification failure does not count as a verified success.
- [ ] Worker exception paths restore controls and continue independent files where appropriate.
- [ ] Verifier configuration fallback uses `apppaths` only when needed.
- [ ] Existing user configuration is never silently overwritten.
- [ ] Superseded active defaults and formatting-only opt-in are surfaced accurately.
- [ ] Test-only changes run CI.
- [ ] Performance improvement targets the measured main paragraph alignment cost.
- [ ] Documentation matches final behavior and historical changelog entries remain historical.
- [ ] Full tests, corpus coverage, and Word validation are reported honestly and separately.

## 16. Ready-to-use assignment text

### Integrating coding agent

Implement the required work in this plan against the current SpecCleanse checkout. Start with W00, compare the current code with the reviewed baseline, and establish a small source-evidence/result contract before delegating edits that depend on it. Preserve the flat architecture and current runtime dependencies. Treat the listed synthetic observations as reproducible leads; add actual DOCX package integration tests and independent damaged-output verifier tests. Integrate work in small coherent stages. Do not claim a verifier defect is an observed cleaner deletion. Do not treat faster runtime as permission to weaken verification. Complete W09 and report manual Word/corpus limitations explicitly. Publishing a release is outside scope.

### Policy and verification agent

Own W01–W03. Correct unsafe default whole-paragraph policy, make formatting-only removal opt-in for default configurations, and preserve removal scope/prerequisites/source intervals during verification. Add preserve-style evidence and part/story identity. Make exact permitted redactions work regardless of retained-text percentage, while injected extra loss still fails. Agree public metadata and outcome changes with the integrator before dependent agents code against them. Avoid a new rules framework and avoid using processor-produced targets as verification authority. Return tests, compatibility notes, and clear documentation changes.

### XML safety agent

Own W04–W05 after agreement on source protection and verification evidence interfaces. Fix covered separator elements, preserve simple fields, protect supported internally referenced bookmark targets, and handle tables emptied by accepted revisions without losing required parent blocks. Add focused numbering review notices without renumbering or rewriting references. Test actual XML structures and actual package processing, including nested/wrapped runs and text boxes. Verify relevant OOXML semantics and distinguish a custom lint result from observed Word behavior. Coordinate edits to shared modules with the integrator.

### GUI and CI agent

Own W06 and the agreed W07 files. Build and validate a complete batch manifest before writes; reject destination collisions and destinations that overwrite selected inputs. Replace conflated success reporting with explicit verified/needs-review/failed outcomes, identifying any written but unverified output. Preserve Tkinter threading discipline. Ensure test-only changes receive CI and that verifier fallback configuration is covered through the shared resolver in coordination with its owner. Do not add an unrelated GUI redesign or settings system.

### Performance agent

Own W08 only after the corrected verifier is integrated. Reproduce the large-input bottleneck with phase-separated profiling; the earlier 6,000-paragraph probe spent about 1.3% of verification time in survivor pairing, so prioritize the outer paragraph alignment. Preserve scope, part identity, order, multiplicity, and all injected-damage failures. Benchmark the original repeated-heading shape and additional adversarial shapes. Optimize container scans only with a readable equivalent safety predicate. Return measured before/after results and stop before speculative ZIP/concurrency work.

## 17. Primary references for standards-dependent checks

These sources support specific implementation questions; they do not replace testing the actual supported Word version and package representation.

- [GitHub Actions workflow syntax and path filters](https://docs.github.com/en/actions/reference/workflows-and-actions/workflow-syntax#patterns-to-match-file-paths): root-level `*` versus recursive matching.
- [Microsoft Open XML simple fields](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.simplefield?view=openxml-3.0.1): `w:fldSimple` and its field instruction.
- [Microsoft: working with WordprocessingML tables](https://learn.microsoft.com/en-us/office/open-xml/word/working-with-wordprocessingml-tables): table, row, and cell structure; verify edge-case acceptance in Word before claiming compatibility.

When implementing bookmark field parsing, numbering inheritance, or break semantics, consult the applicable Microsoft/OOXML primary documentation and record the supported subset. Avoid inferring a general schema rule from one synthetic fixture.
