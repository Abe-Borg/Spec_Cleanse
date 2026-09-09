# SpecCleanse implementation report

Final report for the work packages defined in `SPECCLEANSE_IMPLEMENTATION_PLAN.md`,
required by §17.5.

**No release tag, push, publication, or deployment is part of this report.** The
work is on `master`; cutting a version remains a separate decision.

One structural caveat before anything else. §6 assigns implementation and
independent review to separate parties, and §17.5 asks the two to report
jointly. **Both roles were carried out by the same agent**, so the independence
that separation exists to provide was not present. What partly substituted for
it was an automated reviewer (Codex) on every pull request, which found a real
defect on eight of nine packages — including two false negatives that would have
reported damaged output as verified, one that would have led a user to delete a
good document, and one configuration value that crashed every file. That record
is the argument for the separation, not evidence it was unnecessary.

---

## 1. Work packages

| ID | Scope | State |
|---|---|---|
| W00 | Immediate corrections: batch manifest rejection, outcome separation, separator redaction, exact permitted redaction | Complete |
| W01 | Baseline, fixtures, census and corpus utilities | Complete; **census never run** (§5) |
| W02 | Scoped default policy and the formatting-only decision | Complete; **default chosen without corpus evidence** (§6) |
| W03 | Source-evidence verification contract | Complete; **acceptance criteria not evaluable** (§6) |
| W04 | Pairing, part identity, and ordering | Complete |
| W05 | Inline XML redaction and field carriers | Complete |
| W06 | Reference integrity, revision-emptied tables, numbering notices | Complete; **notice rate on real files unknown** |
| W07 | Full batch and outcome model | Complete |
| W08 | Configuration fallback, validation, CI | Complete; §15.3 was already satisfied by an earlier package |
| W09 | Measured performance improvements | Complete; **all fixtures synthetic** |
| W10 | Integrated validation and documentation | Complete except §17.2 and §17.3, both outstanding for want of inputs |

**Explicitly deferred**, unchanged from the plan's own deferral list (§3.3):
automatic retention of referenced bookmark ranges, cross-reference repair,
renumbering, and any structural or style optimisation stage. Nothing was
deferred beyond that list.

## 2. Final behaviour, and what changed for existing users

The cleaner still performs a single-pass content removal followed by an
automatic verification pass, and still touches tracked changes and comments only
when the user opts in. What changed:

- **Narrower removal rules.** Several shipped patterns were narrowed because
  they deleted requirements. A user's own `patterns.yaml` is never overwritten,
  so an installed copy keeps the older, broader rules — `config_notices()`
  reports each superseded rule that is still active, matched on the exact prior
  string so someone who edited a rule themselves is not told their own work is
  stale.
- **`specifier_notes.formatting_only_removal` defaults to `false`.** This is the
  one path that removes text on no content evidence at all, and real
  specification text is routinely red and italic. Users who want the old
  behaviour opt in; users who already have it on are told so, and every file in
  such a run is reported as needing review.
- **Verification judges against source evidence** rather than a flattened list
  of patterns, and never asks the processor what it did.
- **Three outcomes, and Needs-review names its category** — ambiguous alignment,
  detected damage, configuration notice, or reference/numbering warning. A
  failure states whether a file was nonetheless written.
- **New safeguards**: simple fields protected, nested field carriers counted
  separately, references this run broke reported per consumer, tables emptied by
  accepted revisions removed, numbering-sensitive removals noticed.
- **Configuration is validated** where a mistake used to fail silently: quoted
  Booleans are refused, colours must be six hex digits with a leading `#`
  normalised away.
- **Batches are rejected before any write** when two inputs would collide on one
  destination, or an output would overwrite a selected input.

**No dependency changed.** `requirements.txt` is byte-identical to its state
before W00, as §17.4 expects.

## 3. Automated test results

| | Linux 3.11 (this workspace) | Windows 3.12 (CI) | Linux 3.10 (CI) |
|---|---|---|---|
| Passed | 446 | 446 | 446 |
| Skipped | 15 | 0 | 0 |
| Expected failures | 0 | 0 | 0 |
| Unexpected successes | 0 | 0 | 0 |
| Failures / errors | 0 | 0 | 0 |

Against the W01 baseline of **92 passed, 4 skipped** on Linux with Python 3.11
and lxml 6.0.2: **+354 tests**. The 15 skips are the GUI tests, which skip where
`tkinter` is absent; `actions/setup-python` ships it, so both CI lanes run them
and report no skips at all. Expected failures returned to zero as §8.1 requires —
W00 added one under `unittest.expectedFailure`, and W03 closed it with the
unexpected success that said the decorator could go.

**GUI tests skipped on Linux establish nothing about Windows GUI behaviour**, and
this mattered concretely: during W07 a local run reported `Ran 349 tests ... OK
(skipped=9)` while four GUI tests were already broken. Only a run under a stubbed
`tkinter` showed it. Read the skip count.

Coverage of the two layers §17.1 singles out:

- **Injected-damage verification** — `tests/test_verify.py` carries the twelve
  cases V01–V12, each building source and damaged output independently, never by
  running the cleaner. Eleven must fail; V09 is the false-alarm guard that a
  *correct* clean still passes, which is what stops the contract being satisfied
  by a verifier that distrusts everything. `tests/test_alignment.py` extends the
  same discipline to W09's fast path, comparing every verdict against the
  ordinary comparison in both directions.
- **Synthetic package integration** — every processor and verifier test builds a
  real `.docx` ZIP, runs the public path, and rereads the output XML. Nothing
  mocks the parser, writer, or structural inspection.

W10 added a layer the plan implies but did not have: the fixtures were checked as
*documents*. Two defects surfaced, both invisible to every existing test and both
disqualifying for the Word validation in §17.3.

- Every generated part named `w14` in `mc:Ignorable` without declaring it, which
  Markup Compatibility does not permit.
- Every package carried a relationship to `word/comments.xml` whether or not that
  part was present — a dangling relationship, which is precisely what
  `test_revisions.test_package_bookkeeping_is_updated` asserts the *cleaner* must
  never leave behind. The project already treated this as package-breaking when
  its own code did it, and did it by construction in every fixture.

The second was found by review, after the first round of this work had asserted
the fixtures were valid. The test written to establish that had checked member
*names* were present and never that relationships resolved — which is the same
error one level up, and worth recording: checking the defect you just thought of
is not the same as checking the class it belongs to.

## 4. Performance

Measured with `tools/benchmark_pipeline.py` on synthetic fixtures. Seconds,
median of unprofiled runs, **one machine and one interpreter** (Python 3.11,
CPython, Linux x86-64). These are review targets, not thresholds asserted in CI.

| family | 2,000 before | 2,000 after | 12,000 after |
|---|---|---|---|
| `adversarial_headings` (verify) | 21.192 | **0.203** | 1.38 |
| `repeated_requirements` (verify) | 21.571 | **~0.20** | — |
| `realistic_requirements` (verify) | 1.049 | **0.327** | 2.15 |
| `adversarial_headings` (clean) | 0.252 | **0.090** | 0.50 |

Growth on the adversarial family fell from roughly n^2.8 to roughly n^1.5.
**The realistic family's 3.2× is the number that matters most**: §16.4 forbids
gating on the adversarial family, which repeats three strings through an entire
document and is the worst case by construction rather than a plausible one.

Boundaries, stated because they bound the claim:

- The **12,000-paragraph baseline was never run.** By the measured curve it would
  have taken roughly three quarters of an hour. The improvement factor at that
  size is extrapolated, not observed.
- Every fixture is synthetic. No real specification was timed.
- Small documents show no regression at 50 and 200 paragraphs over five runs.

## 5. Census A and B, and corpus findings

**Not run. No specification corpus was available in this workspace at any point.**

The tools exist and are tested — `tools/census_formatting`, `tools/census_references`,
`tools/corpus_compare`, `tools/actions` — and each is pointed at a directory of
`.docx` files the maintainer supplies. None has been pointed at one.

Consequently, and stated rather than glossed:

- **Census A** (what turning formatting-only removal off costs) never produced a
  number. The default was chosen on the argument that removing text with no
  content evidence is the wrong failure to risk, not on measurement.
- **Census B** (removals inside referenced bookmark ranges) never produced a
  number. Automatic retention stays deferred partly because its cost is unknown.
- **The §17.2 corpus diff was never run**, so no decision-level comparison
  between the baseline and candidate builds exists.

No precision or recall figure appears anywhere in this work, and none should be
inferred. §17.2 permits finishing the synthetic and code work with corpus
validation labelled outstanding; that is what happened.

## 6. §10.6 acceptance criteria

| # | Criterion | State |
|---|---|---|
| 1 | Every independently judged correct output is reported Verified | **Not evaluable** |
| 2 | No unexplained new failure relative to the W01 baseline | **Not evaluable** |
| 3 | Ambiguous alignment, damage, configuration and reference/numbering measured separately | **Machinery complete, never measured** |
| 4 | All injected-damage tests continue to fail for their stated reason | **Met** |
| 5 | Any permitted exception individually documented | **Not evaluable; none claimed** |

Criteria 1, 2, 3 and 5 take their denominator from an acceptance set of
independently judged correct outputs on real documents. There is no such set.
Criterion 3's four-way split now exists in code and reaches the user, which was
this project's non-gating obligation — but a category that has never been counted
on real files is a prerequisite met, not the criterion met.

**No permitted exceptions are claimed, because no acceptance run occurred.**

What exists instead is the property those criteria protect, asserted on synthetic
documents: V09 requires that a correct clean still passes, and the whole suite is
a standing false-alarm check since every clean in it must verify. That is weaker
than the criteria ask for and should not be described as meeting them.

## 7. Word validation

**Not performed. No Windows Word was available.**

Every check in §17.3 is outstanding, and they are listed here so they can be run
rather than assumed:

1. Open original and cleaned copies of representative documents — no repair or
   unreadable-content prompt on the cleaned copy that the original does not also
   produce.
2. Requirement text, headings and section boundaries present.
3. Headers, footers, note anchors, pictures and fields behaving.
4. Simple fields surviving; references refreshed in a disposable copy and
   inspected.
5. Referenced bookmark targets: newly broken references were reported, and no
   *unreported* new broken reference appears.
6. **Last-row and last-cell revision cases.** §17.3 singles this out as the case
   that most needs Word rather than lint, and W06 proved the point — a table
   emptied of its last row passed every structural check in this project until
   the lint was given a rule for it. A clean lint is not evidence Word will open
   a file.
7. Numbering changes absent or clearly reported.
8. Saving and reopening a disposable cleaned copy introduces nothing new.

Open the original as a control, to separate pre-existing problems from new ones,
and do not save over it. Until this is done, **no claim of Word compatibility is
made by this project.** CI builds the Windows executable and runs the suite on
Windows; neither opens a document in Word.

## 8. Active-configuration transition for installed users

An existing `patterns.yaml` is never replaced. `apppaths.resolve_config_path()`
returns, in order: the file beside the modules when running from source; a copy
beside the `.exe` if one is there; otherwise a per-user copy under
`%APPDATA%\SpecCleanse`, seeded from the bundled default on first run and
thereafter the user's own. Seeding is best-effort — a read-only profile falls back
to the bundled copy rather than raising, because a traceback from a windowed
build goes nowhere anyone can see.

The consequence is deliberate and worth stating: **an update does not narrow an
installed user's rules.** A copy made before a rule was narrowed keeps the older,
broader rule, and would go on deleting what that rule deletes. `config_notices()`
is the entire mechanism by which they learn this, and W07 made it consequential —
a run under such a configuration reports every file as needing review, naming
"configuration notice" as the reason.

Stricter validation also arrives with an update: a configuration that previously
loaded with `formatting_only_removal: 'false'` or a colour written `"bright red"`
now fails to load, with a message naming the section, key and expected shape.
That is a deliberate compatibility break — both values were being silently
misread — but it is a break, and an installed user meets it as a startup error.

## 9. Remaining risks and unsupported cases

**Unmeasured on real documents.** The largest risk is the one running through
§5 and §6: every policy decision here was taken on argument and synthetic
evidence. The specific unknowns are the formatting-only default's cost, the
false-alarm rate of each Needs-review category, and how often numbering notices
fire on a heavily numbered master specification. If Needs-review turns out to be
dominated by ambiguous alignment, that is a reason to revisit the comparison; if
by configuration notices, the finding is about a user's `patterns.yaml`.

**Verification is a consistency check, not an independent one.** It shares its
compiled patterns with the cleaner, so a rule that removes the wrong thing is
reported as expected. The genuinely independent layers are the formatting signals
read from the source and the structural inspection.

**Structural checks cover the shapes they have rules for.** The emptied-table
case is the standing demonstration that a clean lint reports nothing when it has
no rule, which is not the same as reporting that nothing is wrong.

**Known unsupported or deliberately unhandled cases:**

- `.doc` (legacy binary) is not handled at all.
- A reference whose target this run broke is reported, never repaired: no target
  is invented, no field retargeted, no replacement bookmark created.
- Numbering is never renumbered and no cross-reference is rewritten. A notice
  says the numbers a reader sees may differ; it does not assert that a specific
  reference broke.
- Field results are never recalculated. What is promised is that the instruction
  and its wrapper survive to recalculate from.
- A redaction spanning a page or column break is abandoned rather than
  half-applied, leaving the placeholder in place — the lesser cost, since
  stranding a break reflows the document from that point on.
- A table that arrived with no rows is linted but not removed: it is the
  document's own problem, and rewriting a file that had no revisions to accept
  would be the cleaner doing something unrelated to its task.
- `w:altChunk`, embedded objects beyond the tags in `EMBEDDED_CONTENT_TAGS`, and
  content controls holding block content are not specially reasoned about beyond
  the paragraph-protection rules.

## 10. Files changed

41 files, +12,507 / −503 against the pre-W00 tree. By purpose:

| Group | Files | Relation to the objective |
|---|---|---|
| Detection policy | `detection.py`, `patterns.yaml` | Narrower rules; the source-evidence types that replaced a flattened pattern list; configuration notices |
| Verification | `verify.py` | The whole source-evidence contract, part-aware comparison, references, fields, tables, numbering, review categories, and W09's alignment work |
| Document plumbing | `docx_xml.py` | Signatures, field instructions, reference targets, numbering resolution, container rules, configuration validation |
| Processing | `processor.py` | Separator redaction, layout-break abandonment, field-carrier protection, revision acceptance |
| Outcome and batch | `batch.py`, `gui.py` | Collision rejection, the three outcomes and four review categories, the run loop moved where it can be tested |
| Runtime locations | `apppaths.py` | Where `patterns.yaml` lives from source and in a frozen build |
| Measurement | `tools/` (5 modules) | Census, corpus comparison, action inventory, pipeline benchmark — none imported by the application |
| Tests | `tests/` (22 modules) | 92 → 446 |
| Packaging and CI | `packaging/`, `.github/workflows/` | Windows build; the test workflow that runs on every pull request |
| Documentation | `README.md`, `CLAUDE.md`, `CHANGELOG.md`, the plan, this report | Behaviour changes, and the corrections in §17.4 |

## 11. What a reader should not conclude

- That the cleaner has been validated against real specifications. It has not.
- That a PASS means Word will open the file. It does not; §7 is outstanding.
- That the performance figures transfer to another machine, or to real documents.
- That the §10.6 criteria are met. Four of five are not evaluable here.
- That an independent reviewer signed this off. One agent did both roles.
