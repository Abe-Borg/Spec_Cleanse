# SpecCleanse — Review Brief for an Independent Reviewer

**Audience:** a capable reasoning model asked to review this codebase end to end.
**Author:** a prior agent that read every module, ran the suite, profiled the pipeline, and probed edge cases.
**Status of this document:** a starting point, not a verdict. Everything below is a hypothesis with the evidence I have. Verify it, contradict it, and go past it.

---

## 0. Ground rules, first

**You are not being asked to change anything.** You are being asked to reach your own conclusions.

If, after your own reading, you conclude the software is fine as it stands, say so plainly and explain why the concerns raised below do not warrant action. That is a fully acceptable outcome and a more useful one than a manufactured improvement plan. A refactor that trades a working, well-documented, well-tested program for a marginally faster one that its author no longer recognises is a net loss.

What *is* asked of you:

1. Read the code yourself. Do not take my findings on faith — several are stated confidently and I could be wrong about any of them.
2. Explore beyond this brief. My checklist in §5 is what I could think of; it is not the boundary of what matters.
3. Deliver **a detailed report** of what you found: what is sound, what is not, what is uncertain, with evidence.
4. Deliver **a proposed implementation plan** — *if and only if* one is warranted. Sequenced, scoped, with an explicit statement of what each change buys and what it risks. If parts of it are not worth doing, say which and why.
5. Rank by expected value. This is a single-maintainer tool used in professional practice; a hundred-file refactor is not a serious proposal, and neither is a list of style nits.

The author is a mechanical/fire-sprinkler designer with a CS degree who builds and maintains this himself, on Windows, in his own time. Optimise your recommendations for *a competent solo maintainer keeping this alive for years*, not for an imaginary team.

---

## 1. What this program is, and why correctness is asymmetric

SpecCleanse strips editorial noise out of construction **specification documents** (`.docx`) so that what a downstream LLM ingests is the specification, not the publisher's commentary about it.

The input documents are master-spec templates and their project-edited descendants — **MasterSpec / ARCOM, BSD SpecLink**, and firm-internal masters derived from them. These documents interleave, in the same file and often in the same list, two entirely different kinds of text:

* **Requirements** — the contract. "Sprinkler piping shall be Schedule 40 black steel." Losing one of these is a professional liability event: an unbuilt requirement, a missed submittal, a bid that omits scope.
* **Editorial apparatus** — instructions to the person editing the master. `Retain or delete paragraph below.` `[Verify quantity with Owner].` `Copyright 2026 by The American Institute of Architects.` Specifier notes, hidden text, SpecAgent watermarks, retain/delete instructions, fill-in placeholders. Keeping one of these is a nuisance: it wastes downstream tokens and can confuse an LLM into treating an instruction as a requirement.

**The error costs are wildly asymmetric.** A false negative costs tokens. A false positive can delete a fire-protection requirement from a document that goes out for permit and construction on a hyperscale data center. Every design decision in this program should be read against that asymmetry, and so should every recommendation you make. When in doubt, the reading that keeps text is the correct one — the codebase already says this in several places, and it is right.

The current pipeline is deliberately **single-pass, shallow, content-only**. Earlier structural-optimisation and style-pruning stages were retired and archived in `legacy/`; they changed document metadata without improving extracted text. Do not propose bringing them back without a strong argument.

---

## 2. Architecture, module by module

Flat layout, no package, no `pyproject.toml`. Entry point is `python gui.py`. Two runtime dependencies (`lxml`, `PyYAML`); everything else is stdlib. Direct `lxml` XML manipulation rather than `python-docx`, on purpose — the program needs control over structures `python-docx` abstracts away.

```
input.docx
  → unpack ZIP to temp dir → load word/styles.xml into a StyleIndex
  → for each content part (document, header*, footer*, footnotes, endnotes, glossary):
      parse → [optionally accept tracked changes / strip comments]
      → walk paragraphs → run detectors → collect targets
      → apply: redactions, then run removals, then paragraph removals
      → write part back only if it changed
  → repack to a temp file, os.replace into place
  → verify: re-unpack BOTH input and output, diff paragraph text, classify
            every removal and modification, compare structure
  → output_cleaned.docx + a PASS/FAIL log
```

### `docx_xml.py` (840 lines) — WordprocessingML plumbing
Namespaces, tag constants, structure rules, text extraction, span arithmetic, style resolution, revision handling, config loading. **Holds no detection policy** — that separation is stated in `CLAUDE.md` and is actually honoured. Key exports:

* `iter_paragraphs(root, skip_alternate_fallback)` — every `w:p` in document order. `mc:AlternateContent` stores the same shape twice (`mc:Choice` / `mc:Fallback`); processing walks both, text extraction counts one.
* `iter_own_runs(para)` — runs belonging to *this* paragraph, descending through `w:hyperlink`, `w:ins`, `w:sdtContent`, `w:fldSimple`, `w:smartTag`, but **not** into a paragraph nested in a text box (that inner paragraph is visited in its own right). Using `para.iter(w:r)` instead would double-count text-box content.
* `iter_text_nodes(scope)` / `element_text` — yields `(element, text)` so character offsets map back onto the nodes they came from, which is what in-place redaction depends on. Tabs, breaks and non-breaking hyphens contribute real whitespace, so `\s+` patterns and `^`-anchored preserve patterns behave as written. `w:delText` and `w:instrText` are deliberately not extracted.
* `is_on()` / `toggle_on()` — WordprocessingML **toggle** semantics. `<w:vanish/>` is on; `<w:vanish w:val="0"/>` is off. Presence is not truth.
* `StyleIndex` — resolves a style ID through its `w:basedOn` chain, matching folded style **IDs and display names**. `is_hidden()` implements toggle XOR along the chain (a derived style repeating its base's `w:vanish` renders *visible*), and treats an explicit `w:val="0"` anywhere in the chain as off.
* Structure predicates — `field_chars_balanced`, `has_section_properties`, `has_embedded_content`, `can_delete_paragraph`, `orphaned_range_markers`, `block_children`.
* `accept_revisions` — accepts tracked changes. Note the genuinely subtle part: a deleted table **row** keeps ordinary `w:t` and records the deletion only in `w:trPr`, so it must be removed whole; only run-level deletions hide in `w:delText`.
* `load_config` — reads `patterns.yaml` as **UTF-8** explicitly (the file contains `©`, `–`, `—`; cp1252 on Windows would silently mangle them into patterns that match nothing) and compiles every regex up front so a bad one is reported once, with its section and index.

### `detection.py` (653 lines) — the policy layer
`BaseDetector` strategy hierarchy: `SpecifierNoteDetector`, `CopyrightDetector`, `HiddenTextDetector`, `SpecAgentDetector`, `EditorialArtifactDetector`, `PreserveDetector`. `DetectionEngine` orchestrates and owns the compiled patterns.

Confidence model, with evidence deliberately split into two buckets:
* **Content evidence** (what the text *is*): specifier-note pattern 0.6, copyright 0.7, high-confidence editorial 0.8, editorial style 0.8, hidden/SpecAgent/preserve 1.0, low-confidence editorial 0.3.
* **Formatting evidence** (how it *looks*): italic +0.2, editorial colour +0.3.
* Threshold 0.5. Italic + colour is **exactly** 0.5 — so formatting alone can carry a removal, governed by `specifier_notes.formatting_only_removal` (default `true`). Such removals carry `Detection.formatting_only` and are labelled in Preview.
* Preserve patterns and preserve styles short-circuit removal regardless of score.

`editorial_artifacts` has three tiers: `text_patterns` (whole paragraph on text alone), `low_confidence_patterns` (need a formatting signal to cross 0.5), `inline_patterns` (produce `INLINE_PLACEHOLDER` detections with character spans, cut out in place — `should_remove()` ignores them). `removal_patterns(include_inline=False)` exists so verification can ask the *narrower* question of the inline tier: an inline match never justifies losing a whole paragraph, only "cutting every placeholder leaves nothing behind" does.

### `processor.py` (502 lines) — unpack, remove, redact, repack
Targets are collected during the walk and applied afterwards (mutating a tree while iterating it skips elements), in a fixed order: **redactions → runs → paragraphs**, because redaction spans are offsets into the paragraph text as it stands.

`_remove_paragraph` refuses to delete the `w:p` element in four situations, emptying it in place instead:

| Situation | Why | Check |
|---|---|---|
| A field begins inside and ends later | Unbalanced `w:fldChar` is a file Word won't open | `field_chars_balanced()` |
| The paragraph carries `w:pPr/w:sectPr` | Deleting it merges the section into the next, losing its headers, footers and page setup | `has_section_properties()` |
| It holds a picture, object, field, or note reference | Invisible to text patterns, visible to the reader | `has_embedded_content()` |
| Its parent would be left with no block content | `CT_HdrFtr` etc. carry `minOccurs="1"`; a cell must also *end* with a `w:p` | `can_delete_paragraph()` |

Half-open bookmark and comment-range markers are relocated beside the paragraph before deletion. Repacking writes to a `.tmp` beside the destination and `os.replace`s it, so an interrupted run can never leave a truncated `.docx`. This whole layer is thoughtful and I found no correctness bug in it.

### `verify.py` (711 lines) — the safety net
Re-reads input and output and asks three questions: which paragraphs disappeared and does each match a rule meant to remove it; which surviving paragraphs lost text and was the loss asked for; is the output still structurally valid.

The module's own docstring is honest about its central limitation, and you should hold it to that: **the pattern layer shares compiled patterns with the detection engine, so it agrees with the cleaner by construction.** A rule that removes the wrong thing is reported as *expected*. Only two layers are genuinely independent: the formatting signals read from the source DOCX, and the structural inspection.

`passed` is `True` only when unexpected removals, unexpected modifications, preserve violations, structural violations and added paragraphs are all empty. The verdict is **advisory** — a FAIL still writes the output file.

### `gui.py` (705 lines) — Tkinter front end
Worker thread, queue-drained log, progress bar, overwrite confirmation, a "file is open in Word" error path, controls disabled during a run with inputs snapshotted on the main thread. Competent threading discipline. This is the only entry point: **there is no CLI.**

### `apppaths.py` (104 lines) — where `patterns.yaml` lives
From source, beside the modules. Frozen: a copy beside the `.exe` wins, else a per-user copy under `%APPDATA%\SpecCleanse` seeded from the bundled default on first run, best-effort (a read-only profile falls back to the bundled copy rather than raising, because a traceback from a windowed build goes nowhere anyone can see). Correct and well-tested.

### `tests/` — 92 stdlib `unittest` tests, all passing
`tests/docx_builder.py` assembles synthetic `.docx` files with `zipfile`, so structural edge cases are covered with no binary fixtures and no test dependency. This is genuinely good test design; do not propose replacing it with pytest without a concrete reason.

```
$ python -m unittest discover -s tests -t .
Ran 92 tests in 0.849s — OK (skipped=4)   # 4 GUI tests skip without tkinter
```

---

## 3. What I actually measured

Reproduce before you trust. Synthetic documents, ~1 in 12 paragraphs editorial, ~1 in 12 carrying an inline placeholder, built with `tests/docx_builder.py` (script in §7).

| Body paragraphs | Clean | Verify |
|---:|---:|---:|
| 500 | 0.06 s | 0.03 s |
| 2,000 | 0.31 s | 0.29 s |
| 6,000 | 1.42 s | 4.70 s |
| 12,000 | 5.02 s | **35.52 s** |
| 24,000 | 16.80 s | **274.69 s** (4.6 min) |

Clean scales at roughly **n^1.75**; verify at roughly **n^2.95** — doubling the document from 12k to 24k paragraphs multiplied verification time by 7.7. A single MasterSpec section is small (hundreds of paragraphs) and neither number matters. A consolidated project manual — which is how a full spec book is often exported, and is exactly the artifact you would hand an LLM — is 20,000–60,000 paragraphs, where verification already takes minutes and is still accelerating away from the cleaner it is supposed to be checking.

`cProfile` at 6,000 paragraphs:

* **Verify:** `difflib.find_longest_match` — 12.7 s of ~16 s total, 3,500 calls, 44.8 M `dict.get` calls. Everything else is noise.
* **Clean:** `docx_xml.py:343` (`block_children`'s list comprehension) — 0.716 s across **500 calls**. That is the single largest cost in the cleaner, and it comes from one line.

---

## 4. My candidate findings

Each has a claim, the evidence I have, and my confidence. **Treat these as leads to verify, not conclusions to implement.**

### 4.1 `can_delete_paragraph` is O(removals × body size) — *high confidence, measured*

`processor._remove_paragraph` → `docx_xml.can_delete_paragraph(para)` → `block_children(parent)`, which materialises **every block child of `w:body`** to answer "is at least one other block child left, and does a cell still end with a paragraph?" That is a full scan of the document body, once per removed paragraph. At 6,000 paragraphs / 500 removals it is 3 M element visits and the top line in the clean profile. At 30,000 paragraphs / 3,000 removals it is ~90 M.

Both questions are answerable in O(1) from sibling links (`getnext()` / `getprevious()`) without building a list. Worth checking: is there a formulation that is both O(1) *and* obviously as safe as the current one? The current version's virtue is that it is trivially easy to read and verify against the schema rule it encodes, and that virtue is worth real CPU time. **Judge whether the speed is worth the legibility.**

### 4.2 Verification runs two character-level diffs per candidate pair — *high confidence, measured*

`verify._pair_with_survivor`, for each unmatched input paragraph, scans every unpaired output paragraph in the replace block and for each candidate runs:

1. `_removed_fragments(text, after)` — a full `SequenceMatcher.get_opcodes()` over characters, then
2. `difflib.SequenceMatcher(...).ratio()` — **a second full character diff of the same pair**, to compute a similarity it already has enough information to derive.

For a pure deletion, every character of `after` survives, so `ratio == 2·len(after) / (len(before) + len(after))` — computable in constant time from the opcodes already in hand. Dropping the second matcher should roughly halve the cost with no behavioural change.

The outer loop is separately O(k²) in the size of a replace block, each iteration paying a character-level diff. Cheap prefilters that cannot change the outcome: `len(after) <= len(before)`; and since a pass requires `ratio >= 0.5`, `len(after) >= len(before)/3`. Both are O(1) and would eliminate nearly every candidate before the expensive diff. Beyond that, consider whether pairing needs to scan the whole block at all, or only a bounded window around the diff alignment.

### 4.3 A correct inline redaction can be reported as a FAIL — *high confidence, reproduced*

`MIN_PAIR_SIMILARITY = 0.5` means a paragraph that loses more than about half its characters to redaction is not recognised as the same paragraph. It is then reported **both** as an unexpected removal **and** as an added paragraph, and the run fails.

Reproduced exactly:

```python
# Input:  "Provide [Verify quantity with the Owner and the AHJ prior to bid] units."
# Output: "Provide units."                    <- the processor did the right thing
# Verify: pass=False, unexpected_removals=1, added=1
#   UNEXP RM: 'Provide [Verify quantity with the Owner and the AHJ prior to bid] units.'
#   ADDED:    'Provide units.'
```

MasterSpec writes placeholders that dominate short paragraphs routinely, so this is not exotic. The cost is not a wrong document — the output is correct — it is **a false alarm that teaches the user to stop reading the verification log**, which is the one thing standing between a bad pattern and a lost requirement. I consider that the most damaging kind of bug this program can have.

Worth thinking about: should pairing use the redaction spans the processor actually applied, rather than re-deriving similarity blind? The information exists on the `ProcessingResult`; verification currently throws it away and re-diffs from scratch. That is a deliberate design choice (verification stays an input/output comparison and cannot be fooled by a processor that lies about what it did) and there is a real tension here — resolving it in favour of speed may cost independence. Reach your own conclusion.

### 4.4 Copyright patterns fire on real specification prose, and verification blesses it — *high confidence, reproduced*

`CopyrightDetector` is the only detector with no formatting or style gate: a text match alone scores 0.7 and takes the whole paragraph. Combined with the global `re.DOTALL` in `compile_patterns` and unbounded `.*?`, these fire on ordinary AEC language:

| Real spec sentence | Pattern that removes it |
|---|---|
| "Shop Drawings submitted under this Section **may not be reproduced** for use on other projects." | `may\s+not\s+be\s+reproduced` |
| "Contractor shall verify that **duplication** of sprinkler coverage in adjacent zones is **prohibited** by the AHJ." | `duplication.*?prohibited` |
| "**Unauthorized** personnel shall not have access to the fire pump room; **reproduction** of access keys is not permitted." | `unauthorized.*?reproduction` |

The middle two match across an entire sentence because `.*?` under `DOTALL` is unbounded — the two anchor words need not be related at all. And because verification classifies against the same patterns, **every one of these is reported as an *expected* removal and the run reports PASS.** I verified this end to end: `copyright_fp` → `pass: True, removed: 1, unexpected: 0`.

Candidate mitigations, in rough order of my confidence: bound the gaps (`.{0,40}?`); drop the global `DOTALL` or apply it per-pattern; consider whether the copyright detector should require a formatting/style/position signal the way `low_confidence_patterns` do, or whether copyright boilerplate is reliably enough located (headers, footers, first/last paragraph of a part) to gate on position. Note the tension: MasterSpec copyright *is* usually plain body text, so a formatting gate may cost real recall. **This one needs domain judgement, not just regex craft.** The maintainer has the corpus; you do not.

### 4.5 `formatting_only_removal` defaults to `true` — *design question, not a bug*

Italic + editorial colour sums to exactly the threshold, so **any** red-or-blue italic run is deleted with no pattern match at all. It is documented, switchable, labelled in Preview, and verification honours the same switch. But it is the single most aggressive default in the program, and in firm-edited masters, red italic is also a common convention for *project-specific edits that must be kept*. Consider whether "safe by default, opt into aggression" is the better polarity for a tool whose false positives are expensive. Consider also whether it deserves a GUI checkbox next to the tracked-changes one, rather than living only in a YAML file.

### 4.6 No numbering awareness — *confirmed absent; importance is a domain call*

Nothing in the codebase reads `w:numPr` or `word/numbering.xml`. Specifications are numbered documents whose own cross-references cite paragraph numbers ("as specified in Paragraph 1.6.B"). Deleting a paragraph that participates in a numbering sequence **renumbers everything after it**, silently invalidating those references.

In practice MasterSpec's `CMT`/note styles are usually outside the requirement numbering, which is probably why this has not bitten. But I confirmed the program removes a `w:numPr`-carrying paragraph without any warning, and verification reports PASS. The cheap version is not a fix but a *signal*: notice when a removed paragraph carried `w:numPr` and say so in the log. Decide whether that is worth the noise.

### 4.7 A bookmark wholly inside a removed paragraph is dropped silently — *confirmed*

`orphaned_range_markers` relocates only **half-open** ranges. A bookmark whose start *and* end are inside the doomed paragraph is deleted with it, and any `REF` field elsewhere pointing at it becomes "Error! Reference source not found." in Word. Verification does not notice; my `bookmark_target_lost` probe reports PASS. Whether this matters depends on whether editorial paragraphs are ever bookmark targets — probably rare, but the failure is invisible and lands in a document that goes out the door.

### 4.8 `verify.py` bypasses `apppaths` — *low severity, real*

`verify.py:538` hardcodes `config_path = Path(__file__).parent / "patterns.yaml"`. In a frozen build that resolves inside PyInstaller's temporary extraction directory — precisely the bug `apppaths.py` exists to fix, and which `CHANGELOG.md` records as fixed. It is latent because the GUI always passes `engine=`, so the fallback never runs today. It is a trap for the next caller, and the CI cannot catch it. Two lines.

### 4.9 There is no CLI — *the biggest missing capability, in my judgement*

Everything runs through Tkinter. The engine, processor and verifier are all cleanly importable and the wiring is already done in `gui.py`'s `_clean_one` / `_preview_one` — a `__main__.py` or `cli.py` is on the order of 60 lines. Without it:

* No batch cleaning from a script or scheduled task.
* No integration into a document-ingestion pipeline, which is the stated purpose of the program.
* No way to regression-test pattern changes against a real corpus of spec files.
* No headless use on a build server.

That last one matters more than it sounds. **The single most valuable thing that could be built on top of this program is a corpus regression harness**: run the cleaner across a folder of real spec sections, record every removal, and diff that record against the last run. Then a pattern edit shows its blast radius before it ships, and §4.4-class problems surface as a diff rather than as a lost requirement. A CLI is the precondition.

### 4.10 Smaller things I noticed

* **Redundant work per paragraph.** `_process_paragraph`, `_should_remove_paragraph` and `_group_run_detections` each independently call `iter_own_runs()` and re-extract `run_text()` for the same paragraph — text extraction runs three to four times per paragraph. Fixable by threading one computed list through, at some cost in signature noise.
* **Per-run allocation in the hot loop.** `SpecifierNoteDetector.detect` rebuilds `[c.upper() for c in fmt_colors]` for **every run in the document**. Precompute a folded `frozenset` on `PatternConfig` once.
* **Repeated subtree walks.** `has_embedded_content` and `field_chars_balanced` each walk a full subtree, and are called per removal candidate and again per touched run during redaction.
* **`_classify_modification` reports `verdicts[0][0]`** as the category even when different fragments were classified differently — a reporting inaccuracy, not a safety one.
* **Element identity as a set key.** `_group_run_detections` and `_should_remove_paragraph` rely on `Detection.element` being the *same lxml proxy object* returned by a later `iter_own_runs()`. lxml guarantees this only while a reference is held — which it is, via the `Detection`. It works, but it is an undocumented load-bearing assumption. Worth a comment at minimum. Check whether any code path can violate it.
* **`verbose=True` prints to stdout**, which goes nowhere under `pythonw.exe`; the GUI never sets it. Dead path.
* **`_preview_one` creates a temp dir and an output path that a dry run never writes to.** Harmless, but confusing.
* **`repack_docx` walks with `os.walk` in filesystem order**, so output bytes are not deterministic across runs. `sorted()` would make cleaned files byte-comparable, which is useful for caching and for diffing two runs.
* **`processor.process` catches bare `Exception`** and stringifies it, discarding the traceback. When a pattern change causes a crash, the user gets "Processing error: ..." and no location. Consider an opt-in debug path.
* **Broad config validation gap.** `load_config` validates regex lists only. A colour written `#FF0000` or `red`, or a misspelled style name, silently matches nothing forever. Consider validating `formatting_signals.colors` shape, and — more useful — reporting **per-pattern hit counts** in Preview so dead patterns become visible.
* **CI path filter misses the tests.** `.github/workflows/release.yml` filters on `"*.py"`, which in GitHub path syntax matches root-level files only. A PR touching only `tests/**` runs no CI at all. Also, the suite currently runs *only* inside a Windows packaging job; a 30-second `ubuntu-latest` unittest job on every push would be faster feedback and cheaper minutes.
* **Shared mutable engine.** `DetectionEngine.bind_styles()` mutates engine state per document, and `DocxProcessor` keeps `self._temp_dir`. Both are fine serially and both block any future parallelism. Note it as a constraint, not a defect.

### 4.11 Things I checked and found *sound* — do not spend effort here

So you can allocate attention rather than re-derive:

* **Paragraph-deletion safety rules** (`field_chars_balanced`, `has_section_properties`, `has_embedded_content`, `can_delete_paragraph`) — correct, well-reasoned, well-tested. The four-way table is right.
* **Toggle-property semantics**, including `w:vanish` XOR along `w:basedOn` chains and `w:val="0"` short-circuit. This is a subtle corner of the spec and the implementation reads correctly to me, with tests covering three-deep chains.
* **Text-box / `iter_own_runs` nesting discipline** — the double-counting trap is understood and avoided.
* **Tracked-row deletion handling** — the insight that a deleted row keeps plain `w:t` and records deletion only in `w:trPr` is correct and non-obvious.
* **`os.replace` repack** — no truncated-output window.
* **UTF-8 pinning on `patterns.yaml`** — correct and tested, including a test that demonstrates the cp1252 failure.
* **`apppaths` frozen-build resolution** — correct, and the tests simulate `sys.frozen` / `sys._MEIPASS` properly.
* **GUI threading** — inputs snapshotted on the main thread, log queued, widget updates marshalled through `root.after`. I found no race.
* **Zip Slip** — `ZipFile.extractall` sanitises member paths in CPython, and the threat model (the user's own spec files) does not warrant more. Decompression-bomb limits are arguably missing but I judge them out of scope; disagree if you see it differently.
* **Documentation accuracy.** `README.md`, `CLAUDE.md` and `CHANGELOG.md` are unusually good and I found no place where they describe behaviour the code does not have. If you change behaviour, they must move with it — the maintainer's stated preference is to update the README when implementation changes and `requirements.txt` when dependencies do, and to let documents grow rather than be trimmed.

---

## 5. Exhaustive checklist

Work through this, but do not stop at it. Anything you add is a contribution.

**Correctness — document safety**
1. Can any path produce a `.docx` Word refuses to open? Enumerate the ways and check each is covered.
2. Are the four `_remove_paragraph` exemptions individually necessary and jointly sufficient?
3. Is the redaction → run → paragraph ordering safe under every interleaving? What if a run is both redacted empty and separately marked for removal?
4. Do redaction span offsets stay valid when a paragraph mixes `w:t`, `w:tab`, `w:br`, `w:noBreakHyphen`, and nested `w:hyperlink`/`w:ins`/`w:sdt`?
5. Are `w:sdt` (content controls), `w:smartTag`, `w:fldSimple` and `mc:AlternateContent` handled correctly in *both* processing and extraction?
6. What happens to a run split mid-placeholder across `w:t` boundaries by Word's spell-checker? (There is a test; is it sufficient?)
7. `w:pPr/w:rPr/w:vanish` — a hidden *paragraph mark*. Handled? Should it be?
8. Table structure: merged cells (`w:vMerge`, `w:gridSpan`), nested tables, a cell whose only paragraph is editorial.
9. Section properties: the final `w:sectPr` lives on `w:body`, not a paragraph. Is that path safe?
10. Bookmarks and `REF` fields (§4.7). Comment ranges when `strip_revisions` is off.
11. Numbering (§4.6): `w:numPr`, list restarts, `w:numbering.xml` orphans after removal.
12. Header/footer references: if a header part is emptied of content, does the `w:headerReference` still resolve?
13. `accept_revisions`: nested `w:ins` inside `w:del`, moves (`w:moveFrom`/`w:moveTo`) whose partners live in different parts, deleted paragraph marks.
14. `remove_comment_parts`: is every reference to a removed part cleaned (`document.xml.rels`, `[Content_Types].xml`, sidecar `.rels`, `people.xml`, `commentsExtensible`)?
15. Round-trip a real document through Word after cleaning. This is the only test that actually matters and no automated suite substitutes for it.

**Correctness — detection quality**
16. Enumerate false positives on real specification prose (§4.4 is a start, not the list).
17. Global `re.DOTALL` — audit every pattern against it. Which ones are wrong under it?
18. Unbounded `.*?` in patterns meant to match within a delimiter — which can span a sentence or a paragraph?
19. Is 0.5 the right threshold, and is "italic + colour lands exactly on it" a designed coincidence or an accident? What happens at 0.51 or 0.49?
20. Does the additive-confidence model actually express what it means, or would explicit rules be clearer and safer?
21. Preserve-pattern coverage: what heading forms exist in MasterFormat / SpecLink / firm masters that the five preserve patterns miss?
22. Style-based preserve (`SCT`/`PRT`/`ART`/`EOS`): correct for MasterSpec — what about SpecLink and firm-internal style vocabularies?
23. Inline placeholder patterns: any that could eat real bracketed content? Specs contain legitimate brackets — units, references, options.
24. `spec\s*agent` — could it match legitimate text? ("...the Owner's spec agent shall...")
25. Are detector interactions right? Should a preserve *style* protect against inline redaction as strongly as it does against removal?

**Verification**
26. Is the shared-pattern design (consistency, not independence) the right trade? What would a genuinely independent check look like — a content fingerprint over requirement counts, section headings, numbering sequences, table counts?
27. `MIN_PAIR_SIMILARITY` (§4.3): tune it, replace the heuristic, or use the processor's actual spans?
28. Should `added` paragraphs really be a hard FAIL when the pairing heuristic can manufacture them?
29. Does the FAIL-but-still-write policy serve the user? Should a FAIL rename the output, or write it beside a report?
30. Is `_classify_modification`'s "worst verdict wins" logic actually implemented as documented?
31. Structural lint coverage: which Word-fatal structures are still unchecked?
32. Does verification cover the same parts as processing under every configuration? (`collect_content_parts` is shared — confirm nothing bypasses it.)

**Performance and cost** — see §6
33. Confirm or refute my complexity claims with your own measurements.
34. Where is the boundary between "worth optimising" and "premature"? Real inputs are one section at a time; the pathological case is a consolidated manual. Is that case real enough to design for?
35. Is unpacking each document up to **five** times per file (once to clean, twice for `extract_paragraphs`, twice more for `inspect_structure`) worth fixing, or is disk I/O irrelevant next to the diff cost?
36. Would parallelism across files pay, given the shared-engine constraint (§4.10) and PyInstaller's `multiprocessing` requirements on Windows?

**Architecture and maintainability**
37. Is the module split still right at ~3,300 lines? `docx_xml.py` at 840 lines is the largest — is it one thing or three?
38. Is the strategy-pattern detector hierarchy earning its keep, or would data-driven rules be simpler?
39. Flat layout with no packaging: at what point does that stop scaling, and is this project near it?
40. Error handling: broad `except Exception` in the processor, per-file continuation in the GUI. Right level of granularity?
41. Test coverage gaps: no large-document test, no performance regression test, no real-corpus test, no round-trip-through-Word test.
42. CI (§4.10): path filters, platform, cost.
43. Documentation currency after any change you propose.

**Product**
44. CLI (§4.9).
45. Should the tool emit a plain-text or Markdown sidecar alongside the cleaned `.docx`? See §6.
46. Should Preview report *what was kept and why* as well as what would go? For a tool whose failure mode is over-removal, the kept set is the interesting one.
47. Per-pattern hit counts, so dead and over-eager patterns are visible.
48. Is `patterns.yaml` the right configuration surface for a non-programmer? (In this case the user *is* a programmer, so probably yes.)

---

## 6. "Cheaper to run" — my reading, for you to challenge

The request that prompted this brief asked how the app might be **cheaper to run without sacrificing quality**. That phrase has at least three readings and I do not know which was meant, so treat all three, and say which you think dominates.

### Reading A — downstream LLM token cost (I believe this is the real one)

**SpecCleanse makes no LLM API calls.** It has no model, no key, no network. Its own dollar cost is zero. What it exists to reduce is the cost of the *next* step: feeding specifications to a paid model. So "cheaper to run" almost certainly means *the pipeline this program feeds*, and the levers are:

* **Measure the reduction.** `VerificationResult.removed_characters` already computes characters removed, and the GUI already logs it — but nothing reports it as a *ratio*, and nothing estimates tokens. A one-line "38,400 characters removed, 41% of the document, ≈9,600 tokens saved per analysis pass" makes the value visible and, more importantly, makes a *regression* in cleaning power visible. Cheap, high value.
* **Emit a text sidecar.** Today the output is a `.docx`, which the downstream consumer must convert to text — a conversion that can reintroduce noise the cleaner just removed (repeated headers and footers, field results, table scaffolding). Writing `*_cleaned.txt` or `.md` from the paragraph list the verifier *already builds* would cost almost nothing and would hand the LLM exactly the text SpecCleanse vouched for. **I think this is the highest-value idea in this document**, but I hold it loosely: it widens the program's scope, and scope is not free.
* **Deduplicate repeated content.** A 60-section project manual repeats the same header and footer 60 times. In a per-section workflow that is invisible; in a whole-book ingestion it is pure repeated tokens.
* **Chunk by section.** Emitting one file per `SECTION xx xx xx` boundary lets a retrieval pipeline send the relevant section instead of the book. Large token saving, but it is a real feature with real scope — judge whether it belongs here or in a separate tool.
* **Content-hash caching.** Skip a section whose bytes have not changed since the last clean, and let downstream reuse the prior analysis. Only pays off if the same manual is processed repeatedly, which in spec revision cycles it is.
* **Whitespace and tab collapse in the emitted text.** Small, real, and safe *in a text sidecar* — do not do it in the `.docx`, where whitespace is layout.

Be skeptical of anything in this list that increases the chance of removing a requirement to save tokens. The asymmetry in §1 dominates: an engineering error costs more than every token this program will ever save.

### Reading B — local compute cost

CPU seconds and wall-clock on the user's Windows laptop, plus the human waiting for it. §3 and §4.1–§4.2 are the whole story: two identified hot spots account for nearly all superlinear cost, and both look fixable without touching behaviour. Beyond those:

* Unpack once and share parse trees between processor and verifier — currently up to five full ZIP extractions and re-parses per file.
* Or skip the temp directory entirely and work through `zipfile` streams in memory.
* Parallelise across files (constraint: §4.10; and PyInstaller + `multiprocessing` on Windows needs `freeze_support()`).
* The per-run allocations in §4.10.

Weigh this honestly. If real inputs are 500-paragraph sections cleaned one at a time, the current code is already instant and every one of these optimisations is waste. If the workflow is "clean the whole project manual," §4.1 and §4.2 are the difference between usable and not. **Find out which before recommending work** — and if you cannot find out, say so and scope the recommendation conditionally.

### Reading C — cost of being wrong

The most expensive thing this program can do is delete a requirement that then propagates into a design, a bid, or an installed system. Against that, CPU seconds and tokens are rounding errors. §4.3 (false alarms training the user to ignore the log) and §4.4 (false positives that the verifier reports as PASS) are, on this reading, the two most important findings in this document — not because they are slow, but because they erode the one mechanism that would catch a real loss.

---

## 7. Reproduction scripts

Environment: Python 3.11, `pip install -r requirements.txt`. Run from the repo root.

**Suite:**
```bash
python -m unittest discover -s tests -t .
```

**Scaling benchmark:**
```python
import sys, time; sys.path.insert(0, ".")
from pathlib import Path
from tests import docx_builder as db
from detection import DetectionEngine
from docx_xml import load_config
from processor import DocxProcessor
from verify import verify_clean

OUT = Path("./_bench"); OUT.mkdir(exist_ok=True)

def make(n):
    parts = []
    for i in range(n):
        if i % 12 == 0:
            parts.append(db.para(db.run("Specifier Note: retain or delete paragraph below.",
                                        italic=True, color="FF0000")))
        elif i % 12 == 1:
            parts.append(db.text_para(f"PART {i%3+1} - GENERAL", style="PRT"))
        elif i % 12 == 2:
            parts.append(db.text_para(f"Provide two [Verify quantity with Owner] spare sprinklers for item {i}."))
        else:
            parts.append(db.text_para(f"{i}. Sprinkler piping shall be hydrostatically tested at 200 psi per NFPA 13, section {i}."))
    return db.build_docx(OUT / f"b_{n}.docx", db.document(*parts))

cfg = load_config(Path("patterns.yaml"))
for n in (500, 2000, 6000, 12000):
    src = make(n); eng = DetectionEngine(cfg); dst = OUT / f"b_{n}_c.docx"
    t0 = time.perf_counter(); r = DocxProcessor(eng).process(src, dst); t1 = time.perf_counter()
    assert not r.errors, r.errors
    v = verify_clean(src, dst, engine=eng); t2 = time.perf_counter()
    print(f"paras={n:6d}  clean={t1-t0:6.2f}s  verify={t2-t1:8.2f}s  pass={v.passed}")
```
Profile either phase with `cProfile` sorted by `tottime`.

**False positives on real prose (§4.4):**
```python
import sys; sys.path.insert(0, ".")
from pathlib import Path
from docx_xml import load_config
from detection import DetectionEngine

eng = DetectionEngine(load_config(Path("patterns.yaml")))
for t in [
  "Shop Drawings submitted under this Section may not be reproduced for use on other projects.",
  "Contractor shall verify that duplication of sprinkler coverage in adjacent zones is prohibited by the AHJ.",
  "Unauthorized personnel shall not have access to the fire pump room; reproduction of access keys is not permitted.",
]:
    hits = [(c, p.pattern) for c, p in eng.removal_patterns() if p.search(t)]
    print(("HIT " if hits else "ok  "), t[:70], hits)
```

**Verification false alarm (§4.3):**
```python
# Build a one-paragraph document whose text is mostly placeholder and clean it:
#   db.text_para("Provide [Verify quantity with the Owner and the AHJ prior to bid] units.")
# Then verify_clean(...) and inspect .passed, .unexpected_removals, .added
```

---

## 8. What I did not do

Say so in your report if any of these turn out to matter.

* **I never opened a real specification document.** Every observation about pattern precision comes from synthetic text I wrote myself, informed by domain knowledge but not by the actual corpus. My §4.4 examples are plausible spec prose, not sampled spec prose. If your conclusions depend on how often these patterns fire in practice, that is a question only the maintainer's own files can answer — say so rather than guessing.
* **I never opened an output in Word.** The structural rules read correctly and the tests exercise them, but "Word opens it without complaint" is the real acceptance criterion and I could not run it.
* **I did not review `legacy/`** (1,400 lines, not imported).
* **I did not review the packaging** (`packaging/*.spec`, `*.iss`) or the release workflow beyond the CI path-filter observation.
* **I did not test on Windows**, which is the only platform that matters here — including `tkinter` behaviour, `%APPDATA%` resolution, and path handling.
* **I did not evaluate detection recall at all** — only precision. How much editorial content *survives* a clean is unmeasured and is the other half of the quality question.

---

## 9. Deliverables

1. **A detailed report.** What you verified, what you refuted, what you found that this brief missed, and what remains uncertain. Evidence for each claim — measurements, reproductions, or the specific lines that prove it. Where you disagree with me, say so directly; I would rather be corrected than agreed with.
2. **A proposed implementation plan, if warranted.** Sequenced and scoped, each item stating what it buys, what it costs, what it risks, and how it would be tested. Separate "clearly worth doing" from "defensible" from "considered and rejected" — the rejections are as useful as the acceptances.
3. **An explicit judgement on §6**: which reading of "cheaper to run" actually dominates for this program, and what follows from that.
4. **Your assessment of whether any change is warranted at all.** If the answer is "not much," say that clearly and without padding. This program is small, well-tested, well-documented, and does a job that has a real professional consequence attached to getting it wrong. Stability has value here that a diff does not show.

If you touch the code: the suite must stay green, `README.md` and `CLAUDE.md` must move with any behaviour change, `requirements.txt` must move with any dependency change, and `CHANGELOG.md` follows Keep a Changelog. The maintainer works on Windows and prefers documentation to grow rather than be trimmed.
