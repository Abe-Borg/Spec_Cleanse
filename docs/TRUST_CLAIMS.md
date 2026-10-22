# Trust claims ledger

Every factual claim made by the two trust surfaces — **Why trust it?** (Level 1)
and the **"I'm not convinced"** dossier (Level 2), both written in `trust.py` and
drawn by `trust_view.py` — is listed here with the code that makes it true and
how it is checked.

**The trust surfaces are a contract.** Any change to a behaviour they describe
updates them, and this ledger, in the same change. See *Keeping it true* at the
bottom for what triggers that.

"Verified by" names a test in `tests/test_trust.py` unless it says otherwise.
"manual" means no automated test can reach it, and says what was done instead.

## Inventory (Phase 1)

Taken from the code on the branch that added this file, not from the README.

| | Count | What |
|---|---|---|
| User actions in the window | 15 | Add Files..., Remove Selected, Clear, Output Folder..., Open Output Folder, the *Strip comments and accept tracked changes* checkbox, Preview, CLEAN, About, License, Third-party notices, Close (every dialog), Why trust it?, I'm not convinced — show me exactly what runs →, closing the main window (the only way to stop a run) |
| User actions outside the window | 1 | Editing `patterns.yaml` in a text editor |
| Automatic behaviours | 8 | Unpacking the bundle (frozen build only); finding `patterns.yaml` and seeding the per-user copy on first run (frozen only); the startup log line; the 100 ms log refresh; re-reading and validating `patterns.yaml` at every Preview/CLEAN click; the destination check before CLEAN; verification after every cleaned file; deleting temporary folders |
| Network destinations contacted by SpecCleanse | 0 | — |
| Network traffic a firewall can attribute to *using* it | 2, neither made by SpecCleanse | Windows SmartScreen on first launch of the unsigned `.exe`; ordinary file access when a file you chose lives on a network drive or synced folder |
| AI models or model calls | 0 | — |
| Retries | 0 | Nothing is retried; a failure is reported once |
| Money spent | none | No paid service is called |

Developer tools under `tools/` are not part of the application (`tools/__init__.py`
says so and nothing in the window imports them); they get one runtime card because
the configuration notice tells users about `tools.census_formatting`.

## Claims

### Boundary: network, AI, processes

| ID | Claim | Source | Verified by |
|---|---|---|---|
| C01 | SpecCleanse makes no network connections. No module it ships imports a networking library (`socket`, `ssl`, `http`, `urllib`, `ftplib`, `smtplib`, `asyncio`, `webbrowser`, `requests`, …) | imports of every root module: `gui.py`, `trust.py`, `trust_view.py`, `theme.py`, `about.py`, `apppaths.py`, `batch.py`, `detection.py`, `processor.py`, `verify.py`, `docx_xml.py` | `NoNetworkTests.test_shipped_modules_import_no_networking_module`, `NoNetworkTests.test_every_root_module_is_scanned` |
| C02 | There is no update check, telemetry, crash reporting, licence check or analytics | follows from C01; no such code exists | same as C01 |
| C03 | The only other program SpecCleanse starts is your file browser, from **Open Output Folder** (`os.startfile` on Windows, `open`/`xdg-open` elsewhere) | `gui.SpecCleanseGUI._open_output_dir` | `NoNetworkTests.test_only_open_output_folder_starts_a_process` |
| C04 | The XML parser does not resolve external entities and has network access off | `docx_xml.parse_xml` uses lxml's `XMLParser` defaults: `no_network=True`, `resolve_entities='internal'` | `UntrustedInputTests.test_the_xml_parser_does_not_resolve_external_entities` |
| C05 | No AI, machine-learning model or remote service is involved anywhere | C01; every decision is in `detection.py` regular expressions and formatting checks | C01 tests; `DossierStructureTests.test_every_card_says_ai_none` |
| C06 | Links in About and in the dossier are text to copy; the app opens no browser | `about.about_text`, `trust_view` (no `webbrowser`) | C01 (`webbrowser` is on the forbidden list); `NoExternalAssetTests.test_trust_files_reference_only_the_further_reading_urls` |
| C07 | Windows SmartScreen may check the unsigned `.exe` on first launch; that is Windows, not SpecCleanse | `packaging/speccleanse.spec` (`codesign_identity=None`); README *Installation* | manual — Windows behaviour, outside the code |
| C08 | Files on a network drive or synced folder move over the network as ordinary file access by Windows | consequence of reading/writing the paths you choose (`processor.DocxProcessor.process`, `processor.repack_docx`) | manual — operating-system behaviour |
| C09 | The trust dialogs load no external asset: no images, fonts or scripts from anywhere; the diagram is drawn on a Tk canvas | `trust_view.py` | `NoExternalAssetTests.test_trust_files_reference_only_the_further_reading_urls`, `NoExternalAssetTests.test_trust_view_loads_no_image_or_file` |

### What happens to your files

| ID | Claim | Source | Verified by |
|---|---|---|---|
| C10 | The original is opened read-only and never written to, by CLEAN or by its verification | `processor.DocxProcessor._unpack_docx`, `_is_valid_docx`, `verify._unpack` (all `ZipFile(path, "r")`) | `OriginalUntouchedTests.test_clean_and_verify_leave_the_input_byte_identical` |
| C11 | The cleaned file is named `<name>_cleaned.docx`, in the chosen output folder or beside the original | `batch.OUTPUT_SUFFIX`, `batch.output_for` | `NumbersTests.test_output_suffix` (copy interpolates the constant) |
| C12 | The cleaned file is built as `<name>_cleaned.docx.tmp` and moved into place only when complete | `processor.repack_docx` | `OriginalUntouchedTests.test_no_tmp_file_is_left_after_a_clean`; existing `tests/test_processor.py` |
| C13 | A batch whose destinations collide, or whose destination is another selected input, is refused with nothing written | `batch.plan_batch`, `gui.SpecCleanseGUI._reject_batch` | existing `tests/test_batch.py` |
| C14 | An existing cleaned file is replaced only after you confirm | `gui.SpecCleanseGUI._confirm_overwrite` | existing behaviour; manual (message box) |
| C15 | Preview writes nothing to your folders | `gui._preview_one` (`dry_run=True`), `processor.DocxProcessor.process` | `GuiWorkerTrustTests.test_preview_writes_nothing_beside_the_input` |
| C16 | Parts of the package the clean does not change are copied byte for byte (properties, images, styles, settings) | `processor.DocxProcessor._process_xml_file` writes a part only when something changed; `repack_docx` copies every file | `OriginalUntouchedTests.test_unchanged_parts_are_copied_byte_for_byte` |
| C17 | Temporary working folders go in the system temp folder with prefixes `speccleanse_` (clean and preview), `speccleanse_verify_` (verification) and `speccleanse_preview_` (preview), and are deleted when each file is done | `processor.DocxProcessor.process`, `verify._unpack`, `gui._preview_one` | `TempFolderTests.test_core_temp_prefixes_match_the_copy`, `TempFolderTests.test_temp_folders_are_removed_after_a_clean`, `GuiWorkerTrustTests.test_preview_uses_the_documented_temp_prefix` |
| C18 | The log lives only in the window: it is not saved, and each Preview or CLEAN clears it. It refreshes ten times a second | `gui.SpecCleanseGUI._log`, `_clear_log`, `_drain_log` (`after(100, …)`); no file write in `gui.py` | `NoNetworkTests.test_gui_writes_no_file_of_its_own`, `NumbersTests.test_log_refresh_interval` |
| C19 | You can select and copy text from the log | `tk.Text` in state `disabled` still takes a mouse selection and `<<Copy>>` | manual — checked with Tk 8.6.14 under Xvfb |
| C20 | Nothing is remembered between sessions: no settings file, no recent-files list, no history. The checkbox and the output folder start fresh every time | `gui.SpecCleanseGUI.__init__` (`strip_revisions` defaults to `False`, `output_dir` to `None`) | `DialogTests.test_the_window_starts_with_defaults`; C18 scan |
| C21 | `patterns.yaml` is read from beside the modules (source), from beside the `.exe` if one is there, or from `%APPDATA%\SpecCleanse\patterns.yaml`, copied from the shipped default on first run | `apppaths.resolve_config_path`, `apppaths.CONFIG_NAME`, `apppaths.APP_NAME` | existing `tests/test_apppaths.py`; `NumbersTests.test_config_locations` |
| C22 | Uninstalling leaves your `%APPDATA%\SpecCleanse\patterns.yaml` in place | `packaging/installer.iss` `[UninstallDelete]` (empty, with comment) | `NumbersTests.test_uninstall_keeps_patterns` |
| C23 | The installer is per-user by default and needs no administrator rights | `packaging/installer.iss` `PrivilegesRequired=lowest` | `NumbersTests.test_installer_is_per_user_by_default` |
| C24 | The frozen executable unpacks itself to a temporary folder for the session and removes it on exit | PyInstaller one-file mode, `packaging/speccleanse.spec` (`runtime_tmpdir=None`) | manual — PyInstaller behaviour |
| C25 | `patterns.yaml` is re-read and re-validated at every Preview or CLEAN click, so an edit takes effect on the next click | `gui.build_engine`, `gui.SpecCleanseGUI._load_engine` | `GuiWorkerTrustTests.test_build_engine_rereads_patterns_each_time` |
| C26 | The first log line names the `patterns.yaml` in use, and every run repeats it | `gui.SpecCleanseGUI._build_ui`, `_load_engine` | manual; existing behaviour |

### What the cleaner decides, and how

| ID | Claim | Source | Verified by |
|---|---|---|---|
| C30 | Decisions come from regular expressions, Word formatting (hidden, italic, colour) and style names in `patterns.yaml`; nothing else | `detection.DetectionEngine`, detector classes | `ShippedRulesTests` |
| C31 | Patterns are matched case-insensitively, and `^` anchors at the start of a paragraph | `docx_xml.compile_patterns` (`re.IGNORECASE \| re.DOTALL`) | `ConfigSafetyTests.test_patterns_are_case_insensitive_and_dotall` |
| C32 | The same document and the same rules give the same decisions | detection is a pure function of the XML and config | `DeterminismTests.test_two_runs_make_the_same_decisions` |
| C33 | Removal threshold is 0.5 | `detection.DetectionEngine.should_remove` | `NumbersTests.test_removal_threshold` |
| C34 | Italic counts 0.2 and an editorial colour 0.3; together exactly 0.5 | `detection.SpecifierNoteDetector.detect` | `NumbersTests.test_formatting_alone_scores_exactly_the_threshold` |
| C35 | Removing on formatting alone is off by default (`specifier_notes.formatting_only_removal: false`) | `patterns.yaml`; `detection.PatternConfig` | `NumbersTests.test_formatting_only_removal_is_off_by_default`; existing `tests/test_config.py` |
| C36 | The editorial colours are FF0000, C00000, 0000FF and 0070C0 by default | `patterns.yaml` `specifier_notes.formatting_signals.colors` | `NumbersTests.test_editorial_colours` |
| C37 | Low-confidence editorial patterns score 0.3 and need a formatting signal to be removed | `detection.EditorialArtifactDetector.detect` | `NumbersTests.test_low_confidence_needs_formatting` |
| C38 | A preserve rule (pattern or style) protects a paragraph outright | `detection.DetectionEngine.detect_in_element`, `should_remove` | `NumbersTests.test_preserve_wins` |
| C39 | Shipped rule counts: specifier notes 8 patterns; copyright 6; SpecAgent 1; editorial 13 high-confidence, 12 low-confidence, 16 inline placeholders; preserve 5 patterns and 4 styles; editorial styles 6 paragraph and 3 character | `patterns.yaml` | `ShippedRulesTests.test_counts_match_patterns_yaml` |
| C40 | Inline placeholders are cut out of the sentence; the paragraph goes only if nothing else is left | `processor.DocxProcessor._redaction_spans` | existing `tests/test_inline.py` |
| C41 | A placeholder straddling a page or column break is left in place, with a note in the log | `processor.DocxProcessor._drop_spans_over_layout_breaks` | existing `tests/test_inline.py` |
| C42 | A paragraph is emptied rather than deleted when it carries a section break, a picture/field/note reference, an unbalanced field, or is the last block its container may hold | `processor.DocxProcessor._remove_paragraph` | existing `tests/test_processor.py` |
| C43 | The cleaner adds no words: it deletes paragraphs, runs and character spans | `processor.DocxProcessor` (no code path writes new text; `set_text` only shortens) | `OnlyRemovesTests.test_every_cleaned_paragraph_comes_from_the_original`, `OnlyRemovesTests.test_verification_fails_an_added_paragraph` |
| C44 | Tracked changes and comments are untouched unless the checkbox is ticked; ticked, insertions are kept, deletions removed, comments and comment parts deleted | `processor.DocxProcessor.strip_revisions`, `docx_xml.accept_revisions`, `docx_xml.remove_comment_parts`, `docx_xml.COMMENT_PARTS` | existing `tests/test_revisions.py`; `NumbersTests.test_comment_parts` |
| C45 | A deleted paragraph mark is not acted on (paragraphs are not merged) | `docx_xml.accept_revisions` | existing `tests/test_revisions.py` |
| C46 | Parts scanned: the document body, headers, footers, footnotes, endnotes and the glossary | `docx_xml.collect_content_parts` | `NumbersTests.test_content_parts` |
| C47 | Preview's log shortens each paragraph to 90 characters | `gui._shorten` | `GuiWorkerTrustTests.test_log_previews_are_90_characters` |

### Verification

| ID | Claim | Source | Verified by |
|---|---|---|---|
| C50 | Every CLEAN is followed automatically by verification of that file | `gui._clean_one` | existing `tests/test_gui.py` `test_clean_reports_a_pass` |
| C51 | Verification judges the output against what the rules permit losing from the original; it never reads the cleaner's report | `verify.verify_clean` (evidence from `extract_paragraphs(input_path, engine)`), `detection.DetectionEngine.paragraph_evidence` | existing `tests/test_verify.py` `InjectedDamageTests` |
| C52 | Verification uses the same rules as the cleaner, so a rule that removes the wrong thing is reported as expected | `gui._clean_one` passes `engine` to `verify_clean` | manual — stated limit; README *Verification* |
| C53 | Outcomes: Verified, Needs review (with one or more of four reasons), Failed | `batch.FileOutcome`, `batch.ReviewCategory` | `NumbersTests.test_outcome_and_category_names` |
| C54 | A file that needs review is still written, and its path is named | `gui._clean_one` | existing `tests/test_gui.py` |
| C55 | Pairing falls back to a similarity ratio of 0.5 only when no exact answer exists, and a finding reached that way is reported as ambiguous alignment | `verify.MIN_PAIR_SIMILARITY`, `verify._pair_with_survivor` | `NumbersTests.test_pairing_similarity` |
| C56 | A configuration notice (formatting-only removal on, or a superseded rule still active) makes every file in the run need review | `detection.config_notices`, `gui.SpecCleanseGUI._run_clean` | existing `tests/test_gui.py` |
| C57 | One file's failure does not stop the batch | `batch.run_batch` | existing `tests/test_batch.py` |
| C58 | Broken cross-references and numbering a removal may shift are reported; nothing is repaired or renumbered | `verify._compare_structure`, `verify.NumberingNotice` | existing `tests/test_references.py`, `tests/test_numbering.py` |

### Untrusted input and security

| ID | Claim | Source | Verified by |
|---|---|---|---|
| C60 | A ZIP entry cannot be unpacked outside the temporary folder | Python `zipfile.ZipFile.extractall` strips drive letters, absolute paths and `..` | `UntrustedInputTests.test_a_zip_entry_cannot_escape_the_temp_folder` |
| C61 | `patterns.yaml` cannot run code: it is loaded with `yaml.safe_load` | `docx_xml.load_config` | `ConfigSafetyTests.test_patterns_yaml_cannot_construct_python_objects` |
| C62 | Nothing in the document is executed; macros and embedded objects are carried through untouched | no execution path in `processor.py` | C01/C03 tests |
| C63 | No credentials, keys or accounts exist | no such code | C01 tests |
| C64 | No listening port or local server | C01 (`socket` forbidden) | C01 tests |
| C65 | The executable is not code-signed | `packaging/speccleanse.spec` `codesign_identity=None` | `NumbersTests.test_executable_is_unsigned` |
| C66 | Runtime libraries are pinned: lxml 6.0.2, PyYAML 6.0.3 | `requirements.txt` | `NumbersTests.test_pinned_versions` |

### Limits (the Honesty section)

| ID | Claim | Source | Verified by |
|---|---|---|---|
| C70 | There is no Stop button; closing the window is the only way to stop a run | `gui.SpecCleanseGUI` (no cancel control; worker is a daemon thread) | `DialogTests.test_there_is_no_stop_button` |
| C71 | Closing mid-run can leave a `.tmp` file beside the destination and a `speccleanse_` folder holding an unpacked copy of the document in the temp folder | daemon thread killed at exit; `finally` blocks do not run | manual — reasoning from `threading` semantics; reported as a finding |
| C77 | If deleting a temporary folder fails, the cleaner reports an error, but verification leaves its copies silently | `processor.DocxProcessor.process` (`shutil.rmtree` raises into `errors`), `verify.extract_paragraphs` and `verify.inspect_structure` (`shutil.rmtree(..., ignore_errors=True)`) | manual — reported as a finding |
| C78 | Documents are processed only after a Preview or CLEAN click, on a background thread so the window stays responsive | `gui.SpecCleanseGUI._start_preview`, `_start_clean` (`threading.Thread`) | `NoNetworkTests.test_processing_starts_only_from_a_click` |
| C72 | No size or time limit: a very large or crafted file is unpacked in full, and a badly written pattern can run for a very long time | no cap in `processor.DocxProcessor._unpack_docx`, no timeout anywhere | manual — reported as a finding |
| C73 | Only `.docx`; the file check only asks whether `word/document.xml` exists | `processor.DocxProcessor._is_valid_docx` | existing `tests/test_processor.py` |
| C74 | Document properties (author, company, custom properties) are not scrubbed | C16 | `OriginalUntouchedTests.test_unchanged_parts_are_copied_byte_for_byte` |
| C75 | No real specification corpus has been measured, and no cleaned file has been opened in Word by any automated check | README *What has and has not been validated* | manual — project statement |
| C76 | Tk exposes little to screen readers | Tkinter on Windows | manual |

### The trust dialogs themselves

| ID | Claim | Source | Verified by |
|---|---|---|---|
| C80 | **Why trust it?** sits beside About and inside the About dialog, and stays usable during a run | `gui.SpecCleanseGUI._build_ui` (`btn_trust`, not in `_run_controls`), `_show_about` | `DialogTests.test_the_trust_buttons_stay_usable_during_a_run`, `DialogTests.test_about_offers_why_trust_it` |
| C81 | Its button opens the dossier on top of it, not instead of it | `trust_view.WhyTrustDialog.open_dossier` | `DialogTests.test_level1_button_opens_the_dossier` |
| C82 | One Escape closes exactly one dialog | `trust_view._StackedDialog` (`<Escape>` bound on each toplevel) | `DialogTests.test_one_escape_closes_one_dialog` |
| C83 | Focus moves into a dialog when it opens, the top dialog holds the input, and focus returns to the opener on close | `trust_view._StackedDialog.take_input`, `close` | `DialogTests.test_focus_moves_in_and_returns_to_the_opener`, `DialogTests.test_the_top_dialog_holds_the_input` |
| C84 | The dossier is at most 88% of the screen height | `trust_view.DOSSIER_MAX_HEIGHT` | `DialogTests.test_dossier_is_at_most_88_percent_of_the_screen_height` |
| C85 | Every section has an anchor; the contents rail navigates and follows the scroll, and hides below 900 px wide | `trust_view.DossierDialog.go_to`, `_on_scroll`, `RAIL_MIN_WIDTH` | `DialogTests.test_every_section_has_an_anchor_and_a_rail_entry`, `DialogTests.test_the_rail_follows_the_scroll` |
| C86 | Every runtime card shows the same five rows | `trust.CARD_ROWS`, `trust_view.DossierDialog._card` | `DossierStructureTests.test_every_card_has_the_five_rows`, `DialogTests.test_every_card_renders_all_five_rows` |

The dialog tests need a display and skip without one. CI's Windows lane has one;
its Linux lane does not, so there they report as skipped.

## Numbers quoted in the copy

`tests/test_trust.py` (`DossierStructureTests.test_every_number_in_the_copy_is_accounted_for`)
fails if the copy contains a number not on this list, and each entry is pinned by
the test named beside it.

| Number | Meaning | Pinned by |
|---|---|---|
| 0.5 | removal threshold; also the pairing-similarity fallback | `test_removal_threshold`, `test_pairing_similarity` |
| 0.2 | italic signal | `test_formatting_alone_scores_exactly_the_threshold` |
| 0.3 | editorial colour signal; low-confidence pattern score | `test_formatting_alone_scores_exactly_the_threshold`, `test_low_confidence_needs_formatting` |
| 90 | log preview width, characters | `test_log_previews_are_90_characters` |
| FF0000, C00000, 0000FF, 0070C0 | default editorial colours | `test_editorial_colours` |
| 8, 6, 1, 13, 12, 16, 5, 4, 3 | shipped rule counts (C39) | `test_counts_match_patterns_yaml` |
| 6.0.2, 6.0.3 | lxml, PyYAML versions | `test_pinned_versions` |
| two (in words) | the exceptions a firewall may show | `test_every_network_exception_is_in_the_boundary` |
| ten (in words) | log refreshes per second | `test_log_refresh_interval` |

## Keeping it true

Update `trust.py`, this ledger and `tests/test_trust.py` in the same change as any of:

- a new user action — a button, menu item, shortcut, file drop, dialog;
- a new network call of any kind, including an update check or a link the app opens;
- a new automatic behaviour — a timer, background job, retry, automatic follow-up;
- a model or AI service, or a change to one;
- a new or changed limit, threshold, default, file location or temporary file;
- a change to what is written, kept, or deleted on disk.

`DossierStructureTests.test_every_button_has_a_runtime_card` fails when a button is
added to `gui.py` or `trust_view.py` without a card that names it.
