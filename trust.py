#!/usr/bin/env python3
"""
What the two trust surfaces say: "Why trust it?" and the dossier behind it.

Purpose. A specifier, engineer or reviewer who puts their name on a cleaned
specification needs to know what this program does to it before relying on it.
Level 1 (``LEVEL1_*``) is the short answer for a reader in a hurry. Level 2
(``DOSSIER``) is the long one for the reader who has to sign: the mechanism
behind every claim, action by action, and how to check all of it.

Reader. A working engineer or specifier, not a programmer. Plain words, "you",
and a technical term defined the first time it appears.

The rule. If the implementation changes, this changes with it. A trust document
that has drifted from the code is worse than none. Every claim here is listed in
``docs/TRUST_CLAIMS.md`` with the code that makes it true, and
``tests/test_trust.py`` pins the numbers, names and locations to their sources —
several of them are read straight from the constants below rather than retyped.

Kept apart from ``trust_view.py``, which draws it, so the words can be tested
where tkinter is absent.
"""

from __future__ import annotations

from dataclasses import dataclass

import about
from apppaths import APP_NAME, CONFIG_NAME
from batch import OUTPUT_SUFFIX, FileOutcome, ReviewCategory

# ---------------------------------------------------------------------------
# Building blocks.  The view draws each kind the same way wherever it appears,
# so every section of the dossier looks alike.  ``**bold**`` is the only markup.
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Point:
    """A short title stating a fact, and the mechanism that makes it true."""
    title: str
    body: str


@dataclass(frozen=True)
class Para:
    text: str


@dataclass(frozen=True)
class Bullets:
    items: tuple[str, ...]


@dataclass(frozen=True)
class Table:
    """Rows drawn as records: the first column as a heading, the rest labelled."""
    columns: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]


@dataclass(frozen=True)
class Contrast:
    """Two lists side by side."""
    left_title: str
    left: tuple[str, ...]
    right_title: str
    right: tuple[str, ...]


@dataclass(frozen=True)
class Note:
    title: str
    text: str


@dataclass(frozen=True)
class Term:
    """A loaded word: what it claims exactly, and what it does not."""
    word: str
    claim: str
    not_claimed: str


@dataclass(frozen=True)
class RuntimeCard:
    """One action, always answered in the same five rows, in this order."""
    title: str
    you_do: str
    what_runs: str
    what_is_sent: str
    ai_involved: str
    bounded_by: str


#: The five row labels, in order.  The view and the tests both read this.
CARD_ROWS = (
    ("you_do", "You do"),
    ("what_runs", "What runs"),
    ("what_is_sent", "What is sent"),
    ("ai_involved", "AI involved"),
    ("bounded_by", "Bounded by"),
)


@dataclass(frozen=True)
class Diagram:
    """The data-flow drawing.  ``description`` restates it in words."""
    description: str


@dataclass(frozen=True)
class Link:
    label: str
    url: str


@dataclass(frozen=True)
class LinkBox:
    title: str
    links: tuple[Link, ...]


@dataclass(frozen=True)
class Section:
    id: str
    kicker: str
    title: str
    blocks: tuple


# ---------------------------------------------------------------------------
# Facts quoted in the copy, gathered so the tests can hold them to the code.
# ---------------------------------------------------------------------------

#: Where a cleaned file goes, as the copy states it.
OUTPUT_NAME = f"<name>{OUTPUT_SUFFIX}.docx"
TEMP_NAME = f"{OUTPUT_NAME}.tmp"
USER_CONFIG = rf"%APPDATA%\{APP_NAME}\{CONFIG_NAME}"

#: Temporary-folder prefixes, as the copy states them.
TEMP_PREFIXES = ("speccleanse_", "speccleanse_verify_", "speccleanse_preview_")

#: Network traffic a firewall may show while you use SpecCleanse.  None of it is
#: made by SpecCleanse; every entry must appear in the boundary section.
NETWORK_EXCEPTIONS = (
    "Windows SmartScreen",
    "network drive or synced folder",
)

#: The outcomes and review reasons the copy names, read from the code.
OUTCOME_NAMES = {outcome: outcome.value.capitalize() for outcome in FileOutcome}
CATEGORY_NAMES = {category: category.value for category in ReviewCategory}

#: What the shipped patterns.yaml holds, as the copy states it.  A test holds
#: every count to the file.
SHIPPED_RULE_COUNTS = {
    ("specifier_notes", "text_patterns"): 8,
    ("copyright_notices", "text_patterns"): 6,
    ("specagent_references", "text_patterns"): 1,
    ("editorial_artifacts", "text_patterns"): 13,
    ("editorial_artifacts", "low_confidence_patterns"): 12,
    ("editorial_artifacts", "inline_patterns"): 16,
    ("preserve_patterns", "text_patterns"): 5,
    ("style_based_detection", "preserve_styles"): 4,
    ("style_based_detection", "paragraph_styles"): 6,
    ("style_based_detection", "character_styles"): 3,
}
#: The shipped editorial colours, with the names the copy gives them.
EDITORIAL_COLOURS = (
    ("red", "FF0000"),
    ("dark red", "C00000"),
    ("blue", "0000FF"),
    ("light blue", "0070C0"),
)

_NOT_CONVINCED = "I'm not convinced — show me exactly what runs →"

# ---------------------------------------------------------------------------
# Level 1 — "Why trust it?"
# ---------------------------------------------------------------------------

LEVEL1_TITLE = "Why trust it?"

LEVEL1_LEAD = (
    "SpecCleanse assumes a qualified person reads what it produces, so it only "
    "ever removes, it runs entirely on your computer, and it tells you about "
    "anything it cannot account for instead of guessing."
)

LEVEL1_POINTS = (
    Point(
        "Everything runs on your computer",
        "The program contains no networking code at all: no update check, no "
        "usage reporting, no licence check. Two things can still show up in a "
        "firewall log, and neither is SpecCleanse: Windows SmartScreen checking "
        "the unsigned program the first time you run it, and Windows reading "
        "or writing your files when you keep them on a network drive or synced "
        "folder.",
    ),
    Point(
        "No AI is involved anywhere",
        f"What gets removed is decided by text-matching rules, Word formatting "
        f"and style names listed in {CONFIG_NAME} — a plain text file you can "
        f"open and read. The same document and the same rules always give the "
        f"same decisions.",
    ),
    Point(
        "It only removes, never adds",
        "The cleaner deletes paragraphs, pieces of paragraphs and placeholders "
        "like [Verify quantity]. It has no way to write new words, and the "
        "check after every clean fails any file containing a paragraph the "
        "original did not have.",
    ),
    Point(
        "Your original is never changed",
        f"The original is opened read-only. The result is a new file, "
        f"{OUTPUT_NAME}, built under a temporary name and moved into place only "
        f"when complete. A batch that would overwrite one of its own files is "
        f"refused, and an earlier cleaned file is replaced only after you say so.",
    ),
    Point(
        "Every clean is checked independently",
        "After writing each file, SpecCleanse compares it with the original and "
        "asks, for every lost paragraph, whether the rules allowed that loss — "
        "working from the original, never from the cleaner's account of what it "
        "did. Anything unexplained marks the file Needs review and says why.",
    ),
    Point(
        "Nothing starts without your click",
        f"Documents are processed only after you click Preview or CLEAN — on a "
        f"background thread, so the window stays responsive. There are no "
        f"scheduled jobs, retries or updates. Only two things happen without a "
        f"click of their own: at start-up the program finds "
        f"{CONFIG_NAME} (the installed program copies a default to "
        f"{USER_CONFIG} the first time), and after each CLEAN the check runs "
        f"automatically.",
    ),
    Point(
        "Your documents are not kept",
        "Each file is unpacked into your temporary folder while it is worked on "
        "and deleted when it is done. A copy can be left behind if the program "
        "is closed or crashes part-way through a file, or if Windows refuses "
        "the deletion — and for the check's own copies that happens without a "
        "word. The log exists only in the window. There "
        "is no history, no saved settings and nothing to sign in to.",
    ),
    Point(
        "A pass is not a promise",
        "The check uses the same rules as the cleaner, so a rule that removes "
        "the wrong thing is reported as expected. No cleaned file has yet been "
        "opened in Word by any automated test. Read the log, and compare the "
        "result against the original before you rely on it.",
    ),
)

LEVEL1_CLOSING_TITLE = "Not convinced?"
LEVEL1_CLOSING_TEXT = (
    "Fair. The points above are claims — here is the mechanism behind each one, "
    "action by action, plus how to audit any of it yourself."
)
LEVEL1_BUTTON = _NOT_CONVINCED

# ---------------------------------------------------------------------------
# Level 2 — the dossier
# ---------------------------------------------------------------------------

DOSSIER_TITLE = "Exactly what runs"
DOSSIER_SUBTITLE = (
    f"Everything {APP_NAME} does to your documents, action by action, and how to "
    "check it. Written from the code as it ships; where a setting changes "
    "something, it says so and names the setting."
)

_START = Section(
    "start", "Start here", "The short answer",
    (
        Point(
            "Every word in the cleaned file has one origin: your original.",
            "SpecCleanse removes text and never writes any. What it removes is "
            f"decided by the rules in {CONFIG_NAME}, and the log says which kind "
            "of rule removed what.",
        ),
        Point(
            "Nothing runs that you did not start — with every exception named.",
            "No network, no scheduled or unprompted work, no AI. The only things "
            "that happen "
            "without a click are listed in the Runtime section below, and each "
            "one stays on your computer.",
        ),
        Point(
            "Every automated judgment can be re-derived.",
            "The rules are a text file, the decisions are the same every time, "
            "and the check after each clean recomputes from the original what "
            "should have gone. You can repeat any of it with Preview, or with "
            "Word's own Compare.",
        ),
    ),
)

_PROVENANCE = Section(
    "provenance", "Provenance", "Where the cleaned file comes from",
    (
        Para(
            "A .docx file is a ZIP archive of XML files — text files that "
            "describe the document. SpecCleanse unpacks it, edits the XML that "
            "holds the text, and packs it again."
        ),
        Table(
            ("Origin", "What it actually is", "How you can tell"),
            (
                (
                    "Your original document",
                    "Every word in the cleaned file. Parts the clean did not "
                    "change — document properties, pictures, styles, settings — "
                    "are copied across byte for byte.",
                    "The check reports any paragraph in the output that was not "
                    "in the input as a failure. Word's Review → Compare should "
                    "show deletions only.",
                ),
                (
                    f"The rules in {CONFIG_NAME}",
                    "Decide what is removed: text patterns (regular expressions "
                    "— a precise way of describing a piece of text, such as "
                    "\"[Verify\" followed by anything up to \"]\"), Word "
                    "formatting, and style names.",
                    f"The first line of every run's log names the {CONFIG_NAME} "
                    "in use. Preview lists every removal under the kind of rule "
                    "that caused it, with its score.",
                ),
                (
                    "Your own edits to the rules",
                    f"Exactly the same as above: {CONFIG_NAME} is the only "
                    "place rules come from. An edit takes effect at the next "
                    "Preview or CLEAN click.",
                    "A configuration notice in the log, and Needs review on "
                    "every file, when a setting that can remove text on looks "
                    "alone is switched on.",
                ),
                (
                    "Accepted tracked changes (only if you tick the box)",
                    "Insertions already in your document become ordinary text; "
                    "tracked deletions and comments are removed.",
                    "Off by default. The check counts text lost this way as "
                    "expected only when the box was ticked.",
                ),
                (
                    f"{APP_NAME} itself",
                    "No words. It re-saves each XML part it changed, which can "
                    "change the file's layout but not its content; moves "
                    "bookmark markers out of a deleted paragraph "
                    "so the bookmark still spans the same text, and removes the "
                    "bookkeeping entries of deleted comment files.",
                    "Unpack both files (rename a copy to .zip) and compare the "
                    "XML.",
                ),
                (
                    "An AI model",
                    "Nothing. None is used.",
                    "There is no code that could call one; see The boundary.",
                ),
            ),
        ),
        Note(
            "Things that are not happening",
            "No rules are downloaded — the only rules are the ones in your "
            f"{CONFIG_NAME}. Nothing learns from your documents: no run "
            "remembers anything from the last one. Your documents are not "
            "uploaded, shared with other users, or used to train anything, "
            "because the program has no means of sending them anywhere. There "
            "is no hidden list of phrases outside that file.",
        ),
    ),
)

_ENGINE = Section(
    "engine", "Components", "What the engine actually is",
    (
        Para(
            "There is no model and no remote service. Every job is done by "
            "ordinary, inspectable code, all of it running on your computer."
        ),
        Table(
            ("Job", "Component", "Why"),
            (
                ("Opening and writing the .docx",
                 "Python's zipfile module, on your computer",
                 "A .docx is a ZIP archive. The original is opened read-only."),
                ("Reading and editing the XML",
                 "lxml 6.0.2, on your computer",
                 "Editing the XML directly keeps every piece of formatting the "
                 "clean does not touch exactly as it was. Network access in the "
                 "parser is off."),
                ("Deciding what to remove",
                 "Regular expressions and formatting checks in detection.py, "
                 f"from {CONFIG_NAME}, on your computer",
                 "Matched without regard to capital letters; a rule's \"^\" "
                 "means the start of a paragraph. Each finding gets a score and "
                 "is removed at 0.5 or above."),
                ("Reading the rules",
                 "PyYAML 6.0.3, safe mode, on your computer",
                 "Safe mode reads plain values only; a rules file cannot run "
                 "code."),
                ("Checking the result",
                 "Python's difflib plus SpecCleanse's own comparison, in "
                 "verify.py, on your computer",
                 "Works out what the rules permitted from the original, then "
                 "compares the output with it."),
                ("The window",
                 "Tkinter, Python's built-in toolkit, on your computer",
                 "Draws the window, including this one. It loads nothing from "
                 "the internet."),
            ),
        ),
        Para(
            "The shipped rules file holds, by default: "
            f"{SHIPPED_RULE_COUNTS['specifier_notes', 'text_patterns']} "
            "specifier-note patterns, "
            f"{SHIPPED_RULE_COUNTS['copyright_notices', 'text_patterns']} "
            "copyright patterns, "
            f"{SHIPPED_RULE_COUNTS['specagent_references', 'text_patterns']} "
            "SpecAgent pattern, "
            f"{SHIPPED_RULE_COUNTS['editorial_artifacts', 'text_patterns']} "
            "high-confidence and "
            f"{SHIPPED_RULE_COUNTS['editorial_artifacts', 'low_confidence_patterns']} "
            "low-confidence editorial patterns, "
            f"{SHIPPED_RULE_COUNTS['editorial_artifacts', 'inline_patterns']} "
            "placeholder patterns, "
            f"{SHIPPED_RULE_COUNTS['preserve_patterns', 'text_patterns']} "
            "protection patterns and "
            f"{SHIPPED_RULE_COUNTS['style_based_detection', 'preserve_styles']} "
            "protected styles, and "
            f"{SHIPPED_RULE_COUNTS['style_based_detection', 'paragraph_styles']} "
            "paragraph and "
            f"{SHIPPED_RULE_COUNTS['style_based_detection', 'character_styles']} "
            "character styles treated as editorial notes. Hidden text needs no "
            "pattern: Word's own hidden setting is read directly."
        ),
        Para(
            "Scores in plain words: a specifier-note pattern on its own is "
            "enough to remove; italic adds 0.2 and an editorial colour ("
            + ", ".join(f"{name} {code}" for name, code in EDITORIAL_COLOURS)
            + ") adds 0.3 "
            "— together exactly 0.5 — but by default looks alone are never "
            "enough, because real specification text is often red and italic. "
            "Editorial patterns marked low-confidence score 0.3 and are removed "
            "only with a formatting signal behind them."
        ),
    ),
)

_BOUNDARY = Section(
    "boundary", "The boundary", "What leaves your computer, and where it goes",
    (
        Diagram(
            "Data-flow diagram. Inside your computer: your original .docx and "
            f"{CONFIG_NAME} flow into SpecCleanse, which works in your "
            f"temporary folder and writes {OUTPUT_NAME}. There is no line from "
            "SpecCleanse to the internet. Two dashed lines, neither drawn from "
            "SpecCleanse, cross the edge: Windows SmartScreen, which Windows "
            "itself may contact the first time you run the unsigned program; "
            "and a network drive or synced folder, which Windows reads and "
            "writes only if your files live there."
        ),
        Contrast(
            "Leaves your machine",
            (
                "Nothing, from SpecCleanse. It contains no code that opens a "
                "network connection.",
                "Windows SmartScreen (Windows, not SpecCleanse): may check the "
                "unsigned program's reputation with Microsoft the first time you "
                "run it.",
                "Network drive or synced folder (Windows, not SpecCleanse): if "
                "an original, the output folder or your rules file lives on one, "
                "Windows moves those bytes over the network as it would for any "
                "program.",
            ),
            "Stays on your machine",
            (
                "Your original .docx — read, never written.",
                f"{OUTPUT_NAME}, in the folder you chose or beside the original.",
                f"{CONFIG_NAME}, your rules.",
                "Temporary working copies, deleted when each file is done.",
                "The log — in the window only, never saved.",
            ),
        ),
        Note(
            "Before you install",
            "Downloading the program from the releases page, or installing "
            "lxml and PyYAML with pip to run from source, uses the internet — "
            "once, at your hands, before SpecCleanse ever runs.",
        ),
    ),
)

_NONE = "None."

_RUNTIME = Section(
    "runtime", "Runtime", "What runs when you click",
    (
        Para(
            "One card per action, every card answering the same five questions "
            "in the same order. A paragraph is Word's unit of text ending in a "
            "return; a run is a stretch of text inside it with one set of "
            "formatting."
        ),
        RuntimeCard(
            "Starting SpecCleanse (automatic)",
            "Open the program.",
            "The installed or portable program first unpacks itself into a "
            "temporary folder for the session. Then it works out which "
            f"{CONFIG_NAME} to use — beside the program if you put one there, "
            f"otherwise {USER_CONFIG}, copied from the shipped default the "
            "first time — and writes its path as the first line of the log. "
            "The log window refreshes ten times a second while open.",
            "Nothing.",
            _NONE,
            "Copying the default is best-effort: if your profile is read-only "
            "it uses the shipped copy and carries on. The rules file is not "
            "read until you click Preview or CLEAN.",
        ),
        RuntimeCard(
            "Choosing files",
            "Click Add Files... and pick one or more .docx files.",
            "Your system's file dialog. The paths are added to the list; the "
            "files are not opened yet.",
            "Nothing.",
            _NONE,
            "A file already in the list is not added twice. The dialog can "
            "show all files; one that is not a Word document is reported when "
            "it is processed, and the rest carry on.",
        ),
        RuntimeCard(
            "Removing files from the list",
            "Select files in the list and click Remove Selected.",
            "The selected paths are taken off the list.",
            "Nothing.",
            _NONE,
            "Nothing on disk is touched.",
        ),
        RuntimeCard(
            "Emptying the list",
            "Click Clear.",
            "Every path is taken off the list.",
            "Nothing.",
            _NONE,
            "Nothing on disk is touched.",
        ),
        RuntimeCard(
            "Choosing where cleaned files go",
            "Click Output Folder... and pick a folder.",
            "Your system's folder dialog. The folder is remembered until you close "
            "the program.",
            "Nothing.",
            _NONE,
            "Without a choice, each cleaned file goes beside its original.",
        ),
        RuntimeCard(
            "Opening the output folder",
            "Click Open Output Folder.",
            "Asks Windows to open the folder in File Explorer — the only other "
            "program SpecCleanse ever starts.",
            "Nothing from SpecCleanse. If the folder is on a network drive, "
            "Explorer reads it as usual.",
            _NONE,
            "If no folder is known yet, or it does not exist, the log says so "
            "and nothing opens.",
        ),
        RuntimeCard(
            "Accepting tracked changes and removing comments",
            "Tick Strip comments and accept tracked changes.",
            "Nothing yet. At the next Preview or CLEAN, every tracked change is "
            "accepted before detection — insertions kept, deletions removed — "
            "comment markers are removed, and on CLEAN the comment files are "
            "deleted from the package.",
            "Nothing.",
            _NONE,
            "Off every time the program starts. A deleted paragraph mark is "
            "not acted on, so two paragraphs are never merged. Deleted table "
            "rows and cells go whole.",
        ),
        RuntimeCard(
            "Previewing",
            "Click Preview.",
            f"Re-reads and checks {CONFIG_NAME}, then for each file: unpacks "
            "it into a temporary folder, runs every rule over the body, "
            "headers, footers, footnotes, endnotes and glossary, and lists "
            "what would be removed, which placeholders would be cut, and what "
            "is protected. Files are done one at a time, in list order.",
            "Nothing.",
            _NONE,
            "Writes nothing to your folders. A file that cannot be read is "
            "reported and the rest continue. Paragraphs in the log are "
            "shortened to 90 characters. There is no time limit.",
        ),
        RuntimeCard(
            "Cleaning — before anything is written",
            "Click CLEAN.",
            f"Works out every destination ({OUTPUT_NAME}) and checks the whole "
            "set before opening a single file. If two selected files would land "
            "on one destination, or a destination is another selected file, "
            "the run is refused. If cleaned files from an earlier run exist, it "
            f"asks before replacing them. Then {CONFIG_NAME} is re-read and "
            "checked.",
            "Nothing.",
            _NONE,
            "A refused batch writes nothing and the log names the files. "
            "Answering No to the overwrite question writes nothing. A broken "
            "rules file is reported, naming the section and entry, and nothing "
            "is processed.",
        ),
        RuntimeCard(
            "Cleaning — each file",
            "Nothing more; it follows the check above.",
            "One file at a time: unpack into a temporary folder; accept tracked "
            "changes if the box is ticked; run every rule; cut placeholders "
            "out of sentences, then remove flagged runs, then remove flagged "
            "paragraphs — emptying a paragraph instead of deleting it where "
            "deleting it would damage the file; pack the result as "
            f"{TEMP_NAME} and rename it into place; delete the temporary "
            "folder.",
            "Nothing.",
            _NONE,
            "Nothing is retried. A file that fails is reported once and the "
            "batch continues. A file open in Word is reported as such. The "
            "output appears under its real name only when complete.",
        ),
        RuntimeCard(
            "The check after each file (automatic)",
            "Nothing; it runs after every cleaned file.",
            "Unpacks the original and the output again, works out from the "
            "original what the rules permitted losing, and compares: removed "
            "paragraphs, shortened paragraphs, added paragraphs, structure "
            "Word depends on (section breaks, fields, bookmarks, tables), "
            "broken cross-references and numbering a removal may shift. Then "
            "it deletes both copies; if a deletion fails, it carries on "
            "without saying so.",
            "Nothing.",
            _NONE,
            "Never reads the cleaner's report. A file that does not pass is "
            "still written, marked Needs review, and its path is logged. If "
            "the check itself crashes, the file is marked Failed and the log "
            "calls the written file UNVERIFIED.",
        ),
        RuntimeCard(
            "When the batch finishes (automatic)",
            "Nothing; it follows the last file.",
            "One summary line counting Verified, Needs review (with each "
            "reason counted separately) and Failed files.",
            "Nothing.",
            _NONE,
            "The log is not saved anywhere. Select it and copy it if you want "
            "to keep it.",
        ),
        RuntimeCard(
            "Stopping",
            "Close the main window. There is no Stop button.",
            "The program exits. A file that is half-way through is abandoned.",
            "Nothing.",
            _NONE,
            "Files already finished stay. Closing mid-file can leave "
            f"{TEMP_NAME} beside the destination and a {TEMP_PREFIXES[0]} "
            "folder holding an unpacked copy of the document in your temporary "
            "folder; delete both by hand. A half-written file never takes the "
            "real name.",
        ),
        RuntimeCard(
            "About, licence and notices",
            "Click About, then License or Third-party notices; Close to finish.",
            "Shows text built into the program, and opens the licence and "
            "third-party notices bundled inside it in a read-only window.",
            "Nothing. The web addresses shown are text to copy; nothing is "
            "opened.",
            _NONE,
            "Usable even while a batch runs.",
        ),
        RuntimeCard(
            "This explanation",
            f"Click Why trust it?, then {_NOT_CONVINCED}; Close or Escape to "
            "go back.",
            "Shows text built into the program. The diagram is drawn by the "
            "program itself.",
            "Nothing. Like About, the addresses under Further reading are text "
            "to copy.",
            _NONE,
            "Each Escape closes one window. Usable even while a batch runs.",
        ),
        RuntimeCard(
            "Changing the rules (outside the program)",
            f"Open {CONFIG_NAME} in a text editor, edit it, save it.",
            "Nothing until your next Preview or CLEAN, which reads the file "
            "again and checks every pattern before processing.",
            "Nothing.",
            _NONE,
            "A mistake is reported by section and entry, and nothing is "
            "processed. Updating SpecCleanse never overwrites your copy, and "
            "uninstalling leaves it in place.",
        ),
        RuntimeCard(
            "Developer tools (source checkout only)",
            "Run python -m tools.census_formatting or another tool in the "
            "tools folder from a command prompt.",
            "Reads the documents you name and reports what a rule change would "
            "do. Not part of the installed program.",
            "Nothing.",
            _NONE,
            "Never modifies a document. Two tools write a report file, only "
            "where you tell them to.",
        ),
    ),
)

_REACH = Section(
    "reach", "Blast radius", "What the program may touch, and what it cannot",
    (
        Para(
            "There is no AI to constrain. The question that matters is what "
            "the program itself can reach."
        ),
        Table(
            ("Capability", "What it can do", "Constraint"),
            (
                ("Read files",
                 f"The .docx files you pick, and {CONFIG_NAME}.",
                 "Originals are opened read-only."),
                ("Write files",
                 f"{OUTPUT_NAME} (via {TEMP_NAME}); its own temporary folders; "
                 f"on the installed program's first run, the default {USER_CONFIG}.",
                 "Destinations are checked for collisions first; existing "
                 "cleaned files are replaced only after you confirm."),
                ("Delete files",
                 "Its own temporary folders and its own .tmp file.",
                 "Nothing you chose is ever deleted."),
                ("Start other programs",
                 "File Explorer, from Open Output Folder.",
                 "No other program is started."),
                ("Edit document content",
                 "Remove paragraphs, runs and placeholder text matched by the "
                 "rules; with the box ticked, accept tracked changes and "
                 "remove comments.",
                 "Protected headings are never removed or trimmed; "
                 "structure-bearing paragraphs are emptied, not deleted."),
            ),
        ),
        Para(
            "It categorically cannot: open a network connection, run a macro "
            "or anything else inside a document, change the original, write "
            "outside the folders above, or keep anything of its own between "
            "sessions other than the rules file."
        ),
    ),
)

_TERMS = Section(
    "terms", "Definitions", "What the key terms mean, exactly",
    (
        Term(
            OUTCOME_NAMES[FileOutcome.VERIFIED],
            "The file was written, and every difference between it and the "
            "original was accounted for by a configured rule, with no new "
            "structural problem found.",
            "Not that the rules are right, not that the document is correct, "
            "and not that Word will open it without complaint.",
        ),
        Term(
            "PASS",
            "The log's word for a file whose check found nothing unexplained.",
            "The same limits as Verified.",
        ),
        Term(
            OUTCOME_NAMES[FileOutcome.NEEDS_REVIEW],
            "The file was written, but something needs a person: "
            f"{CATEGORY_NAMES[ReviewCategory.DETECTED_DAMAGE]} (text gone or "
            "added, and the comparison knew exactly where), "
            f"{CATEGORY_NAMES[ReviewCategory.AMBIGUOUS_ALIGNMENT]} (something "
            "unexplained, but which paragraph it belongs to is a guess), "
            f"{CATEGORY_NAMES[ReviewCategory.CONFIGURATION]} (your rules can "
            "remove text on looks alone), or "
            f"{CATEGORY_NAMES[ReviewCategory.REFERENCE_NUMBERING]} (a "
            "cross-reference broke or numbering may shift).",
            "Not necessarily that anything is wrong — only that the program "
            "cannot vouch for it.",
        ),
        Term(
            OUTCOME_NAMES[FileOutcome.FAILED],
            "Processing failed, or the check could not run. If a file was "
            "written anyway, the log calls it UNVERIFIED.",
            "Not that your original was harmed; it is never written to.",
        ),
        Term(
            "Expected removal",
            "A loss the rules permit, worked out from the original: a rule "
            "matched the whole paragraph, or the permitted pieces cover every "
            "character it had.",
            "Not that the removal was a good idea. A wrong rule produces "
            "expected removals.",
        ),
        Term(
            "Preserve violation",
            "A protected paragraph — matched by a preserve pattern or style, "
            "such as PART and SECTION headings — lost text.",
            "Protection covers only what the preserve rules name.",
        ),
        Term(
            "Formatting-only",
            "Preview's label for a removal made on italic plus an editorial "
            "colour alone, with no pattern or style behind it. Possible only "
            "if you switch formatting_only_removal on.",
            "Not a judgment that the text was a note — only that it looked "
            "like one.",
        ),
        Term(
            "Score (confidence)",
            "A number from the rules: points for each pattern, style or "
            "formatting signal. At 0.5 or above the item is removed.",
            "Not a probability, and not a measure of how likely the removal is "
            "to be right.",
        ),
        Term(
            "Preview",
            "Exactly the detection CLEAN would do, with nothing written.",
            "Preview does not run the check; only CLEAN does.",
        ),
    ),
)

_NO_AI = Section(
    "no-ai", "No AI here", "The parts with no AI in them",
    (
        Para("All of it. Specifically:"),
        Bullets((
            "Finding hidden text, from Word's own hidden setting on the text or "
            "its style.",
            "Matching specifier notes, copyright lines, SpecAgent references and "
            "editing instructions with the patterns in your rules file.",
            "Cutting placeholders such as [Verify quantity] out of sentences.",
            "Protecting PART, SECTION, ARTICLE and END OF SECTION headings.",
            "Deciding when a paragraph must be emptied rather than deleted.",
            "Accepting tracked changes and removing comments, when you ask.",
            "Checking destinations before anything is written.",
            "The check after every clean, and the summary.",
        )),
        Note(
            "With the network disconnected",
            "Everything works: Preview, CLEAN, the check, About and this "
            "explanation. Only files on a network drive become unreachable, as "
            "they would for any program.",
        ),
    ),
)

_SECURITY = Section(
    "security", "Security", "Security and privacy, mechanism by mechanism",
    (
        Table(
            ("Concern", "How it is handled"),
            (
                ("Passwords, keys, accounts",
                 "There are none. Nothing to sign in to, nothing stored."),
                ("A local server or open port",
                 "None. The program contains no networking code to open one."),
                ("A malicious .docx",
                 "Nothing inside a document is run. Python's unzip strips "
                 "drive letters, absolute paths and \"..\" from entry names, so "
                 "no entry can land outside the temporary folder. The XML "
                 "parser does not fetch external references."),
                ("A malicious rules file",
                 "Read in safe mode, which accepts plain values only."),
                ("Macros and embedded objects",
                 "Never run. Carried through to the output unchanged."),
                ("Copies of your documents",
                 "Temporary folders only, deleted after each file. Closing "
                 "the program, or a crash, part-way through a file can leave "
                 "one behind (see Stopping). So can a deletion Windows refuses "
                 "— antivirus holding a file, say: the cleaner reports that as "
                 "an error, but the check leaves its copies without a word."),
                ("What is logged",
                 "Paragraph text and file paths, in the window only, never "
                 "written to disk."),
                ("Administrator rights",
                 "Not needed. The installer installs for you alone by default."),
                ("Is this the program that was published?",
                 "The executable is not code-signed. Releases are built from a "
                 "tagged commit by the project's GitHub Actions workflow, and you "
                 "can build your own from the source."),
                ("Runtime libraries",
                 "Pinned exactly: lxml 6.0.2 and PyYAML 6.0.3."),
            ),
        ),
    ),
)

_LIMITS = Section(
    "limits", "Honesty", "What this does not do",
    (
        Bullets((
            "**The check shares the cleaner's rules.** A rule that matches "
            "real requirement text removes it, and the check calls that "
            "expected. Guard: run Preview first, read the removals, and keep "
            "formatting-only removal off unless your notes are marked only that "
            "way.",
            "**Rules can match real requirements.** Several shipped rules were "
            "narrowed after deleting real text; an older copy of your rules "
            "file may still carry them. Guard: a configuration notice names any "
            "such rule still active, and every file in that run needs review.",
            "**Not validated on real documents.** No real specification "
            "collection has been measured, and no automated test opens a "
            "cleaned file in Word. Guard: the manual Word checklist in the "
            "README.",
            "**No Stop button, no limits.** There is no size or time limit: a "
            "very large or deliberately crafted file is unpacked in full, and "
            "a badly written pattern can run for a very long time. Closing the "
            "window is the only stop, and it can leave a temporary copy behind.",
            "**Only .docx.** Not .doc. The file check only asks whether the "
            "archive holds a Word document part.",
            "**Not a privacy scrubber.** Author, company and other document "
            "properties are copied unchanged, as are pictures, embedded files "
            "and anything else it does not edit.",
            "**Empty paragraphs can remain.** Where deleting a paragraph would "
            "damage the file, it is emptied instead.",
            "**Some notes survive on purpose.** A placeholder that straddles a "
            "page break is left whole, with a note in the log. A note that no "
            "rule describes is not removed.",
            "**Numbers and references may shift.** Removing a numbered "
            "paragraph can change the numbers a reader sees; removing a bookmarked paragraph "
            "breaks references to it. Both are reported; neither is repaired.",
            "**Ambiguous alignment can mislead.** When the comparison falls "
            "back to guessing which paragraph became which (similarity below "
            "0.5 is treated as no match), the fragments it quotes may be "
            "artefacts of the guess. It says so: those files are labelled "
            "ambiguous alignment, not damage.",
            "**The log is not kept.** Close the window and it is gone.",
            "**Accessibility.** The window is built with Tkinter, which "
            "exposes little to screen readers.",
        )),
    ),
)

_CHECK = Section(
    "check", "The point", "Check it yourself",
    (
        Bullets((
            "**Read the log** (a minute). Its first line names the rules file in "
            "use; each file shows what was removed, what was cut and what was "
            "protected, and the check's verdict.",
            "**Preview before you clean** (a minute). Every removal is listed "
            "under the kind of rule that caused it, with its score, and nothing "
            "is written.",
            "**Compare in Word** (five minutes). Review → Compare → Compare…, "
            "original against cleaned. With the tracked-changes box off, you "
            "should see deletions only — no insertions.",
            f"**Open the rules** (five minutes). {CONFIG_NAME} is plain text; "
            "every pattern is listed there, with comments explaining each "
            "group.",
            "**Pull the network plug** (five minutes). Disconnect, then run "
            "Preview and CLEAN on a local file. Everything works.",
            "**Watch the network** (ten minutes). Open Resource Monitor → "
            "Network, or your firewall's log, and run a batch. SpecCleanse "
            "makes no connections of its own.",
            "**Look inside the files** (fifteen minutes). Copy a .docx, rename "
            "it .zip, open it, and read word/document.xml before and after.",
            "**Read the code and run the tests** (an afternoon). The whole "
            "program is a handful of Python files. The test suite runs with "
            "python -m unittest discover -s tests -t . and "
            "docs/TRUST_CLAIMS.md lists every claim here with the code and test "
            "behind it.",
        )),
        LinkBox(
            "Further reading",
            (
                Link("Source code", about.PROJECT_URL),
                Link("Releases and release notes", about.RELEASES_URL),
                Link("Report a problem", about.SUPPORT_URL),
                Link("The claims ledger",
                     f"{about.PROJECT_URL}/blob/master/docs/TRUST_CLAIMS.md"),
                Link("Python's zipfile (how entries are unpacked)",
                     "https://docs.python.org/3/library/zipfile.html"),
                Link("lxml's parser options (network access, entities)",
                     "https://lxml.de/parsing.html"),
                Link("PyYAML (safe loading)",
                     "https://pyyaml.org/wiki/PyYAMLDocumentation"),
            ),
        ),
        Note(
            "Still not convinced?",
            "Good. Treat every cleaned file as a draft until you have compared "
            "it with its original. Run Preview on a section you know well and "
            "see whether every listed removal is one you would have made. Keep "
            "the original beside the cleaned copy, never save over it, and "
            "report anything the program removed that it should not have — "
            "that is how a rule gets narrowed.",
        ),
    ),
)

DOSSIER = (
    _START,
    _PROVENANCE,
    _ENGINE,
    _BOUNDARY,
    _RUNTIME,
    _REACH,
    _TERMS,
    _NO_AI,
    _SECURITY,
    _LIMITS,
    _CHECK,
)


def runtime_cards() -> list[RuntimeCard]:
    """Every runtime card, in order."""
    return [block for block in _RUNTIME.blocks if isinstance(block, RuntimeCard)]


def further_reading() -> list[Link]:
    """The only external addresses either trust surface shows."""
    return [
        link
        for section in DOSSIER
        for block in section.blocks if isinstance(block, LinkBox)
        for link in block.links
    ]


def _block_texts(block) -> list[str]:
    if isinstance(block, Point):
        return [block.title, block.body]
    if isinstance(block, (Para,)):
        return [block.text]
    if isinstance(block, Bullets):
        return list(block.items)
    if isinstance(block, Table):
        return [*block.columns, *(cell for row in block.rows for cell in row)]
    if isinstance(block, Contrast):
        return [block.left_title, *block.left, block.right_title, *block.right]
    if isinstance(block, Note):
        return [block.title, block.text]
    if isinstance(block, Term):
        return [block.word, block.claim, block.not_claimed]
    if isinstance(block, RuntimeCard):
        return [block.title, *(getattr(block, field) for field, _ in CARD_ROWS)]
    if isinstance(block, Diagram):
        return [block.description]
    if isinstance(block, LinkBox):
        return [block.title, *(link.label for link in block.links)]
    raise TypeError(f"unknown block {block!r}")


def all_copy() -> list[str]:
    """Every string either level shows, except the further-reading addresses."""
    texts = [LEVEL1_TITLE, LEVEL1_LEAD, LEVEL1_CLOSING_TITLE, LEVEL1_CLOSING_TEXT,
             LEVEL1_BUTTON, DOSSIER_TITLE, DOSSIER_SUBTITLE]
    for point in LEVEL1_POINTS:
        texts += [point.title, point.body]
    for section in DOSSIER:
        texts += [section.kicker, section.title]
        for block in section.blocks:
            texts += _block_texts(block)
    return texts
