"""The trust surfaces say only what the code does.

"Why trust it?" and the dossier behind it (``trust.py``, drawn by
``trust_view.py``) make claims a reader may check against a firewall log or the
source. Every number, name and location they quote is held here to the code it
comes from, so changing the code fails a test until the words follow.
``docs/TRUST_CLAIMS.md`` lists each claim with the test that pins it.

Three kinds of test, by what they need:

* most need nothing beyond the project's own modules and run everywhere;
* a few need ``gui.py`` importable, which needs tkinter;
* the dialog behaviour — focus, Escape, the contents rail — needs a display as
  well, and skips without one, saying so.
"""

import ast
import re
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path, PureWindowsPath
from unittest.mock import patch

from lxml import etree

import about
import apppaths
import batch
import trust
import verify
from detection import ContentType, Detection
from docx_xml import W, collect_content_parts, load_config, parse_xml
from processor import DocxProcessor
from tests import docx_builder as db
from tests.support import CONFIG_PATH, PROJECT_ROOT, DocxTestCase

try:
    import gui
    GUI_IMPORTABLE = True
except ImportError:  # pragma: no cover - environments without Tk
    GUI_IMPORTABLE = False


LEDGER = (PROJECT_ROOT / "docs" / "TRUST_CLAIMS.md").read_text(encoding="utf-8")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

DOCUMENT = db.document(
    db.text_para("SECTION 21 13 13 - WET-PIPE SPRINKLER SYSTEMS"),
    db.text_para("PART 1 - GENERAL"),
    db.text_para("Retain subparagraph below for wet-pipe systems."),
    db.text_para("Provide two [Verify quantity with Owner] spare sprinklers."),
    db.para(db.run("Coordinate hangers with structural.", italic=True, color="FF0000")),
    db.para(db.run("Hidden note.", vanish=True), db.run(" Provide gauges.")),
    db.text_para("END OF SECTION 21 13 13"),
)


def _copy() -> str:
    return "\n".join(trust.all_copy())


def _root_modules() -> list[Path]:
    return sorted(PROJECT_ROOT.glob("*.py"))


def _tree(path: Path) -> ast.Module:
    return ast.parse(path.read_text(encoding="utf-8"), filename=str(path))


def _imported_roots(tree: ast.Module) -> set[str]:
    roots = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            roots |= {alias.name.split(".")[0] for alias in node.names}
        elif isinstance(node, ast.ImportFrom) and node.module and node.level == 0:
            roots.add(node.module.split(".")[0])
    return roots


def _enclosing_functions(tree: ast.Module):
    """Yield ``(function name or None, node)`` for every node in the module."""
    def visit(node, owner):
        for child in ast.iter_child_nodes(node):
            is_function = isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef))
            name = child.name if is_function else owner
            yield name, child
            yield from visit(child, name)
    yield from visit(tree, None)


def _dotted(node) -> str:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        return f"{_dotted(node.value)}.{node.attr}"
    return ""


# ---------------------------------------------------------------------------
# The boundary
# ---------------------------------------------------------------------------

NETWORK_MODULES = {
    "socket", "ssl", "http", "urllib", "urllib3", "ftplib", "smtplib", "poplib",
    "imaplib", "telnetlib", "nntplib", "xmlrpc", "socketserver", "selectors",
    "asyncio", "webbrowser", "requests", "httpx", "aiohttp", "email",
}


class NoNetworkTests(unittest.TestCase):
    """C01–C03, C18: nothing the program ships can reach the network."""

    def test_shipped_modules_import_no_networking_module(self):
        for path in _root_modules():
            with self.subTest(module=path.name):
                found = _imported_roots(_tree(path)) & NETWORK_MODULES
                self.assertEqual(found, set(), f"{path.name} imports {found}; "
                                 "the trust copy says the program has no network code")

    def test_every_root_module_is_scanned(self):
        # The scan above globs the root, so a new module is covered the moment
        # it exists.  This pins the ones the ledger names, so a rename is seen.
        names = {path.name for path in _root_modules()}
        for name in ("gui.py", "trust.py", "trust_view.py", "theme.py", "about.py",
                     "apppaths.py", "batch.py", "detection.py", "processor.py",
                     "verify.py", "docx_xml.py"):
            self.assertIn(name, names)

    def test_only_open_output_folder_starts_a_process(self):
        starters = set()
        for path in _root_modules():
            for owner, node in _enclosing_functions(_tree(path)):
                if not isinstance(node, ast.Call):
                    continue
                name = _dotted(node.func)
                if (name.startswith("subprocess.") or name in {
                        "os.startfile", "os.system", "os.popen"}
                        or name.startswith(("os.exec", "os.spawn"))):
                    starters.add((path.name, owner))
        self.assertEqual(starters, {("gui.py", "_open_output_dir")})

    def test_gui_writes_no_file_of_its_own(self):
        # The window keeps nothing: no log file, no settings, no history.  Its
        # only filesystem writes are the preview's temporary folder and its
        # removal; everything else is the processor's, writing the output.
        writers = {"open", "write_text", "write_bytes", "mkdir", "makedirs",
                   "copyfile", "copy2", "copytree", "move", "rename", "touch",
                   "dump", "mkdtemp", "rmtree", "unlink"}
        for name in ("gui.py", "trust.py", "trust_view.py", "theme.py", "about.py"):
            calls = set()
            for _, node in _enclosing_functions(_tree(PROJECT_ROOT / name)):
                if isinstance(node, ast.Call):
                    leaf = _dotted(node.func).rsplit(".", 1)[-1]
                    if leaf in writers:
                        calls.add(leaf)
            expected = {"mkdtemp", "rmtree"} if name == "gui.py" else set()
            with self.subTest(module=name):
                self.assertEqual(calls, expected)

    def test_every_network_exception_is_in_the_boundary(self):
        boundary = next(s for s in trust.DOSSIER if s.id == "boundary")
        contrast = next(b for b in boundary.blocks if isinstance(b, trust.Contrast))
        diagram = next(b for b in boundary.blocks if isinstance(b, trust.Diagram))
        leaves = " ".join(contrast.left).lower()
        first_point = trust.LEVEL1_POINTS[0].body.lower()
        self.assertEqual(len(trust.NETWORK_EXCEPTIONS), 2)
        self.assertIn("two things", first_point)
        for exception in trust.NETWORK_EXCEPTIONS:
            with self.subTest(exception=exception):
                self.assertIn(exception.lower(), leaves)
                self.assertIn(exception.lower(), diagram.description.lower())
                self.assertIn(exception.lower(), first_point)
        drawing = (PROJECT_ROOT / "trust_view.py").read_text(encoding="utf-8")
        self.assertIn("Windows SmartScreen", drawing)
        self.assertIn("Network drive or synced folder", drawing)


class UntrustedInputTests(DocxTestCase):
    """C04, C60: a document cannot reach outside its temporary folder."""

    def test_the_xml_parser_does_not_resolve_external_entities(self):
        secret = self.temp_dir / "secret.txt"
        secret.write_text("TOPSECRET", encoding="utf-8")
        part = self.temp_dir / "part.xml"
        part.write_text(
            f'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e SYSTEM "{secret.as_uri()}">]>'
            "<x>&e;</x>", encoding="utf-8")
        # lxml leaves the external entity unloaded, so the part either fails to
        # parse — reported as a processing error for that file — or parses
        # without the file's contents.  Either way nothing is read from outside.
        try:
            parsed = etree.tostring(parse_xml(part))
        except etree.XMLSyntaxError:
            return
        self.assertNotIn(b"TOPSECRET", parsed)

    def test_a_zip_entry_cannot_escape_the_temp_folder(self):
        source = self.build(DOCUMENT)
        with zipfile.ZipFile(source, "a") as zf:
            zf.writestr("../../escaped.txt", "out")
        sandbox = self.temp_dir / "sandbox"
        sandbox.mkdir()
        with patch.object(tempfile, "tempdir", str(sandbox)):
            result = DocxProcessor(self.make_engine()).process(
                source, self.temp_dir / "out" / "x_cleaned.docx")
        self.assertEqual(result.errors, [])
        for where in (sandbox, self.temp_dir):
            self.assertFalse((where / "escaped.txt").exists())


class ConfigSafetyTests(DocxTestCase):
    """C31, C61."""

    def test_patterns_yaml_cannot_construct_python_objects(self):
        bad = self.temp_dir / "patterns.yaml"
        bad.write_text("x: !!python/object/apply:os.getcwd []\n", encoding="utf-8")
        with self.assertRaises(ValueError):
            load_config(bad)

    def test_patterns_are_case_insensitive_and_dotall(self):
        engine = self.make_engine(specagent_references={"text_patterns": ["^spec agent"]})
        detector = next(d for d in engine.detectors if d.content_type == ContentType.SPECAGENT)
        para = etree.fromstring(f'<w:p xmlns:w="{W_NS}"/>')
        self.assertIsNotNone(detector.detect(para, "SPEC AGENT\nsecond line"))
        self.assertIsNone(detector.detect(para, "first line\nspec agent"))


# ---------------------------------------------------------------------------
# What happens to your files
# ---------------------------------------------------------------------------

class OriginalUntouchedTests(DocxTestCase):
    """C10, C12, C16, C74."""

    def test_clean_and_verify_leave_the_input_byte_identical(self):
        source = self.build(DOCUMENT)
        before = (source.read_bytes(), source.stat().st_mtime_ns)
        _, output = self.clean(source)
        verify.verify_clean(source, output, engine=self.make_engine())
        self.assertEqual((source.read_bytes(), source.stat().st_mtime_ns), before)

    def test_no_tmp_file_is_left_after_a_clean(self):
        _, output = self.clean(self.build(DOCUMENT))
        self.assertTrue(output.exists())
        self.assertEqual(list(output.parent.glob("*.tmp")), [])
        self.assertEqual(trust.TEMP_NAME, trust.OUTPUT_NAME + ".tmp")

    def test_unchanged_parts_are_copied_byte_for_byte(self):
        core = ('<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties '
                'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>A. Specifier</dc:creator>'
                '</cp:coreProperties>')
        source = self.build(DOCUMENT, {"docProps/core.xml": core})
        _, output = self.clean(source)
        with zipfile.ZipFile(source) as a, zipfile.ZipFile(output) as b:
            self.assertEqual(a.read("docProps/core.xml"), b.read("docProps/core.xml"))
            self.assertNotEqual(a.read("word/document.xml"), b.read("word/document.xml"))


class OnlyRemovesTests(DocxTestCase):
    """C43: the cleaner deletes; it never writes a word."""

    @staticmethod
    def _is_subsequence(short: str, long: str) -> bool:
        letters = iter(long)
        return all(char in letters for char in short)

    def test_every_cleaned_paragraph_comes_from_the_original(self):
        source = self.build(DOCUMENT)
        _, output = self.clean(source)
        before = self.paragraph_texts(source)
        after = self.paragraph_texts(output)
        self.assertLess(len(after), len(before))
        for text in after:
            with self.subTest(text=text):
                self.assertTrue(any(self._is_subsequence(text, b) for b in before))

    def test_verification_fails_an_added_paragraph(self):
        source = self.build(DOCUMENT, name="in.docx")
        invented = db.document(
            *[db.text_para(t) for t in self.paragraph_texts(source)],
            db.text_para("Provide gold-plated sprinklers."),
        )
        output = self.build(invented, name="out.docx")
        result = verify.verify_clean(source, output, engine=self.make_engine())
        self.assertFalse(result.passed)
        self.assertTrue(result.added)


class DeterminismTests(DocxTestCase):
    """C32."""

    def test_two_runs_make_the_same_decisions(self):
        source = self.build(DOCUMENT)

        def decisions():
            return [(d.content_type, d.text, d.confidence, tuple(d.spans))
                    for d in self.detect(source)]

        self.assertEqual(decisions(), decisions())


class TempFolderTests(DocxTestCase):
    """C17."""

    def test_core_temp_prefixes_match_the_copy(self):
        prefixes = []
        real = tempfile.mkdtemp

        def recording(*args, **kwargs):
            prefixes.append(kwargs.get("prefix"))
            return real(*args, **kwargs)

        source = self.build(DOCUMENT)
        with patch("tempfile.mkdtemp", side_effect=recording):
            _, output = self.clean(source)
            verify.verify_clean(source, output, engine=self.make_engine())
        self.assertEqual(set(prefixes), set(trust.TEMP_PREFIXES[:2]))

    def test_temp_folders_are_removed_after_a_clean(self):
        source = self.build(DOCUMENT)
        sandbox = self.temp_dir / "sandbox"
        sandbox.mkdir()
        with patch.object(tempfile, "tempdir", str(sandbox)):
            _, output = self.clean(source)
            verify.verify_clean(source, output, engine=self.make_engine())
            DocxProcessor(self.make_engine(), dry_run=True).process(
                source, self.temp_dir / "unused.docx")
        self.assertEqual(list(sandbox.iterdir()), [])


# ---------------------------------------------------------------------------
# The numbers
# ---------------------------------------------------------------------------

def _run(*props: str) -> etree._Element:
    return etree.fromstring(
        f'<w:r xmlns:w="{W_NS}"><w:rPr>{"".join(props)}</w:rPr><w:t>x</w:t></w:r>')


class NumbersTests(DocxTestCase):
    """Every number and name the copy quotes, against its source."""

    def _detector(self, engine, kind):
        engine.bind_styles({})
        return next(d for d in engine.detectors if d.content_type == kind)

    def test_removal_threshold(self):
        engine = self.make_engine()
        el = etree.fromstring(f'<w:p xmlns:w="{W_NS}"/>')

        def removes(score):
            return engine.should_remove(
                [Detection(ContentType.SPECIFIER_NOTE, el, "x", score, "")])

        self.assertTrue(removes(0.5))
        self.assertFalse(removes(0.49))
        self.assertIn("removed at 0.5 or above", _copy())

    def test_formatting_alone_scores_exactly_the_threshold(self):
        on = self.make_engine(specifier_notes={"formatting_only_removal": True})
        notes = self._detector(on, ContentType.SPECIFIER_NOTE)
        both = notes.detect(_run("<w:i/>", '<w:color w:val="FF0000"/>'), "x")
        self.assertEqual(both.confidence, 0.5)
        self.assertIsNone(notes.detect(_run("<w:i/>"), "x"))                     # 0.2
        self.assertIsNone(notes.detect(_run('<w:color w:val="FF0000"/>'), "x"))  # 0.3
        self.assertIn("italic adds 0.2", _copy())
        self.assertIn("adds 0.3", _copy())
        self.assertIn("together exactly 0.5", _copy())

    def test_formatting_only_removal_is_off_by_default(self):
        self.assertIs(load_config(CONFIG_PATH)["specifier_notes"]["formatting_only_removal"], False)
        notes = self._detector(self.make_engine(), ContentType.SPECIFIER_NOTE)
        self.assertIsNone(notes.detect(_run("<w:i/>", '<w:color w:val="FF0000"/>'), "x"))
        self.assertIn("by default looks alone are never enough", _copy())

    def test_low_confidence_needs_formatting(self):
        editorial = self._detector(self.make_engine(), ContentType.EDITORIAL_ARTIFACT)
        text = "Provide pumps and revise as required."
        self.assertIsNone(editorial.detect(_run(), text))
        coloured = editorial.detect(_run('<w:color w:val="FF0000"/>'), text)
        self.assertAlmostEqual(coloured.confidence, 0.3 + 0.3)
        self.assertIn("low-confidence score 0.3", _copy())

    def test_preserve_wins(self):
        engine = self.make_engine()
        engine.bind_styles({})
        para = etree.fromstring(f'<w:p xmlns:w="{W_NS}"/>')
        found = engine.detect_in_element(para, "PART 1 - GENERAL [Specifier: delete]")
        self.assertEqual([d.content_type for d in found], [ContentType.PRESERVE])
        self.assertFalse(engine.should_remove(found))

    def test_editorial_colours(self):
        shipped = load_config(CONFIG_PATH)["specifier_notes"]["formatting_signals"]["colors"]
        self.assertEqual([code for _, code in trust.EDITORIAL_COLOURS], shipped)
        for _, code in trust.EDITORIAL_COLOURS:
            self.assertIn(code, _copy())

    def test_output_suffix(self):
        self.assertEqual(batch.output_for(Path("x.docx")).name,
                         trust.OUTPUT_NAME.replace("<name>", "x"))

    def test_config_locations(self):
        self.assertEqual(trust.USER_CONFIG, str(PureWindowsPath(
            "%APPDATA%", apppaths.APP_NAME, apppaths.CONFIG_NAME)))
        with patch.dict("os.environ", {"APPDATA": "appdata"}):
            self.assertEqual(apppaths.user_config_dir(), Path("appdata") / apppaths.APP_NAME)

    def test_uninstall_keeps_patterns(self):
        iss = (PROJECT_ROOT / "packaging" / "installer.iss").read_text(encoding="utf-8")
        section = iss.split("[UninstallDelete]", 1)[1].split("\n[", 1)[0]
        entries = [line for line in section.splitlines()
                   if line.strip() and not line.lstrip().startswith(";")]
        self.assertEqual(entries, [])

    def test_installer_is_per_user_by_default(self):
        iss = (PROJECT_ROOT / "packaging" / "installer.iss").read_text(encoding="utf-8")
        self.assertIn("PrivilegesRequired=lowest", iss.splitlines())

    def test_executable_is_unsigned(self):
        spec = (PROJECT_ROOT / "packaging" / "speccleanse.spec").read_text(encoding="utf-8")
        self.assertIn("codesign_identity=None", spec)
        self.assertIn("not code-signed", _copy())

    def test_comment_parts(self):
        from docx_xml import COMMENT_PARTS
        self.assertIn("comments.xml", COMMENT_PARTS)
        self.assertIn("comment files are deleted", _copy())

    def test_content_parts(self):
        word = self.temp_dir / "word"
        (word / "glossary").mkdir(parents=True)
        names = ["document.xml", "header1.xml", "footer1.xml", "footnotes.xml",
                 "endnotes.xml", "glossary/document.xml", "comments.xml", "settings.xml"]
        for name in names:
            (word / name).write_text("<x/>", encoding="utf-8")
        found = {p.relative_to(word).as_posix() for p in collect_content_parts(word)}
        self.assertEqual(found, set(names[:6]))
        self.assertIn("body, headers, footers, footnotes, endnotes and glossary", _copy())

    def test_outcome_and_category_names(self):
        for outcome in batch.FileOutcome:
            self.assertIn(trust.OUTCOME_NAMES[outcome], _copy())
        for category in batch.ReviewCategory:
            self.assertIn(category.value, _copy())

    def test_pairing_similarity(self):
        self.assertEqual(verify.MIN_PAIR_SIMILARITY, 0.5)
        self.assertIn("similarity below 0.5", _copy())

    def test_pinned_versions(self):
        pins = dict(line.split("==") for line in
                    (PROJECT_ROOT / "requirements.txt").read_text(encoding="utf-8").split())
        self.assertIn(f"lxml {pins['lxml']}", _copy())
        self.assertIn(f"PyYAML {pins['PyYAML']}", _copy())

    def test_log_refresh_interval(self):
        tree = _tree(PROJECT_ROOT / "gui.py")
        delays = [node.args[0].value for owner, node in _enclosing_functions(tree)
                  if owner == "_drain_log" and isinstance(node, ast.Call)
                  and _dotted(node.func).endswith(".after")]
        self.assertEqual(delays, [100])
        self.assertIn("refreshes ten times a second", _copy())


class ShippedRulesTests(unittest.TestCase):
    """C30, C39."""

    def test_counts_match_patterns_yaml(self):
        shipped = load_config(CONFIG_PATH)
        for (section, key), count in trust.SHIPPED_RULE_COUNTS.items():
            with self.subTest(section=section, key=key):
                self.assertEqual(len(shipped[section][key]), count)

    def test_every_detector_section_is_counted(self):
        shipped = load_config(CONFIG_PATH)
        listed = {(section, key) for section, body in shipped.items()
                  for key, value in body.items() if isinstance(value, list)
                  and key not in ("colors",)}
        self.assertEqual(listed, set(trust.SHIPPED_RULE_COUNTS))


# ---------------------------------------------------------------------------
# The shape of the two levels
# ---------------------------------------------------------------------------

QUOTED_NUMBERS = {"0.5", "0.2", "0.3", "90", "6.0.2", "6.0.3",
                  "8", "6", "1", "13", "12", "16", "5", "4", "3"}

DOSSIER_ORDER = ["start", "provenance", "engine", "boundary", "runtime", "reach",
                 "terms", "no-ai", "security", "limits", "check"]


def _button_texts(path: Path) -> set[str]:
    """Every literal label on a button or checkbox the file creates."""
    texts = set()
    for _, node in _enclosing_functions(_tree(path)):
        if not isinstance(node, ast.Call):
            continue
        name = _dotted(node.func)
        label = None
        if name.endswith(("Button", "Checkbutton")):
            for keyword in node.keywords:
                if keyword.arg == "text":
                    label = keyword.value
        elif name == "_button" and len(node.args) > 1:
            label = node.args[1]
        if isinstance(label, ast.Constant) and isinstance(label.value, str):
            texts.add(label.value)
        elif isinstance(label, ast.Attribute) and _dotted(label.value) == "trust":
            texts.add(getattr(trust, label.attr))
    return texts


class DossierStructureTests(unittest.TestCase):

    def test_level1_shape(self):
        self.assertTrue(6 <= len(trust.LEVEL1_POINTS) <= 8)
        for point in trust.LEVEL1_POINTS:
            with self.subTest(point=point.title):
                self.assertTrue(5 <= len(point.title.split()) <= 8)
                self.assertTrue(point.body)
        self.assertEqual(trust.LEVEL1_CLOSING_TITLE, "Not convinced?")
        self.assertEqual(trust.LEVEL1_BUTTON, "I'm not convinced — show me exactly what runs →")
        self.assertTrue(trust.LEVEL1_CLOSING_TEXT.startswith("Fair."))

    def test_sections_in_order_each_with_kicker_and_title(self):
        self.assertEqual([s.id for s in trust.DOSSIER], DOSSIER_ORDER)
        for section in trust.DOSSIER:
            with self.subTest(section=section.id):
                self.assertTrue(section.kicker and section.title and section.blocks)
                self.assertLessEqual(len(section.kicker.split()), 3)

    def test_every_card_has_the_five_rows(self):
        self.assertEqual([label for _, label in trust.CARD_ROWS],
                         ["You do", "What runs", "What is sent", "AI involved", "Bounded by"])
        cards = trust.runtime_cards()
        self.assertTrue(12 <= len(cards) <= 20)
        for card in cards:
            for field, label in trust.CARD_ROWS:
                with self.subTest(card=card.title, row=label):
                    self.assertTrue(getattr(card, field).strip())

    def test_every_card_says_ai_none(self):
        # No model is used anywhere; the day one is, this test is where the
        # card for it has to be written.
        for card in trust.runtime_cards():
            self.assertEqual(card.ai_involved, "None.", card.title)

    def test_every_button_has_a_runtime_card(self):
        labels = _button_texts(PROJECT_ROOT / "gui.py") | _button_texts(PROJECT_ROOT / "trust_view.py")
        self.assertIn("CLEAN", labels)
        self.assertIn(trust.LEVEL1_BUTTON, labels)
        you_do = " ".join(card.you_do for card in trust.runtime_cards())
        for label in labels:
            with self.subTest(button=label):
                self.assertIn(label, you_do)

    def test_automatic_behaviours_have_cards(self):
        titles = " ".join(card.title for card in trust.runtime_cards())
        for word in ("Starting", "The check after each file", "When the batch finishes",
                     "Stopping", "Changing the rules"):
            self.assertIn(word, titles)

    def test_every_number_in_the_copy_is_accounted_for(self):
        text = re.sub(r"\b[0-9A-F]{6}\b", "", _copy())   # the colours, pinned apart
        numbers = set(re.findall(r"\d+(?:\.\d+)*", text))
        self.assertEqual(numbers - QUOTED_NUMBERS, set(),
                         "a number in the copy that no test pins; add it to the ledger "
                         "and to QUOTED_NUMBERS with a test")
        numbers_table = LEDGER.split("## Numbers quoted in the copy", 1)[1]
        for number in QUOTED_NUMBERS:
            self.assertRegex(numbers_table, rf"(?<![\d.]){re.escape(number)}(?![\d.])")

    def test_every_test_the_ledger_names_exists(self):
        # The ledger is a contract too: a renamed or deleted test must not
        # leave a claim pointing at nothing.
        import importlib
        classes: dict[str, type] = {}
        for path in sorted((PROJECT_ROOT / "tests").glob("test_*.py")):
            module = importlib.import_module(f"tests.{path.stem}")
            classes.update({name: value for name, value in vars(module).items()
                            if isinstance(value, type) and name.endswith("Tests")})
        cited = re.findall(r"`(\w+Tests)\.(test_\w+)`", LEDGER)
        self.assertGreater(len(cited), 40)
        for cls, name in cited:
            with self.subTest(test=f"{cls}.{name}"):
                self.assertIn(cls, classes)
                self.assertTrue(hasattr(classes[cls], name))
        defined = {name for value in classes.values() for name in dir(value)
                   if name.startswith("test_")}
        for name in re.findall(r"`(test_\w+)`", LEDGER):
            with self.subTest(test=name):
                self.assertIn(name, defined)

    def test_honesty_and_check_sections(self):
        limits = next(s for s in trust.DOSSIER if s.id == "limits")
        bullets = next(b for b in limits.blocks if isinstance(b, trust.Bullets))
        self.assertGreaterEqual(len(bullets.items), 8)
        check = next(s for s in trust.DOSSIER if s.id == "check")
        last = check.blocks[-1]
        self.assertIsInstance(last, trust.Note)
        self.assertEqual(last.title, "Still not convinced?")
        self.assertTrue(last.text.startswith("Good."))
        self.assertTrue(any(isinstance(b, trust.LinkBox) for b in check.blocks))

    def test_the_boundary_has_a_described_diagram(self):
        boundary = next(s for s in trust.DOSSIER if s.id == "boundary")
        diagram = [b for b in boundary.blocks if isinstance(b, trust.Diagram)]
        self.assertEqual(len(diagram), 1)
        self.assertIn("dashed", diagram[0].description)


class NoExternalAssetTests(unittest.TestCase):
    """C06, C09: the trust surfaces load nothing and open nothing."""

    def test_trust_files_reference_only_the_further_reading_urls(self):
        allowed = {link.url for link in trust.further_reading()}
        for link in trust.further_reading():
            self.assertTrue(link.url.startswith("https://"), link.url)
        self.assertTrue(about.PROJECT_URL in allowed)
        for name in ("trust.py", "trust_view.py", "theme.py"):
            source = (PROJECT_ROOT / name).read_text(encoding="utf-8")
            for url in re.findall(r"https?://[^\s\"')]+", source):
                with self.subTest(module=name, url=url):
                    self.assertIn(url, allowed)
        self.assertFalse(re.search(r"https?://", _copy()),
                         "an address in the copy outside Further reading")

    def test_trust_view_loads_no_image_or_file(self):
        tree = _tree(PROJECT_ROOT / "trust_view.py")
        self.assertEqual(_imported_roots(tree), {"__future__", "re", "tkinter", "trust", "theme"})
        self.assertEqual(_imported_roots(_tree(PROJECT_ROOT / "trust.py")),
                         {"__future__", "dataclasses", "about", "apppaths", "batch"})
        source = (PROJECT_ROOT / "trust_view.py").read_text(encoding="utf-8")
        for forbidden in ("PhotoImage", "BitmapImage", "iconbitmap", "iconphoto",
                          "open(", "read_text", "read_bytes"):
            self.assertNotIn(forbidden, source)


# ---------------------------------------------------------------------------
# Needing gui.py, and then a display
# ---------------------------------------------------------------------------

@unittest.skipUnless(GUI_IMPORTABLE, "gui.py needs tkinter, which is unavailable")
class GuiWorkerTrustTests(DocxTestCase):
    """C15, C17, C25, C47 — gui.py's workers, no window needed."""

    def test_preview_writes_nothing_beside_the_input(self):
        source = self.build(DOCUMENT)
        before = sorted(p.name for p in self.temp_dir.iterdir())
        self.assertTrue(gui._preview_one(source, self.make_engine(), lambda _: None))
        self.assertEqual(sorted(p.name for p in self.temp_dir.iterdir()), before)

    def test_preview_uses_the_documented_temp_prefix(self):
        prefixes = []
        real = tempfile.mkdtemp

        def recording(*args, **kwargs):
            prefixes.append(kwargs.get("prefix"))
            return real(*args, **kwargs)

        with patch("tempfile.mkdtemp", side_effect=recording):
            gui._preview_one(self.build(DOCUMENT), self.make_engine(), lambda _: None)
        self.assertEqual(set(prefixes), {trust.TEMP_PREFIXES[0], trust.TEMP_PREFIXES[2]})

    def test_log_previews_are_90_characters(self):
        self.assertEqual(gui._shorten("x" * 200), "x" * 90 + "...")
        self.assertEqual(gui._shorten("x" * 90), "x" * 90)
        self.assertIn("shortened to 90 characters", _copy())

    def test_build_engine_rereads_patterns_each_time(self):
        config = self.temp_dir / "patterns.yaml"
        shutil.copyfile(CONFIG_PATH, config)
        with patch.object(gui, "CONFIG_PATH", config):
            first = gui.build_engine()
            config.write_text(config.read_text(encoding="utf-8").replace(
                "formatting_only_removal: false", "formatting_only_removal: true"),
                encoding="utf-8")
            second = gui.build_engine()
        self.assertFalse(first.config["specifier_notes"]["formatting_only_removal"])
        self.assertTrue(second.config["specifier_notes"]["formatting_only_removal"])


def _display_available() -> bool:
    if not GUI_IMPORTABLE:
        return False
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError:
        return False
    root.destroy()
    return True


DISPLAY = _display_available()


@unittest.skipUnless(DISPLAY, "needs tkinter and a display to open windows")
class DialogTests(unittest.TestCase):
    """The dialogs' behaviour: stacking, Escape, focus, the rail."""

    def setUp(self):
        import tkinter as tk
        self.tk = tk
        self.root = tk.Tk()
        self.root.geometry("760x600+0+0")
        self.app = gui.SpecCleanseGUI(self.root)
        self.addCleanup(self._close_root)
        self.root.update()

    def _close_root(self):
        # The log refresh reschedules itself; cancel it so destroying the
        # window does not leave a callback firing into nothing.
        for pending in self.root.tk.splitlist(self.root.tk.call("after", "info")):
            self.root.after_cancel(pending)
        self.root.destroy()

    def _open(self):
        level1 = self.app._show_trust(self.app.btn_trust)
        self.root.update()
        return level1

    def _escape(self, widget):
        widget.focus_force()
        self.root.update()
        widget.event_generate("<Escape>")
        self.root.update()

    def test_level1_button_opens_the_dossier(self):
        level1 = self._open()
        self.assertEqual(level1.button.cget("text"), trust.LEVEL1_BUTTON)
        level1.button.invoke()
        self.root.update()
        self.assertIsNotNone(level1.dossier)
        self.assertTrue(level1.dossier.winfo_exists())
        # Stacked on top, not replacing: Level 1 is still there beneath.
        self.assertTrue(level1.winfo_exists())

    def test_one_escape_closes_one_dialog(self):
        level1 = self._open()
        dossier = level1.open_dossier()
        self.root.update()
        self._escape(dossier.text)
        self.assertTrue(dossier.closed)
        self.assertTrue(level1.winfo_exists())
        self.assertFalse(level1.closed)
        self._escape(level1.text)
        self.assertTrue(level1.closed)
        self.assertTrue(self.root.winfo_exists())

    def test_focus_moves_in_and_returns_to_the_opener(self):
        level1 = self._open()
        self.assertIs(level1.focus_lastfor(), level1.text)
        dossier = level1.open_dossier()
        self.root.update()
        self.assertIs(dossier.focus_lastfor(), dossier.text)
        dossier.close()
        self.root.update()
        self.assertIs(level1.focus_lastfor(), level1.button)
        level1.close()
        self.root.update()
        self.assertIs(self.root.focus_lastfor(), self.app.btn_trust)

    def test_the_top_dialog_holds_the_input(self):
        level1 = self._open()
        self.assertEqual(str(level1.grab_current()), str(level1))
        dossier = level1.open_dossier()
        self.root.update()
        self.assertEqual(str(dossier.grab_current()), str(dossier))
        dossier.close()
        self.root.update()
        self.assertEqual(str(level1.grab_current()), str(level1))

    def test_dossier_is_at_most_88_percent_of_the_screen_height(self):
        import trust_view
        self.assertEqual(trust_view.DOSSIER_MAX_HEIGHT, 0.88)
        dossier = self._open().open_dossier()
        self.root.update()
        limit = dossier.winfo_screenheight() * trust_view.DOSSIER_MAX_HEIGHT
        self.assertLessEqual(dossier.winfo_height(), limit + 1)
        self.assertLessEqual(dossier.maxsize()[1], limit + 1)

    def test_every_section_has_an_anchor_and_a_rail_entry(self):
        dossier = self._open().open_dossier()
        self.root.update()
        self.assertEqual(list(dossier.rail_buttons), DOSSIER_ORDER)
        positions = [dossier.text.index(dossier.mark_for(sid)) for sid in DOSSIER_ORDER]
        self.assertEqual(positions, sorted(positions, key=lambda i: tuple(map(int, i.split(".")))))
        dossier.rail_buttons["runtime"].invoke()
        self.root.update()
        self.assertEqual(dossier.current_section, "runtime")

    def test_the_rail_follows_the_scroll(self):
        dossier = self._open().open_dossier()
        self.root.update()
        dossier.text.yview(dossier.mark_for("boundary"))
        self.root.update()
        self.assertEqual(dossier.current_section, "boundary")
        dossier.text.yview_moveto(0)
        self.root.update()
        self.assertEqual(dossier.current_section, "start")

    def test_every_card_renders_all_five_rows(self):
        dossier = self._open().open_dossier()
        self.root.update()
        content = dossier.text.get("1.0", "end")
        for _, label in trust.CARD_ROWS:
            self.assertEqual(content.count(label + "\t"), len(trust.runtime_cards()), label)

    def test_the_trust_buttons_stay_usable_during_a_run(self):
        self.app._disable_controls()
        self.assertNotIn(self.app.btn_trust, self.app._run_controls())
        self.assertEqual(str(self.app.btn_trust.cget("state")), "normal")

    def test_about_offers_why_trust_it(self):
        self.app._show_about()
        self.root.update()
        labels = [w.cget("text") for w in self._all_widgets(self.root)
                  if isinstance(w, self.tk.Button)]
        self.assertEqual(labels.count("Why trust it?"), 2)

    def test_there_is_no_stop_button(self):
        labels = [w.cget("text") for w in self._all_widgets(self.root)
                  if isinstance(w, self.tk.Button)]
        self.assertTrue(labels)
        self.assertFalse([t for t in labels if re.search(r"stop|cancel", t, re.I)])

    def test_the_window_starts_with_defaults(self):
        self.assertFalse(self.app.strip_revisions.get())
        self.assertIsNone(self.app.output_dir)
        self.assertEqual(self.app.files, [])

    def _all_widgets(self, widget):
        for child in widget.winfo_children():
            yield child
            yield from self._all_widgets(child)


if __name__ == "__main__":
    unittest.main()
