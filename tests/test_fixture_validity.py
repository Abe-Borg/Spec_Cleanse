"""The fixtures themselves, checked as documents rather than as strings.

§17.1: "Existing minimalist fixtures are not automatically valid Word
compatibility fixtures — for example, verify that every namespace named by
`mc:Ignorable` is actually declared."  They were not: `mc:Ignorable` named
`w14` and nothing declared it, which makes a part non-conformant under Markup
Compatibility and useless as a positive control — Word rejecting it would say
nothing about the change under test.

These are cheap invariants about the builder, kept because every other test in
the suite rests on it. A fixture that is wrong in a way no test looks at makes
every assertion over it worth less than it appears.
"""

import unittest
import zipfile
from pathlib import Path

from lxml import etree

from docx_xml import NAMESPACES, W, collect_content_parts, parse_xml

from tests import docx_builder as db
from tests.support import DocxTestCase

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


class NamespaceDeclarationTests(DocxTestCase):
    """Every prefix a part names must be one the part declares."""

    def _parts(self, path: Path) -> list[tuple[str, etree._Element]]:
        parsed = []
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if not name.endswith(".xml"):
                    continue
                try:
                    parsed.append((name, etree.fromstring(archive.read(name))))
                except etree.XMLSyntaxError as exc:  # pragma: no cover
                    self.fail(f"{name} is not well-formed XML: {exc}")
        return parsed

    def test_every_ignorable_prefix_is_declared(self):
        # The defect §17.1 names, as an assertion rather than an inspection.
        path = self.build(db.document(
            db.text_para("PART 1 - GENERAL"),
            db.text_para("Provide listed sprinklers."),
        ), {"word/header1.xml": db.header(db.text_para("Header"))})

        for name, root in self._parts(path):
            ignorable = root.get(f"{{{MC_NS}}}Ignorable")
            if not ignorable:
                continue
            with self.subTest(name):
                undeclared = [
                    prefix for prefix in ignorable.split()
                    if prefix not in root.nsmap
                ]
                self.assertEqual(
                    undeclared, [],
                    f"{name}: mc:Ignorable names prefixes nothing declares",
                )

    def test_the_builder_declares_what_it_marks_ignorable(self):
        # Guards the declaration string directly, so a future prefix added to
        # mc:Ignorable cannot be forgotten even if no fixture happens to use it.
        declared = {
            token.split("=")[0].removeprefix("xmlns:")
            for token in db.NS_DECL.split()
            if token.startswith("xmlns:")
        }
        marked = db.NS_DECL.split('mc:Ignorable="')[1].split('"')[0].split()

        self.assertTrue(marked, "nothing is marked ignorable; check the fixture")
        self.assertTrue(set(marked) <= declared, f"{marked} vs {sorted(declared)}")

    def test_generated_parts_parse_with_the_project_namespace_map(self):
        # The map the application uses must resolve against what the builder
        # emits; a fixture the code cannot address is not exercising anything.
        path = self.build(db.document(db.text_para("Provide listed sprinklers.")))

        with zipfile.ZipFile(path) as archive:
            root = etree.fromstring(archive.read("word/document.xml"))

        self.assertTrue(root.xpath("//w:p", namespaces=NAMESPACES))


class PackageShapeTests(DocxTestCase):
    """A generated package has the parts the application expects to walk."""

    def test_the_required_package_parts_are_present(self):
        path = self.build(db.document(db.text_para("Provide listed sprinklers.")))

        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())

        for required in ("[Content_Types].xml", "_rels/.rels",
                         "word/document.xml", "word/_rels/document.xml.rels"):
            self.assertIn(required, names)

    def test_extra_parts_are_reached_by_the_part_collector(self):
        # processor.py and verify.py both walk collect_content_parts, so a
        # fixture whose header it cannot see is silently untested.
        path = self.build(
            db.document(db.text_para("Body.")),
            {
                "word/header1.xml": db.header(db.text_para("Header.")),
                "word/footnotes.xml": db.footnotes(db.text_para("Note.")),
            },
        )
        unpacked = self.temp_dir / "unpacked"
        with zipfile.ZipFile(path) as archive:
            archive.extractall(unpacked)

        found = {p.name for p in collect_content_parts(unpacked / "word")}

        self.assertEqual(found, {"document.xml", "header1.xml", "footnotes.xml"})

    def test_a_bookmark_fixture_defines_both_ends(self):
        # A half-open range is a real case the cleaner handles, but a fixture
        # that means to define a whole one and does not is testing something
        # else by accident.
        path = self.build(db.document(db.para(
            db.bookmark_start("1", "Target"),
            db.run("Table 1"),
            db.bookmark_end("1"),
        )))
        unpacked = self.temp_dir / "unpacked"
        with zipfile.ZipFile(path) as archive:
            archive.extractall(unpacked)
        root = parse_xml(unpacked / "word" / "document.xml").getroot()

        starts = {e.get(f"{W}id") for e in root.iter(f"{W}bookmarkStart")}
        ends = {e.get(f"{W}id") for e in root.iter(f"{W}bookmarkEnd")}

        self.assertEqual(starts, ends)


if __name__ == "__main__":
    unittest.main()
