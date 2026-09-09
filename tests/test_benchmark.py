"""The benchmark tool, held to the standard of the code it measures.

A wrong number here is what a performance decision gets taken on, and §16.1
names two specific ways this tool could lie: comparing ZIP bytes instead of
content, and letting a "faster" verification stop reporting damage.
"""

import unittest
from pathlib import Path

from processor import DocxProcessor

from tests import docx_builder as db
from tests.support import DocxTestCase

from tools import benchmark_pipeline as bench


class FixtureFamilyTests(DocxTestCase):
    """Every family §16.1 names exists and builds a document that cleans."""

    def test_every_family_the_plan_names_is_present(self):
        self.assertEqual(sorted(bench.FAMILIES), [
            "adversarial_headings",
            "damaged_output",
            "long_paragraphs",
            "many_placeholders",
            "no_anchors",
            "realistic_requirements",
            "repeated_requirements",
            "tables_and_parts",
        ])

    def test_each_family_builds_a_document_that_cleans_and_verifies(self):
        engine = self.make_engine()
        for family in sorted(bench.FAMILIES):
            with self.subTest(family):
                measurement = bench.measure(family, 40, engine, 1, self.temp_dir)

                self.assertGreater(measurement.input_paragraphs, 0)
                self.assertEqual(len(measurement.clean_seconds), 1)
                self.assertEqual(len(measurement.verify_seconds), 1)

    def test_the_damaged_family_must_not_verify(self):
        # The tripwire that stops a "faster" verification being one that has
        # simply stopped looking.
        measurement = bench.measure("damaged_output", 40, self.make_engine(), 1,
                                    self.temp_dir)

        self.assertFalse(measurement.passed)
        self.assertIn("detected damage", measurement.categories)

    def test_the_adversarial_family_really_does_repeat_itself(self):
        # If this stopped repeating, family 1 would quietly become an easy case
        # and the number it produces would mean nothing.
        body = bench.adversarial_headings(90)
        source = db.build_docx(self.temp_dir / "a.docx", db.document(*body))

        texts = self.paragraph_texts(source)

        self.assertEqual(len(texts), 90)
        self.assertLess(len(set(texts)), 5)

    def test_the_realistic_family_is_mostly_unique(self):
        # The contrast the adversarial family is measured against. §16.4
        # forbids gating on family 1 alone, which only means something if this
        # one is genuinely different.
        body = bench.realistic_requirements(90)
        source = db.build_docx(self.temp_dir / "r.docx", db.document(*body))

        texts = self.paragraph_texts(source)

        self.assertGreater(len(set(texts)), 80)

    def test_the_no_anchor_family_leaves_no_paragraph_unchanged(self):
        # Family 4 is load-bearing precisely because an anchor-based alignment
        # has nothing unchanged to anchor on here.
        engine = self.make_engine()
        body = bench.no_anchors(40)
        source = db.build_docx(self.temp_dir / "n.docx", db.document(*body))
        output = self.temp_dir / "n_out.docx"
        DocxProcessor(engine).process(source, output)

        before = set(self.paragraph_texts(source))
        after = set(self.paragraph_texts(output))

        self.assertTrue(before)
        self.assertFalse(before & after, "some paragraph survived unchanged")

    def test_long_paragraphs_vary_length_not_count(self):
        # Kept separate from paragraph-count scaling on purpose, so the two
        # effects cannot be confused with each other.
        few = bench.long_paragraphs(500)
        many = bench.realistic_requirements(500)

        self.assertLess(len(few), len(many))
        self.assertGreater(len(few[0]), 1000)


class SemanticOracleTests(DocxTestCase):
    """§16.1: ZIP bytes are not the behaviour oracle."""

    def _digest(self, path: Path) -> str:
        return bench.semantic_digest(path, self.make_engine())

    def test_two_packages_with_the_same_text_digest_the_same(self):
        # Repacking changes the archive while the document says the same thing.
        # A byte comparison would call that a behaviour change.
        document = db.document(db.text_para("PART 1 - GENERAL"),
                               db.text_para("Provide listed sprinklers."))
        first = db.build_docx(self.temp_dir / "one.docx", document)
        second = db.build_docx(self.temp_dir / "two.docx", document,
                               {"word/extra.xml": "<x/>"})

        self.assertEqual(self._digest(first), self._digest(second))

    def test_losing_a_paragraph_changes_the_digest(self):
        # The guard against the case above being satisfied by a digest that
        # never changes at all.
        first = db.build_docx(self.temp_dir / "a.docx", db.document(
            db.text_para("PART 1 - GENERAL"),
            db.text_para("Provide listed sprinklers."),
        ))
        second = db.build_docx(self.temp_dir / "b.docx", db.document(
            db.text_para("PART 1 - GENERAL"),
        ))

        self.assertNotEqual(self._digest(first), self._digest(second))

    def test_a_lost_field_carrier_is_counted(self):
        # Text-identical, one field poorer: exactly what a text digest cannot
        # see and why the carrier counts are recorded beside it.
        field = db.para(
            '<w:fldSimple w:instr=" REF Target \\h ">' + db.run("Table 1") + "</w:fldSimple>"
        )
        with_field = db.build_docx(self.temp_dir / "f.docx", db.document(field))
        without = db.build_docx(self.temp_dir / "g.docx",
                                db.document(db.text_para("Table 1")))

        self.assertEqual(self._digest(with_field), self._digest(without))
        self.assertEqual(bench.carrier_counts(with_field)[0], 1)
        self.assertEqual(bench.carrier_counts(without)[0], 0)


class ReportTests(unittest.TestCase):
    """What the report says, and what it refuses to let a reader assume."""

    def _report(self) -> dict:
        return {
            "environment": bench.environment(),
            "repeat": 3,
            "measurements": [
                {
                    "family": "adversarial_headings", "size": 500,
                    "input_paragraphs": 500, "output_paragraphs": 333,
                    "removed": 167, "modified": 0,
                    "clean_seconds": [0.1, 0.2, 0.3],
                    "verify_seconds": [1.0, 2.0, 3.0],
                    "output_digest": "abc", "fields": 0, "references": 0,
                    "passed": True, "categories": [],
                },
            ],
        }

    def test_the_median_is_reported_not_the_mean(self):
        text = bench.format_report(self._report())

        self.assertIn("0.200", text)   # median of 0.1/0.2/0.3
        self.assertIn("2.000", text)   # median of 1.0/2.0/3.0
        self.assertIn("median of 3 unprofiled run(s)", text)

    def test_the_environment_is_named(self):
        # Timings from two machines are not comparable, and a table without
        # this invites exactly that comparison.
        text = bench.format_report(self._report())

        self.assertIn(bench.environment()["python"], text)
        self.assertIn("comparable only within one machine", text)

    def test_the_adversarial_family_is_labelled_as_such(self):
        # §2.4 and §16.1 both insist on it: this is the worst case by
        # construction, and a reader who takes it for the typical one will
        # draw the wrong conclusion about real documents.
        text = bench.format_report(self._report())

        self.assertIn("not the representative one", text)


if __name__ == "__main__":
    unittest.main()
