"""Source evidence must agree with what the processor actually does.

``DetectionEngine.paragraph_evidence()`` states the policy verification judges
against.  If it drifts from the cleaner, verification stops being a check and
becomes a second opinion about a different program — the exact failure the old
flattened ``(category, regex)`` list produced.

So every case here is cleaned for real, and the evidence read from the *source*
is compared with the *output*: a paragraph the evidence authorizes losing must
actually be gone, and one it does not must still be there.
"""

import difflib
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

from detection import ParagraphEvidence, RunEvidence, WHITESPACE_ONLY
from docx_xml import (
    iter_paragraphs,
    paragraph_signature,
    run_profile,
    load_styles,
    paragraph_text,
    parse_xml,
    spans_cover,
)

from tests import docx_builder as db
from tests.support import DocxTestCase


STYLES = db.styles(
    db.style_def("CMT", name="CMT"),
    db.style_def("ART", name="ART"),
    db.style_def("SpecifierNote", name="Specifier Note", style_type="character"),
    db.style_def("Hidden", name="Hidden", hidden=True, style_type="character"),
)

#: Each case is a paragraph plus what the cleaner is expected to do with it.
#: The evidence is not consulted to build this list — it is the independent
#: statement of intent both sides are measured against.
CASES = [
    ("plain requirement",
     db.text_para("Provide fire pumps rated 1500 gpm at 100 psi."), "keep"),
    ("specifier note",
     db.text_para("Note to Specifier: delete before issue."), "remove"),
    ("copyright marker",
     db.text_para("Copyright 2025 by the American Institute of Architects."), "remove"),
    ("editorial paragraph style",
     db.text_para("Coordinate hangers with structural.", style="CMT"), "remove"),
    ("preserve style outranks a matching rule",
     db.text_para("Retain or delete manufacturers below.", style="ART"), "keep"),
    ("preserve pattern",
     db.text_para("PART 1 - GENERAL"), "keep"),
    ("hidden run beside a requirement",
     db.para(db.run("Provide listed devices. "),
             db.run("Delete before issue.", vanish=True)), "trim"),
    ("every run hidden",
     db.para(db.run("Delete this. ", vanish=True),
             db.run("And this.", vanish=True)), "remove"),
    ("low-confidence phrase, no formatting",
     db.text_para("Provide pumps and revise as required."), "keep"),
    ("low-confidence phrase in an editorial style",
     db.text_para("Provide pumps and revise as required.", style="CMT"), "remove"),
    ("inline placeholder",
     db.text_para("Provide two [Verify quantity] spare filters."), "trim"),
    ("placeholder is the whole paragraph",
     db.text_para("[Insert manufacturer]"), "remove"),
    ("placeholder straddling a page break",
     db.para(db.run("Provide [Verify", inner='<w:br w:type="page"/>'),
             db.run(" quantity] valves.")), "keep"),
]


class EvidenceCase(DocxTestCase):
    """Shared unpacking so a case reads as source evidence versus output."""

    def _paragraphs(self, path):
        temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_evidence_"))
        self.addCleanup(shutil.rmtree, temp_dir, True)
        with zipfile.ZipFile(path) as zf:
            zf.extractall(temp_dir)
        root = parse_xml(temp_dir / "word" / "document.xml").getroot()
        return temp_dir, [
            para for para in iter_paragraphs(root) if paragraph_text(para).strip()
        ]

    def _evidence_for(self, path, engine):
        """Read every source paragraph's evidence, in document order."""
        temp_dir, paragraphs = self._paragraphs(path)
        engine.bind_styles(load_styles(temp_dir / "word"))
        return [(paragraph_text(p), engine.paragraph_evidence(p)) for p in paragraphs]

    def _output_texts(self, path):
        _, paragraphs = self._paragraphs(path)
        return [paragraph_text(p) for p in paragraphs]


class EvidenceMatchesProcessor(EvidenceCase):
    """The evidence read from the source predicts the actual output."""

    def test_each_case(self):
        for label, xml, expected in CASES:
            with self.subTest(label):
                engine = self.make_engine()
                src = self.build(db.document(xml), {"word/styles.xml": STYLES},
                                 name=f"{abs(hash(label))}.docx")
                _, out = self.clean(src, engine=engine)

                (source_text, evidence), = self._evidence_for(src, engine)
                survivors = self._output_texts(out)
                gone = not survivors

                if expected == "remove":
                    self.assertTrue(
                        gone, f"{label}: expected the paragraph to be removed")
                    self.assertTrue(
                        evidence.authorizes_whole_paragraph(),
                        f"{label}: it was removed, but the evidence does not "
                        "authorize losing the paragraph")
                else:
                    self.assertFalse(
                        gone, f"{label}: the paragraph should have survived")
                    if expected == "keep":
                        self.assertEqual(
                            survivors, [source_text],
                            f"{label}: the text should be untouched")
                    else:
                        self.assertNotEqual(
                            survivors, [source_text],
                            f"{label}: the paragraph should have lost text")
                    self.assertFalse(
                        evidence.authorizes_whole_paragraph(),
                        f"{label}: it survived, but the evidence says the whole "
                        "paragraph could go")

    def test_trimmed_text_lies_inside_authorized_spans(self):
        """What a surviving paragraph actually lost was authorized, positionally."""
        for label, xml, expected in CASES:
            if expected != "trim":
                continue
            with self.subTest(label):
                engine = self.make_engine()
                src = self.build(db.document(xml), {"word/styles.xml": STYLES},
                                 name=f"trim{abs(hash(label))}.docx")
                _, out = self.clean(src, engine=engine)
                (source_text, evidence), = self._evidence_for(src, engine)
                after, = self._output_texts(out)

                spans = evidence.authorized_spans()
                self.assertTrue(spans, f"{label}: no authorized interval")

                matcher = difflib.SequenceMatcher(
                    None, source_text, after, autojunk=False)
                for tag, i1, i2, _j1, _j2 in matcher.get_opcodes():
                    if tag == "equal":
                        continue
                    self.assertEqual(tag, "delete", f"{label}: text was mutated")
                    for idx in range(i1, i2):
                        if source_text[idx].strip():
                            self.assertTrue(
                                spans_cover(spans, idx, idx + 1),
                                f"{label}: lost character {source_text[idx]!r} at "
                                f"{idx} is outside every authorized interval")


class SpanAuthorityShape(unittest.TestCase):
    """Every answer from ``reason_for_span`` has the same shape.

    A whitespace-only interval is filtered out before verification asks, so this
    path is not reachable through ``verify_clean`` today.  It is asserted anyway
    because the cost of it being wrong is an unpacking error in the middle of a
    verdict, and the filter that hides it is one refactor away from moving.
    """

    def _evidence(self):
        return ParagraphEvidence(
            raw_text="Provide  valves.",
            runs=(RunEvidence(0, 16, "Provide  valves.", "hidden_text", "Hidden text"),),
        )

    def test_every_authority_is_a_triple(self):
        evidence = self._evidence()
        for start, end in ((0, 7), (7, 9), (0, 16)):
            with self.subTest(span=(start, end)):
                authority = evidence.reason_for_span(start, end)
                self.assertIsNotNone(authority)
                self.assertEqual(len(authority), 3, authority)
                category, reason, formatting_only = authority
                self.assertIsInstance(category, str)
                self.assertIsInstance(reason, str)
                self.assertIsInstance(formatting_only, bool)

    def test_a_whitespace_only_span_is_labelled_as_such(self):
        evidence = ParagraphEvidence(raw_text="Provide  valves.")
        self.assertEqual(
            evidence.reason_for_span(7, 9),
            (WHITESPACE_ONLY, "adjacent whitespace", False),
        )

    def test_an_unauthorized_character_answers_none(self):
        evidence = ParagraphEvidence(raw_text="Provide  valves.")
        self.assertIsNone(evidence.reason_for_span(0, 7))

    def test_unreliable_offsets_authorize_nothing(self):
        """Rather than guess: if the offsets do not add up, nothing is permitted."""
        evidence = ParagraphEvidence(
            raw_text="Provide valves.",
            runs=(RunEvidence(0, 15, "Provide valves.", "hidden_text", "Hidden text"),),
            offsets_reliable=False,
        )
        self.assertIsNone(evidence.reason_for_span(0, 7))
        self.assertEqual(evidence.authorized_spans(), [])
        self.assertFalse(evidence.covers_all_text())
        self.assertFalse(evidence.authorizes_whole_paragraph())


class RunSignatures(EvidenceCase):
    """Signatures say whether two runs are interchangeable, nothing more."""

    def _paragraph(self, xml, name="sig"):
        src = self.build(db.document(xml), {"word/styles.xml": STYLES},
                         name=f"{name}{abs(hash(xml))}.docx")
        _, paragraphs = self._paragraphs(src)
        return paragraphs[0]

    def _signatures(self, xml):
        return run_profile(self._paragraph(xml))

    def test_identical_text_with_different_formatting_is_distinguishable(self):
        """The whole point: same characters, different run, different signature."""
        first, second = self._signatures(
            db.para(db.run("Keep. ", vanish=True), db.run("Keep. ")))

        self.assertEqual(first[0], second[0], "the text really is identical")
        self.assertNotEqual(first, second, "hidden must separate them")

    def test_identical_runs_have_identical_signatures(self):
        first, second = self._signatures(db.para(db.run("Keep. "), db.run("Keep. ")))

        self.assertEqual(first, second)

    def test_style_colour_italic_and_bold_all_separate_runs(self):
        for label, run in (
            ("style", db.run("X", rstyle="SpecifierNote")),
            ("colour", db.run("X", color="FF0000")),
            ("italic", db.run("X", italic=True)),
            ("bold", db.run("X", bold=True)),
        ):
            with self.subTest(label):
                plain, marked = self._signatures(db.para(db.run("X"), run))
                self.assertNotEqual(plain, marked, f"{label} must separate runs")

    def test_runs_without_substantive_text_are_left_out(self):
        """An empty run carries structure, not content."""
        signatures = self._signatures(
            db.para(db.run("Real text."), db.run(""), db.run("   ")))

        self.assertEqual(len(signatures), 1)

    def test_a_paragraph_signature_carries_the_style_the_runs_do_not(self):
        """Identical runs, different authority — the style has to be in the identity.

        An editorial *paragraph* style permits losing the paragraph while its
        runs say nothing at all.  Without the style here, removing the styled
        copy of a repeated text was blamed on the plain one, and a correct clean
        was reported as damage.
        """
        text = "Coordinate hangers with the structural drawings."
        plain = paragraph_signature(self._paragraph(db.text_para(text), "plain"))
        styled = paragraph_signature(
            self._paragraph(db.text_para(text, style="CMT"), "styled"))

        self.assertEqual(plain[1], styled[1], "the runs really are identical")
        self.assertNotEqual(plain, styled, "the style must separate them")
        self.assertEqual(styled[0], "CMT")


class OffsetIntegrity(EvidenceCase):
    """Run offsets are offsets into the document, not into a tidied copy."""

    def test_offsets_reconstruct_the_paragraph(self):
        xml = db.para(
            db.run("Hangers", inner="<w:tab/>"),
            db.hyperlink(db.run("Section 21 05 29")),
            db.run(" and "),
            db.inserted(db.run("supplementary")),
            db.run(" seismic bracing."),
        )
        engine = self.make_engine()
        src = self.build(db.document(xml), {"word/styles.xml": STYLES})
        (_, evidence), = self._evidence_for(src, engine)

        self.assertTrue(evidence.offsets_reliable)
        self.assertEqual(
            "".join(r.text for r in evidence.runs), evidence.raw_text,
            "run texts must reconstruct the paragraph exactly")
        for run in evidence.runs:
            self.assertEqual(
                evidence.raw_text[run.start:run.end], run.text,
                "a run's offsets must address its own text")

    def test_leading_whitespace_is_not_silently_dropped(self):
        """Offsets address the raw text, which keeps its leading whitespace."""
        engine = self.make_engine()
        src = self.build(
            db.document(db.para(db.run("   Provide valves."))),
            {"word/styles.xml": STYLES})
        (_, evidence), = self._evidence_for(src, engine)
        self.assertTrue(evidence.raw_text.startswith("   "))
        self.assertEqual(evidence.runs[0].start, 0)


if __name__ == "__main__":
    unittest.main()
