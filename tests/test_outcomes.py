"""Why a file needs review, and how a finished run says so.

§10.6 criterion 3 requires four categories measured separately and never
pooled, and carries one non-gating requirement into W07: the same split has to
reach the user.  These cases are what stops the four collapsing back into one.

The damage cases build source and output **independently, never by running the
cleaner**.  Agreement between a broken classifier and the cleaner that produced
its input proves nothing.  Each carries an anchor paragraph present on both
sides, so a case cannot pass or fail on invented text instead of on the
classification under test.
"""

import unittest

from batch import FileOutcome, ReviewCategory, describe_categories, summarise
from verify import verify_clean

from tests import docx_builder as db
from tests.support import DocxTestCase


#: Present on both sides of every damage case.  Without it the output differs
#: from the source by an *addition* as well as the injected damage, and the
#: case could pass for the wrong reason.
ANCHOR = db.text_para("PART 1 - GENERAL")

#: One automatic-numbering list, applied by direct w:numPr.
NUMBERED = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'


def numbered(text: str) -> str:
    return "<w:p>" + NUMBERED + db.run(text) + "</w:p>"


class CategoryTests(DocxTestCase):
    """Each kind of concern reaches its own category, and only its own."""

    def categories(self, source: str, output: str) -> set[ReviewCategory]:
        engine = self.make_engine()
        src = db.build_docx(self.temp_dir / "in.docx", source)
        out = db.build_docx(self.temp_dir / "out.docx", output)
        return verify_clean(src, out, engine=engine).review_categories()

    def test_a_correct_clean_is_in_no_category(self):
        # The false-alarm guard.  Without it every assertion below could be
        # satisfied by a classifier that simply distrusts everything.
        self.assertEqual(self.categories(
            db.document(ANCHOR,
                        db.text_para("Retain or delete manufacturers below."),
                        db.text_para("Provide listed sprinklers.")),
            db.document(ANCHOR, db.text_para("Provide listed sprinklers.")),
        ), set())

    def test_a_deleted_requirement_is_detected_damage(self):
        self.assertEqual(self.categories(
            db.document(ANCHOR, db.text_para("Provide listed sprinklers.")),
            db.document(ANCHOR),
        ), {ReviewCategory.DETECTED_DAMAGE})

    def test_an_invented_paragraph_is_detected_damage(self):
        self.assertEqual(self.categories(
            db.document(ANCHOR),
            db.document(ANCHOR, db.text_para("Provide gold-plated sprinklers.")),
        ), {ReviewCategory.DETECTED_DAMAGE})

    def test_a_loss_paired_only_by_similarity_is_ambiguous_not_damage(self):
        # The pairing here is the MIN_PAIR_SIMILARITY guess, and the fragments
        # it reports are artefacts of where the differ happened to align —
        # 'nd hangers a' is not text anyone edited out.  Reporting that as
        # detected damage would claim more than the comparison established.
        categories = self.categories(
            db.document(
                ANCHOR,
                db.text_para("Provide sprinkler piping and hangers as indicated on drawings."),
            ),
            db.document(ANCHOR, db.text_para("Provide sprinkler piping as indicated.")),
        )

        self.assertEqual(categories, {ReviewCategory.AMBIGUOUS_ALIGNMENT})

    def test_numbering_is_its_own_category_not_damage(self):
        # Nothing was lost that policy did not authorize; what may have changed
        # is the numbers a reader sees.
        self.assertEqual(self.categories(
            db.document(ANCHOR,
                        numbered("Retain or delete manufacturers below."),
                        numbered("Provide listed sprinklers.")),
            db.document(ANCHOR, numbered("Provide listed sprinklers.")),
        ), {ReviewCategory.REFERENCE_NUMBERING})

    def test_a_broken_reference_is_a_warning_and_the_loss_is_damage(self):
        # Two distinct concerns in one file: the bookmark's paragraph really is
        # gone, and a live REF now names nothing.  Both are reported, because
        # choosing one to show is the pooling this split exists to end.
        field = db.para(
            '<w:fldSimple w:instr=" REF TargetA \\h ">' + db.run("Table 1") + "</w:fldSimple>"
        )
        target = db.para(
            db.bookmark_start("1", "TargetA"), db.run("Table 1"), db.bookmark_end("1")
        )

        self.assertEqual(self.categories(
            db.document(ANCHOR, target, field),
            db.document(ANCHOR, field),
        ), {ReviewCategory.DETECTED_DAMAGE, ReviewCategory.REFERENCE_NUMBERING})


class ViolationKindTests(DocxTestCase):
    """A category is decided by a field, never by reading the message."""

    def test_a_broken_reference_is_kind_reference(self):
        engine = self.make_engine()
        field = db.para(
            '<w:fldSimple w:instr=" REF TargetA \\h ">' + db.run("Table 1") + "</w:fldSimple>"
        )
        src = db.build_docx(self.temp_dir / "in.docx", db.document(
            ANCHOR,
            db.para(db.bookmark_start("1", "TargetA"), db.run("Table 1"), db.bookmark_end("1")),
            field,
        ))
        out = db.build_docx(self.temp_dir / "out.docx", db.document(ANCHOR, field))

        result = verify_clean(src, out, engine=engine)

        self.assertEqual([v.kind for v in result.reference_violations], ["reference"])
        self.assertEqual(result.structural_damage, [])


class SummaryTests(unittest.TestCase):
    """What a finished run says, without opening a window."""

    def test_needs_review_names_its_categories(self):
        line = summarise(
            {FileOutcome.VERIFIED: 8, FileOutcome.NEEDS_REVIEW: 2, FileOutcome.FAILED: 1},
            {ReviewCategory.DETECTED_DAMAGE: 1, ReviewCategory.REFERENCE_NUMBERING: 1},
        )

        self.assertEqual(
            line,
            "Done: 8 verified, 2 need review "
            "(1 detected damage, 1 reference/numbering warning), 1 failed",
        )

    def test_a_run_with_nothing_to_report_is_unchanged(self):
        self.assertEqual(
            summarise({FileOutcome.VERIFIED: 3}), "Done: 3 verified"
        )

    def test_category_counts_may_exceed_the_file_count(self):
        # One file in two categories is counted in both.  Forcing one category
        # per file would mean choosing which real concern to hide.
        line = summarise(
            {FileOutcome.NEEDS_REVIEW: 1},
            {ReviewCategory.DETECTED_DAMAGE: 1, ReviewCategory.REFERENCE_NUMBERING: 1},
        )

        self.assertIn("1 needs review (1 detected damage, 1 reference/numbering warning)", line)

    def test_a_failure_says_whether_a_file_was_written(self):
        # §14.1: "Failed; state whether an output exists".  Without this a
        # reader cannot tell there is something in the output folder to delete.
        self.assertEqual(
            summarise({FileOutcome.FAILED: 2}, unverified_outputs=1),
            "Done: 2 failed (1 wrote an unverified file)",
        )

    def test_a_failure_that_wrote_nothing_says_only_that(self):
        self.assertEqual(summarise({FileOutcome.FAILED: 2}), "Done: 2 failed")

    def test_categories_are_listed_in_declaration_order(self):
        # Not by frequency: a summary whose wording reshuffles between runs is
        # harder to read, and there is no severity ranking to sort by.
        self.assertEqual(
            describe_categories({
                ReviewCategory.REFERENCE_NUMBERING: 5,
                ReviewCategory.AMBIGUOUS_ALIGNMENT: 1,
            }),
            "1 ambiguous alignment, 5 reference/numbering warning",
        )

    def test_an_empty_category_map_adds_nothing(self):
        self.assertEqual(summarise({FileOutcome.NEEDS_REVIEW: 1}, {}), "Done: 1 needs review")


if __name__ == "__main__":
    unittest.main()
