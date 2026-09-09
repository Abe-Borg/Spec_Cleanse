"""What the shipped patterns.yaml does to representative specification prose.

The corpus in section 8.2 of the implementation plan, pinned as tests. Two
jobs:

*Negatives* are plausible requirement text that must survive a clean. Five of
them do not today — they are the reproduced F01 and F02 false positives — and
are carried under ``unittest.expectedFailure`` until W02 narrows the rules.
The suite stays green while they stand and turns red the moment one starts
surviving, which is the signal to drop its decorator.

*Positives* are unambiguous editorial content that must keep being cleaned.
They all pass today, and they exist to catch W02 narrowing the rules too far.
Reduced editorial recall is an accepted cost of the F01/F02 fix; losing these
is not.

These are synthetic policy fixtures. They say what the shipped defaults do to
these exact strings, and nothing about how often such strings occur in a real
specification — that is what the censuses in ``tools/`` are for.
"""

import unittest

from tests import docx_builder as db
from tests.support import DocxTestCase


#: Closes the five expected failures below.
NARROWING_PACKAGE = "W02"


class ShippedPolicyTestCase(DocxTestCase):
    """Runs one paragraph through a real clean and reports what survived."""

    def survives(self, sentence: str) -> bool:
        path = self.build(db.document(db.text_para(sentence)))
        _, out = self.clean(path)
        return self.paragraph_texts(out) == [sentence]

    def assertSurvives(self, sentence: str):
        self.assertTrue(
            self.survives(sentence),
            f"requirement text was not retained: {sentence!r}",
        )

    def assertCleaned(self, sentence: str):
        self.assertFalse(
            self.survives(sentence),
            f"editorial text was not cleaned: {sentence!r}",
        )


class CopyrightNegatives(ShippedPolicyTestCase):
    """Requirement prose that the copyright rules currently claim.

    Each of these matches one unanchored phrase, and `CopyrightDetector`
    scores a single match at 0.7 — over the 0.5 threshold on its own. Removing
    the global DOTALL flag fixes none of them: all three match on one line.
    """

    @unittest.expectedFailure
    def test_shop_drawings_reproduction_clause(self):
        """EXPECTED TO FAIL until W02 — matches `may\\s+not\\s+be\\s+reproduced`."""
        self.assertSurvives(
            "Shop Drawings submitted under this Section may not be reproduced "
            "for use on other projects."
        )

    @unittest.expectedFailure
    def test_duplicate_sprinkler_coverage_clause(self):
        """EXPECTED TO FAIL until W02 — matches `duplication.*?prohibited`."""
        self.assertSurvives(
            "Contractor shall verify that duplication of sprinkler coverage in "
            "adjacent zones is prohibited by the AHJ."
        )

    @unittest.expectedFailure
    def test_fire_pump_room_access_clause(self):
        """EXPECTED TO FAIL until W02 — matches `unauthorized.*?reproduction`."""
        self.assertSurvives(
            "Unauthorized personnel shall not have access to the fire pump "
            "room; reproduction of access keys is not permitted."
        )


class EditorialOverreachNegatives(ShippedPolicyTestCase):
    """Requirement prose that the high-confidence editorial tier currently claims."""

    @unittest.expectedFailure
    def test_selection_instruction_to_the_contractor(self):
        """EXPECTED TO FAIL until W02 — matches `^\\s*(?:select|choose)\\s+one\\b`.

        Ambiguous: it reads as an instruction to a specifier and as a
        requirement on a contractor. The plan retains it by default; retaining
        editorial noise is cheaper than deleting a requirement.
        """
        self.assertSurvives("Select one of the listed manufacturers.")

    @unittest.expectedFailure
    def test_requirement_containing_a_bracketed_editorial_marker(self):
        """EXPECTED TO FAIL until W02 — `retain\\s+or\\s+delete` takes the paragraph.

        A whole-paragraph rule fires on a marker embedded in a requirement, so
        the requirement goes with it. The selected behaviour is to retain the
        paragraph complete, brackets and all, rather than invent an inline rule
        to cut them: the marker reaching downstream analysis is the lesser cost.
        """
        self.assertSurvives("Provide two [retain or delete] spare filters per unit.")


class NegativesAlreadyHonoured(ShippedPolicyTestCase):
    """Requirement prose the shipped rules already leave alone.

    Pinned so a W02 narrowing cannot regress them, and so the anchoring that
    protects them is not removed by accident.
    """

    def test_low_confidence_prose_without_editorial_formatting(self):
        # `revise as required` is in the low-confidence tier, which needs a
        # formatting signal to cross the threshold.  Plain text does not.
        self.assertSurvives("Provide pumps and revise as required.")

    def test_retain_that_does_not_point_above_or_below(self):
        # The MasterSpec rule is anchored and must point somewhere in the same
        # sentence, so it cannot fire on a records-retention requirement.
        self.assertSurvives("Retain records of all tests required in Paragraph 1.6.")

    def test_select_one_piece_is_not_select_one(self):
        # The `(?!-)` lookahead is what keeps this a requirement.
        self.assertSurvives("Select one-piece molded fittings for changes in direction.")


class ShippedPolicyPositives(ShippedPolicyTestCase):
    """Editorial content that must keep being cleaned after W02 narrows the rules."""

    def test_copyright_notice(self):
        self.assertCleaned("© 2026 ARCOM. All rights reserved.")

    def test_delimited_specifier_note(self):
        self.assertCleaned("[Specifier: delete this note before issue]")

    def test_retain_instruction_pointing_below(self):
        self.assertCleaned("Retain subparagraph below for wet-pipe systems.")

    def test_copy_instruction_pointing_above(self):
        self.assertCleaned("Copy paragraphs above for each additional riser.")

    def test_inline_placeholder_is_cut_and_the_requirement_kept(self):
        path = self.build(db.document(db.text_para(
            "Provide two [Verify quantity with Owner] spare sprinklers."
        )))
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), ["Provide two spare sprinklers."])


if __name__ == "__main__":
    unittest.main()
