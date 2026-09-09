"""What the shipped patterns.yaml does to representative specification prose.

The corpus in section 8.2 of the implementation plan, pinned as tests. Two
jobs:

*Negatives* are plausible requirement text that must survive a clean. Five of
them did not before W02 — they were the reproduced F01 and F02 false positives
— and were carried under ``unittest.expectedFailure`` until the rules were
narrowed. Narrowing them turned the suite red with five unexpected successes,
which was the signal to drop the decorators; they are ordinary passing tests
now, and they are the regression guard against the old breadth coming back.

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

    Each matched one unanchored phrase, and `CopyrightDetector` scores a
    single match at 0.7 — over the 0.5 threshold on its own. Removing the
    global DOTALL flag would have fixed none of them: all three match on one
    line. W02 narrowed the phrases instead.
    """

    def test_shop_drawings_reproduction_clause(self):
        """Was taken by the unbounded `may not be reproduced`.

        The narrowed rule requires the clause a notice actually uses — "in
        whole or in part", "without written permission" — which this has not.
        """
        self.assertSurvives(
            "Shop Drawings submitted under this Section may not be reproduced "
            "for use on other projects."
        )

    def test_duplicate_sprinkler_coverage_clause(self):
        """Was taken by `duplication.*?prohibited` matching across the sentence.

        Four words separate "duplication" from "is prohibited" here; the
        narrowed rule allows at most two, so the words have to be adjacent the
        way they are in a notice.
        """
        self.assertSurvives(
            "Contractor shall verify that duplication of sprinkler coverage in "
            "adjacent zones is prohibited by the AHJ."
        )

    def test_fire_pump_room_access_clause(self):
        """Was taken by `unauthorized.*?reproduction` spanning a semicolon.

        "Unauthorized" opens one clause and "reproduction" opens another. The
        narrowed rule requires them adjacent.
        """
        self.assertSurvives(
            "Unauthorized personnel shall not have access to the fire pump "
            "room; reproduction of access keys is not permitted."
        )


class CopyingLanguageNegatives(ShippedPolicyTestCase):
    """Restrictions on copying that are requirements, not copyright notices.

    Reproduction and duplication language says the same thing, about the same
    act, in a notice and in a requirement. What differs is who is imposing the
    restriction, and that is not in the text. Three narrowed rules each looked
    safe against the F01 examples and then took one of these; they were removed
    rather than narrowed further, and these are the cases that decided it.

    Door hardware key control and submittal restrictions appear in essentially
    every nonresidential project, so a rule that eats them is not a corner case.
    """

    def test_keys_designed_to_prevent_unauthorized_duplication(self):
        # Took the narrowed `unauthorized (\w+ ){0,2}(reproduction|duplication)`.
        self.assertSurvives("Provide keys designed to prevent unauthorized duplication.")

    def test_cylinders_that_prevent_unauthorized_key_duplication(self):
        self.assertSurvives(
            "Cylinders shall be of a type that prevents unauthorized key duplication."
        )

    def test_patented_keyways_restricting_duplication(self):
        self.assertSurvives(
            "Furnish patented keyways to restrict unauthorized duplication."
        )

    def test_key_duplication_prohibited_without_authorization(self):
        # Took the narrowed `duplication (\w+ ){0,2}is (strictly )?prohibited`.
        self.assertSurvives(
            "Duplication of keys is prohibited without written authorization "
            "from the Owner."
        )

    def test_shop_drawings_not_reproduced_in_whole_or_in_part(self):
        # Took the narrowed `may not be reproduced (in whole|in part|without
        # written)`.  A submittal restriction is boilerplate too, and reaches
        # for the same words a notice does.
        self.assertSurvives(
            "Shop Drawings may not be reproduced in whole or in part without "
            "the Architect's written consent."
        )

    def test_record_drawings_not_reproduced_without_approval(self):
        self.assertSurvives(
            "Record Drawings may not be reproduced without written approval "
            "of the Owner."
        )

    def test_software_licensed_for_use_by_the_owner(self):
        # `licensed for use by` was unanchored and took this.  A licence line
        # opens with the phrase; a requirement embeds it mid-sentence.
        self.assertSurvives(
            "Software licensed for use by the Owner shall be transferable."
        )


class EditorialOverreachNegatives(ShippedPolicyTestCase):
    """Requirement prose that the high-confidence editorial tier currently claims."""

    def test_selection_instruction_to_the_contractor(self):
        """Was taken by the bare `^select one` anchor.

        The opening words are genuinely ambiguous — a specifier and a
        contractor are told to select one of something in the same voice. What
        separates them is the referent: an editorial note points at the
        document's own structure ("of the following paragraphs", "below"), and
        this points at manufacturers. The narrowed rule requires that referent.
        """
        self.assertSurvives("Select one of the listed manufacturers.")

    def test_requirement_containing_a_bracketed_editorial_marker(self):
        """Was taken by an unanchored `retain or delete` inside a requirement.

        The rule is now anchored to the start of the paragraph, so it fires
        when the paragraph *is* the instruction rather than when a requirement
        contains the marker. The paragraph is retained complete, brackets and
        all: cutting the marker would need an inline rule, and one broad enough
        to catch it would put every bracketed phrase in a requirement at risk.
        The marker reaching downstream analysis is the lesser cost.
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

    def test_a_licence_line_opening_with_the_phrase(self):
        self.assertCleaned("Licensed for use by a single user.")

    def test_the_arcom_distribution_notice(self):
        self.assertCleaned(
            "This document is exclusively published and distributed by ARCOM."
        )

    def test_a_dated_copyright_line(self):
        self.assertCleaned(
            "Copyright 2026 by the American Institute of Architects. "
            "All rights reserved."
        )

    def test_inline_placeholder_is_cut_and_the_requirement_kept(self):
        path = self.build(db.document(db.text_para(
            "Provide two [Verify quantity with Owner] spare sprinklers."
        )))
        _, out = self.clean(path)

        self.assertEqual(self.paragraph_texts(out), ["Provide two spare sprinklers."])


if __name__ == "__main__":
    unittest.main()
