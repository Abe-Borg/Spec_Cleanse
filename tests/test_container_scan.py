"""`can_delete_paragraph` decides the same way after §16.3's rewrite.

The old version materialised the parent's whole block-child list on every
call, which made a body of n paragraphs losing k of them do O(k*n) work — 56%
of cleaning time at 4,000 paragraphs.  Scanning until the answer is known is
faster; these cases are here because it must also be the *same* answer.

§16.3 enumerates what has to keep working, and every item is below.  The tests
are on the predicate directly rather than through a clean, so a case cannot
pass because some other rule happened to keep the paragraph.
"""

import itertools
import unittest

from lxml import etree

from docx_xml import (
    BLOCK_CONTAINERS,
    P_TAG,
    TC_TAG,
    W,
    block_children,
    can_delete_paragraph,
)

NS = f'xmlns:w="{W[1:-1]}"'


def parse(xml: str) -> etree._Element:
    return etree.fromstring(xml.encode("utf-8"))


def paragraphs(container: etree._Element) -> list[etree._Element]:
    return [child for child in container if child.tag == f"{W}p"]


class ContainerScanTests(unittest.TestCase):

    def test_a_sole_paragraph_in_a_body_may_not_be_deleted(self):
        body = parse(f"<w:body {NS}><w:p/></w:body>")

        self.assertFalse(can_delete_paragraph(paragraphs(body)[0]))

    def test_one_of_several_may_be(self):
        body = parse(f"<w:body {NS}><w:p/><w:p/></w:body>")

        self.assertTrue(can_delete_paragraph(paragraphs(body)[0]))

    def test_several_removed_consecutively_stop_at_the_last(self):
        # The scan reads the live tree, so each answer reflects the removals
        # already made — the reason nothing is cached across mutations.
        body = parse(f"<w:body {NS}><w:p/><w:p/><w:p/></w:body>")

        removed = 0
        for para in list(paragraphs(body)):
            if can_delete_paragraph(para):
                body.remove(para)
                removed += 1

        self.assertEqual(removed, 2)
        self.assertEqual(len(paragraphs(body)), 1)

    def test_a_table_counts_as_remaining_block_content(self):
        body = parse(f"<w:body {NS}><w:p/><w:tbl/></w:body>")

        self.assertTrue(can_delete_paragraph(paragraphs(body)[0]))

    def test_tables_and_paragraphs_interleaved(self):
        body = parse(f"<w:body {NS}><w:tbl/><w:p/><w:tbl/><w:p/></w:body>")

        for para in paragraphs(body):
            self.assertTrue(can_delete_paragraph(para))

    def test_non_block_markers_are_not_block_content(self):
        # A bookmark is not something Word will accept a body consisting of.
        body = parse(
            f'<w:body {NS}><w:bookmarkStart w:id="1" w:name="a"/>'
            f'<w:p/><w:bookmarkEnd w:id="1"/><w:sectPr/></w:body>'
        )

        self.assertFalse(can_delete_paragraph(paragraphs(body)[0]))

    def test_markers_before_between_and_after_do_not_hide_a_sibling(self):
        body = parse(
            f'<w:body {NS}><w:bookmarkStart w:id="1" w:name="a"/><w:p/>'
            f'<w:bookmarkEnd w:id="1"/><w:p/><w:sectPr/></w:body>'
        )

        self.assertTrue(can_delete_paragraph(paragraphs(body)[0]))

    def test_a_cell_must_still_end_with_a_paragraph(self):
        cell = parse(f"<w:tc {NS}><w:p/><w:tbl/><w:p/></w:tc>")

        first, last = paragraphs(cell)

        self.assertTrue(can_delete_paragraph(first))
        self.assertFalse(can_delete_paragraph(last), "the cell would end in a table")

    def test_a_cell_whose_only_paragraph_is_last_keeps_it(self):
        cell = parse(f"<w:tc {NS}><w:tbl/><w:p/></w:tc>")

        self.assertFalse(can_delete_paragraph(paragraphs(cell)[0]))

    def test_a_sole_paragraph_in_a_cell_may_not_be_deleted(self):
        cell = parse(f"<w:tc {NS}><w:p/></w:tc>")

        self.assertFalse(can_delete_paragraph(paragraphs(cell)[0]))

    def test_a_cell_ending_in_a_paragraph_after_removal_is_fine(self):
        cell = parse(f"<w:tc {NS}><w:p/><w:p/></w:tc>")

        self.assertTrue(can_delete_paragraph(paragraphs(cell)[0]))

    def test_section_properties_at_the_end_of_a_body_are_not_block_content(self):
        # w:sectPr is a sibling of the blocks, not one of them: a body left
        # holding only a sectPr has no content.
        body = parse(f"<w:body {NS}><w:p/><w:sectPr/></w:body>")

        self.assertFalse(can_delete_paragraph(paragraphs(body)[0]))

    def test_a_nested_container_answers_for_its_own_parent(self):
        # The inner cell's paragraph is judged against the cell, not the body.
        body = parse(
            f"<w:body {NS}><w:p/><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body>"
        )
        inner = body.find(f".//{W}tc")

        self.assertFalse(can_delete_paragraph(paragraphs(inner)[0]))
        self.assertTrue(can_delete_paragraph(paragraphs(body)[0]))

    def test_a_structured_document_tag_is_a_block_container_too(self):
        # w:sdtContent is in BLOCK_CONTAINERS, so its sole paragraph is kept
        # like a body's.  (verify.py's lint deliberately skips it, because an
        # *inline* sdtContent legitimately holds runs — a different question.)
        content = parse(f"<w:sdtContent {NS}><w:p/></w:sdtContent>")

        self.assertFalse(can_delete_paragraph(paragraphs(content)[0]))

    def test_a_container_word_does_not_constrain_permits_deletion(self):
        row = parse(f"<w:tr {NS}><w:p/></w:tr>")

        self.assertTrue(can_delete_paragraph(paragraphs(row)[0]))

    def test_a_detached_paragraph_is_never_deletable(self):
        self.assertFalse(can_delete_paragraph(parse(f"<w:p {NS}/>")))


class EquivalenceTests(unittest.TestCase):
    """Every arrangement small enough to enumerate decides the same way.

    A refactor that claims to preserve behaviour should be checked against the
    behaviour, not against a handful of cases someone thought of.  This walks
    every arrangement of up to four children over each container kind and
    compares the scan with the list-building version it replaced.
    """

    @staticmethod
    def _by_materialising(para: etree._Element) -> bool:
        """The previous implementation, kept here as the oracle."""
        parent = para.getparent()
        if parent is None:
            return False
        if parent.tag not in BLOCK_CONTAINERS:
            return True
        remaining = [c for c in block_children(parent) if c is not para]
        if not remaining:
            return False
        if parent.tag == TC_TAG:
            if not any(c.tag == P_TAG for c in remaining):
                return False
            if remaining[-1].tag != P_TAG:
                return False
        return True

    def test_the_scan_agrees_with_the_list_it_replaced(self):
        kinds = ["<w:p/>", "<w:tbl/>",
                 '<w:bookmarkStart w:id="1" w:name="a"/>', "<w:sectPr/>"]
        containers = ["w:body", "w:tc", "w:hdr", "w:ftr", "w:footnote",
                      "w:endnote", "w:txbxContent", "w:sdtContent", "w:tr"]

        checked = 0
        for tag in containers:
            for count in range(1, 5):
                for combo in itertools.product(kinds, repeat=count):
                    container = parse(f"<{tag} {NS}>{''.join(combo)}</{tag}>")
                    for para in paragraphs(container):
                        checked += 1
                        self.assertEqual(
                            can_delete_paragraph(para),
                            self._by_materialising(para),
                            f"{tag} {combo}",
                        )

        # Named so a future change that silently narrows the sweep is visible.
        self.assertGreater(checked, 2000)


if __name__ == "__main__":
    unittest.main()
