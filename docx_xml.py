"""
SpecCleanse Shared DOCX/XML Helpers

Namespace constants, WordprocessingML structure rules, and text extraction
helpers shared by ``detection.py``, ``processor.py``, ``verify.py``, and
``gui.py``.

Everything in this module is document-model plumbing: it knows how
WordprocessingML nests runs inside paragraphs, which containers Word refuses
to open when they are left empty, and how to turn an element subtree into the
text string the detectors match against.  It holds no detection policy.
"""

import re
from collections import Counter
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

import yaml
from lxml import etree

# ---------------------------------------------------------------------------
# Namespaces
# ---------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML_NS = "http://www.w3.org/XML/1998/namespace"

W = f"{{{W_NS}}}"
MC = f"{{{MC_NS}}}"
XML = f"{{{XML_NS}}}"

NAMESPACES = {
    "w": W_NS,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": MC_NS,
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}

# Frequently used tag names
P_TAG = f"{W}p"
R_TAG = f"{W}r"
T_TAG = f"{W}t"
TBL_TAG = f"{W}tbl"
TC_TAG = f"{W}tc"
PPR_TAG = f"{W}pPr"
RPR_TAG = f"{W}rPr"
SECTPR_TAG = f"{W}sectPr"

# ---------------------------------------------------------------------------
# Structure rules
# ---------------------------------------------------------------------------

#: Containers that WordprocessingML requires to hold at least one block-level
#: child (``w:p`` or ``w:tbl``).  Emptying any of these produces a file Word
#: reports as having "unreadable content".  ``CT_HdrFtr`` carries
#: ``minOccurs="1"`` in the schema; table cells additionally have to *end*
#: with a ``w:p``.
BLOCK_CONTAINERS = frozenset({
    TC_TAG,
    f"{W}hdr",
    f"{W}ftr",
    f"{W}footnote",
    f"{W}endnote",
    f"{W}txbxContent",
    f"{W}comment",
    f"{W}body",
    f"{W}sdtContent",
})

#: Block-level children that satisfy the "container is not empty" rule.
BLOCK_LEVEL_TAGS = frozenset({P_TAG, TBL_TAG})

#: Elements that carry content a text pattern can never see but whose loss
#: damages the document: pictures, embedded objects, field plumbing, and
#: footnote/endnote/comment references.
EMBEDDED_CONTENT_TAGS = frozenset({
    f"{W}drawing",
    f"{W}pict",
    f"{W}object",
    f"{W}fldChar",
    f"{W}fldSimple",
    f"{W}instrText",
    f"{W}footnoteReference",
    f"{W}endnoteReference",
    f"{W}commentReference",
    f"{W}footnoteRef",
    f"{W}endnoteRef",
    f"{W}separator",
    f"{W}continuationSeparator",
})

#: Range markers that are legal both inside a paragraph and beside it at
#: block level, so an orphaned half can be relocated instead of dropped.
RANGE_MARKER_TAGS = frozenset({
    f"{W}bookmarkStart",
    f"{W}bookmarkEnd",
    f"{W}commentRangeStart",
    f"{W}commentRangeEnd",
})

#: Structural children kept when a paragraph's content is stripped in place.
KEEP_ON_STRIP = frozenset({PPR_TAG, f"{W}proofErr", f"{W}permStart", f"{W}permEnd"}) | RANGE_MARKER_TAGS

#: Text-carrying leaves removed when a run's text is stripped.  ``w:instrText``
#: is deliberately absent: dropping it breaks the field it belongs to.
TEXT_LEAF_TAGS = frozenset({
    T_TAG,
    f"{W}delText",
    f"{W}tab",
    f"{W}br",
    f"{W}cr",
    f"{W}noBreakHyphen",
    f"{W}softHyphen",
    f"{W}sym",
    f"{W}ptab",
})

#: How each text-carrying leaf contributes to a paragraph's extracted text.
#: Tabs and breaks become real whitespace so that ``\s+`` patterns and the
#: ``^`` anchored preserve patterns behave the way they read.
_TEXT_SEPARATORS = {
    f"{W}tab": "\t",
    f"{W}br": "\n",
    f"{W}cr": "\n",
    f"{W}ptab": "\t",
    f"{W}noBreakHyphen": "-",
}


# ---------------------------------------------------------------------------
# Property helpers
# ---------------------------------------------------------------------------

def is_on(elem: etree._Element | None) -> bool:
    """Return True if a WordprocessingML toggle property is switched on.

    Toggle properties such as ``w:i``, ``w:b`` and ``w:vanish`` are on when
    present with no ``w:val``, and off when ``w:val`` is ``0``/``false``/``off``.
    Word writes ``<w:vanish w:val="0"/>`` to *un-hide* a run that inherits
    hidden from its style, so presence alone is not truth.
    """
    if elem is None:
        return False
    val = elem.get(f"{W}val")
    if val is None:
        return True
    return val.strip().lower() in ("1", "true", "on")


def toggle_on(props: etree._Element | None, tag: str) -> bool:
    """Return True if ``props`` (an ``rPr``/``pPr``) has toggle ``tag`` on."""
    if props is None:
        return False
    return is_on(props.find(tag))


# ---------------------------------------------------------------------------
# Iteration that respects paragraph nesting
# ---------------------------------------------------------------------------

def _iter_own_descendants(node: etree._Element, tag: str) -> Iterator[etree._Element]:
    """Yield descendants with ``tag`` without crossing into a nested ``w:p``.

    Text boxes (``w:txbxContent``) embed whole paragraphs inside a run.  Those
    inner paragraphs are visited in their own right by :func:`iter_paragraphs`,
    so descending into them here would process and count their content twice.
    """
    for child in node:
        if child.tag == P_TAG:
            continue
        if child.tag == tag:
            yield child
            continue
        yield from _iter_own_descendants(child, tag)


def iter_own_runs(para: etree._Element) -> Iterator[etree._Element]:
    """Yield the ``w:r`` elements belonging to ``para`` itself.

    Covers runs nested in ``w:hyperlink``, ``w:ins``, ``w:sdtContent``,
    ``w:fldSimple``, ``w:smartTag`` and friends, but not runs that belong to a
    paragraph inside a text box.
    """
    return _iter_own_descendants(para, R_TAG)


def iter_paragraphs(
    root: etree._Element, skip_alternate_fallback: bool = False
) -> Iterator[etree._Element]:
    """Yield every ``w:p`` in document order, nested paragraphs included.

    ``mc:AlternateContent`` stores the same shape twice — once under
    ``mc:Choice`` and once under ``mc:Fallback`` — so its paragraphs appear
    twice.  Processing wants both branches (each has to be cleaned); text
    extraction wants the content counted once, which is what
    ``skip_alternate_fallback`` is for.
    """
    for elem in root.iter():
        if elem.tag != P_TAG:
            continue
        if skip_alternate_fallback and _in_alternate_fallback(elem):
            continue
        yield elem


def _in_alternate_fallback(elem: etree._Element) -> bool:
    """True if ``elem`` lives under an ``mc:Fallback`` branch."""
    node = elem.getparent()
    while node is not None:
        if node.tag == f"{MC}Fallback":
            return True
        node = node.getparent()
    return False


# ---------------------------------------------------------------------------
# Text extraction
# ---------------------------------------------------------------------------

def iter_text_nodes(scope: etree._Element) -> Iterator[tuple[etree._Element, str]]:
    """Yield ``(element, text)`` pairs for ``scope`` in document order.

    ``w:t`` contributes its own text; tabs, breaks and non-breaking hyphens
    contribute the whitespace or character they render as.  Nested paragraphs
    are skipped, so the offsets of the concatenated string map back onto the
    yielded elements — which is what in-place redaction relies on.
    """
    for child in scope:
        if child.tag == P_TAG:
            continue
        if child.tag == T_TAG:
            yield child, child.text or ""
        elif child.tag in _TEXT_SEPARATORS:
            yield child, _TEXT_SEPARATORS[child.tag]
        else:
            yield from iter_text_nodes(child)


def element_text(scope: etree._Element) -> str:
    """Extract the visible text of a paragraph, run, or any other element."""
    return "".join(text for _, text in iter_text_nodes(scope))


def paragraph_text(para: etree._Element) -> str:
    """Extract a paragraph's own text (excluding nested text-box paragraphs)."""
    return element_text(para)


def run_text(run: etree._Element) -> str:
    """Extract a run's text."""
    return element_text(run)


def has_embedded_content(elem: etree._Element) -> bool:
    """True if ``elem`` contains a picture, object, field, or note reference."""
    for node in elem.iter():
        if node.tag in EMBEDDED_CONTENT_TAGS:
            return True
    return False


#: ``w:br/@w:type`` values that move content instead of just wrapping it.  A
#: break with no type, or ``textWrapping``, is a soft line break.
LAYOUT_BREAK_TYPES = frozenset({"page", "column"})


#: Field keywords whose first argument names a bookmark.  Deliberately short:
#: a general Word field evaluator is not wanted here, only enough grammar to
#: say which bookmark a reference consumes.
REFERENCE_KEYWORDS = frozenset({"REF", "PAGEREF", "NOTEREF"})

#: One field-instruction token: a quoted string, or a run of non-space.
_INSTRUCTION_TOKEN = re.compile(r'"([^"]*)"|(\S+)')


def reference_target(instruction: str) -> str | None:
    """The bookmark a field instruction refers to, or None if it refers to none.

    Handles the quoted form Word writes for a name containing spaces.  Anything
    that is not one of :data:`REFERENCE_KEYWORDS` followed by a name is None —
    including a switch where the name should be, which is a malformed reference
    rather than a reference to a bookmark called ``\\h``.

    This reads an instruction that a field carrier actually held.  Ordinary
    prose containing the word "REF" is never a field and never reaches here.
    """
    tokens = [
        quoted if quoted is not None else bare
        for quoted, bare in (
            (m.group(1), m.group(2)) for m in _INSTRUCTION_TOKEN.finditer(instruction)
        )
    ]
    if len(tokens) < 2 or tokens[0].upper() not in REFERENCE_KEYWORDS:
        return None
    target = tokens[1]
    return None if target.startswith("\\") else target


def bookmark_names(scope: etree._Element) -> set[str]:
    """Every bookmark name defined in ``scope``, case-folded.

    Word matches bookmark names case-insensitively, so the inventory does too.
    """
    return {
        name.casefold()
        for element in scope.iter(f"{W}bookmarkStart")
        if (name := element.get(f"{W}name"))
    }


@dataclass(frozen=True)
class ReferenceConsumer:
    """Something that points at a bookmark, and enough to find it again.

    The name alone is not enough to act on.  A target referenced from the body,
    a header and a footnote breaks in three places, and someone repairing it has
    to be told all three — collapsing them into one message by target name says
    what is wrong without saying where.
    """
    name: str
    kind: str
    detail: str

    @property
    def folded(self) -> str:
        return self.name.casefold()

    def __str__(self) -> str:
        return f"{self.kind} {self.detail}"


def reference_consumers(scope: etree._Element) -> list[ReferenceConsumer]:
    """Everything in ``scope`` that still points at a bookmark.

    Two kinds are supported: a reference field, simple or complex, and an
    internal hyperlink, which names its target in ``w:anchor`` rather than
    through a field at all.  Names are kept as written — matching folds case,
    as Word does, but a report has to name what someone will search for.
    """
    consumers = [
        ReferenceConsumer(target, "field", instruction.strip())
        for instruction in field_instructions(scope)
        if (target := reference_target(instruction))
    ]
    consumers.extend(
        ReferenceConsumer(anchor, "hyperlink to", anchor)
        for link in scope.iter(f"{W}hyperlink")
        if (anchor := link.get(f"{W}anchor"))
    )
    return consumers


def in_tracked_deletion(node: etree._Element) -> bool:
    """True if accepting revisions would take ``node`` with it.

    Covers every markup that carries content away: the inline wrappers in
    :data:`REVISION_DELETE_TAGS` — ``w:del`` and ``w:moveFrom``, the *source*
    half of a move — a deleted table row recording it in ``w:trPr``, and a
    deleted cell in ``w:tcPr``.

    It reads that tuple rather than naming ``w:del`` itself, because the
    question here is exactly "does :func:`accept_revisions` remove this?" and
    a second, shorter list of the answer drifted from the first: a field inside
    a tracked move was reported lost from a run that had correctly accepted the
    move.
    """
    while node is not None:
        if node.tag in REVISION_DELETE_TAGS:
            return True
        if node.tag == f"{W}tr" and node.find(f"{W}trPr/{W}del") is not None:
            return True
        if node.tag == TC_TAG and node.find(f"{W}tcPr/{W}cellDel") is not None:
            return True
        node = node.getparent()
    return False


def field_instructions(scope: etree._Element, skip_deleted: bool = False) -> Counter:
    """Every field instruction inside ``scope``, however the field is written.

    Word records the same field two ways.  A *simple* field is one
    ``w:fldSimple`` carrying its instruction in ``w:instr``; a *complex* one is
    a run sequence delimited by ``w:fldChar``, with the instruction spread over
    ``w:instrText`` nodes in between.  A count of one kind says nothing about
    the other, which is how a document could lose a simple field and still look
    intact.

    Instructions are whitespace-normalised, because Word splits them across
    ``w:instrText`` nodes at arbitrary points.  A Counter rather than a set: a
    document may legitimately hold the same field twice, and losing one of them
    is still a loss.

    ``skip_deleted`` leaves out fields inside a tracked deletion, which a run
    that accepts revisions removes legitimately — the source revision is the
    evidence that explains their absence from the output.
    """
    found: Counter = Counter()

    for field in scope.iter(f"{W}fldSimple"):
        if skip_deleted and in_tracked_deletion(field):
            continue
        found[" ".join((field.get(f"{W}instr") or "").split())] += 1

    # A stack, not a depth counter: a field nested in another field's result is
    # its own carrier.  Accumulating into one buffer merged the two
    # instructions and emitted a single entry, so stripping the inner field's
    # begin/end while leaving its instruction text behind produced an identical
    # inventory — the loss of a live nested field was invisible.
    open_fields: list[tuple[list[str], bool]] = []
    for node in scope.iter(f"{W}fldChar", f"{W}instrText"):
        if node.tag == f"{W}instrText":
            if open_fields:
                open_fields[-1][0].append(node.text or "")
            continue
        kind = node.get(f"{W}fldCharType")
        if kind == "begin":
            open_fields.append(([], skip_deleted and in_tracked_deletion(node)))
        elif kind == "end" and open_fields:
            parts, deleted = open_fields.pop()
            if not deleted:
                found[" ".join("".join(parts).split())] += 1
    return found


def is_layout_break(node: etree._Element) -> bool:
    """True if ``node`` is a ``w:br`` that starts a new page or column.

    Text extraction renders every break as ``\n``, which makes a page break
    look like ordinary whitespace to a pattern.  It is not: it is page setup,
    and deleting one reflows the document from that point on.
    """
    if node.tag != f"{W}br":
        return False
    return (node.get(f"{W}type") or "") in LAYOUT_BREAK_TYPES


#: Note containers whose ``w:id`` identifies one story inside a shared part.
NOTE_TAGS = (f"{W}footnote", f"{W}endnote")


def note_identity(para: etree._Element) -> str | None:
    """The ``w:id`` of the footnote or endnote holding ``para``, if any.

    ``footnotes.xml`` is one part holding many independent stories.  Without
    this, a paragraph in note 3 and an identical paragraph in note 7 are two
    interchangeable entries in a flat list, and losing one can be explained by
    the other.
    """
    node = para
    while node is not None:
        if node.tag in NOTE_TAGS:
            return f"{etree.QName(node).localname}:{node.get(f'{W}id')}"
        node = node.getparent()
    return None


def run_signature(run: etree._Element) -> tuple:
    """A run's identity: its text and the formatting properties a reader sees.

    Pure document fact — nothing here decides whether a run is editorial, and
    no style resolution is involved.  The question it answers is narrower and
    syntactic: are these two runs interchangeable?  Two runs with the same text
    and the same properties are, and it does not matter which of them survived.

    It exists because extracted text cannot always say *which* of two identical
    occurrences a clean removed.  A hidden note followed by an identical visible
    requirement extracts as the same characters twice, so an output holding one
    copy is textually consistent with either having gone — and those two
    outcomes are a correct clean and a lost requirement.
    """
    rpr = run.find(f"{W}rPr")
    color = rpr.find(f"{W}color") if rpr is not None else None
    rstyle = rpr.find(f"{W}rStyle") if rpr is not None else None
    return (
        run_text(run),
        toggle_on(rpr, f"{W}vanish"),
        toggle_on(rpr, f"{W}i"),
        toggle_on(rpr, f"{W}b"),
        color.get(f"{W}val") if color is not None else None,
        rstyle.get(f"{W}val") if rstyle is not None else None,
    )


def run_profile(para: etree._Element) -> tuple:
    """Signatures of the runs in ``para`` that carry text, in document order.

    Runs with no substantive text are left out: they carry structure rather
    than content, and whether an empty one survives says nothing about whether
    the document kept what it had to.
    """
    return tuple(
        signature
        for signature in (run_signature(run) for run in iter_own_runs(para))
        if signature[0].strip()
    )


def paragraph_style(para: etree._Element) -> str | None:
    """The ``w:pStyle`` applied to ``para``, if any."""
    ppr = para.find(f"{W}pPr")
    if ppr is None:
        return None
    pstyle = ppr.find(f"{W}pStyle")
    return pstyle.get(f"{W}val") if pstyle is not None else None


def paragraph_signature(para: etree._Element) -> tuple:
    """A paragraph's full identity: its style and its runs.

    The style has to be here and not only in the runs.  Two paragraphs can hold
    character-for-character identical runs and still differ in what policy
    permits, because an editorial *paragraph* style is authority the runs know
    nothing about — which made a correct clean of the styled copy look like the
    loss of the plain one.
    """
    return (paragraph_style(para), run_profile(para))


def layout_break_offsets(scope: etree._Element) -> list[tuple[int, int]]:
    """Character ranges of the page and column breaks inside ``scope``.

    Offsets are into the same text :func:`element_text` produces, so a span
    computed against a paragraph's text can be tested against them directly.
    Both the processor (which refuses to cut across one) and verification
    (which must not then claim the cut was authorized) read this, so the two
    cannot disagree about where a break sits.
    """
    offsets: list[tuple[int, int]] = []
    offset = 0
    for node, text in iter_text_nodes(scope):
        start = offset
        offset += len(text)
        if is_layout_break(node):
            offsets.append((start, offset))
    return offsets


def strip_text_leaves(elem: etree._Element) -> None:
    """Remove text-carrying leaves from ``elem``, keeping everything else.

    Used to blank a run that has to stay put — because it holds a picture or
    half of a field — without leaving its editorial text behind.
    """
    for node in list(_iter_text_leaves(elem)):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)


def _iter_text_leaves(scope: etree._Element) -> Iterator[etree._Element]:
    for child in scope:
        if child.tag == P_TAG:
            continue
        if child.tag in TEXT_LEAF_TAGS:
            yield child
        else:
            yield from _iter_text_leaves(child)


def set_text(node: etree._Element, text: str) -> None:
    """Set a ``w:t`` node's text, adding ``xml:space`` when Word needs it."""
    node.text = text
    if text != text.strip():
        node.set(f"{XML}space", "preserve")
    elif node.get(f"{XML}space") is not None:
        del node.attrib[f"{XML}space"]


# ---------------------------------------------------------------------------
# Character spans
# ---------------------------------------------------------------------------

def merge_spans(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Sort and coalesce overlapping/adjacent ``(start, end)`` ranges."""
    merged: list[tuple[int, int]] = []
    for start, end in sorted(spans):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def tidy_spans(text: str, spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Widen each span by one trailing space when it sits between two spaces.

    Cutting "[Verify quantity]" out of "two [Verify quantity] spare filters"
    would otherwise leave a double space where the placeholder used to be.
    """
    tidied: list[tuple[int, int]] = []
    for start, end in spans:
        if start > 0 and text[start - 1] == " " and end < len(text) and text[end] == " ":
            end += 1
        tidied.append((start, end))
    return merge_spans(tidied)


def spans_cover(spans: list[tuple[int, int]], start: int, end: int) -> bool:
    """True if some span in ``spans`` contains all of ``[start, end)``."""
    return any(
        span_start <= start and span_end >= end for span_start, span_end in spans
    )


def cut_spans(text: str, spans: list[tuple[int, int]]) -> str:
    """Return ``text`` with every span removed."""
    kept: list[str] = []
    cursor = 0
    for start, end in merge_spans(spans):
        kept.append(text[cursor:start])
        cursor = max(cursor, end)
    kept.append(text[cursor:])
    return "".join(kept)


# ---------------------------------------------------------------------------
# Structure inspection
# ---------------------------------------------------------------------------

def block_children(container: etree._Element) -> list[etree._Element]:
    """Return the block-level (``w:p``/``w:tbl``) children of a container."""
    return [child for child in container if child.tag in BLOCK_LEVEL_TAGS]


def has_section_properties(para: etree._Element) -> bool:
    """True if ``para`` carries a ``w:sectPr`` — i.e. it is a section break.

    Deleting such a paragraph merges its section into the next one and takes
    its headers, footers, and page setup with it.
    """
    ppr = para.find(PPR_TAG)
    return ppr is not None and ppr.find(SECTPR_TAG) is not None


def can_delete_paragraph(para: etree._Element) -> bool:
    """True if the ``w:p`` element itself may be removed from its parent.

    False when removal would empty a container Word requires to hold block
    content, or would leave a table cell whose last child is not a paragraph.
    Callers strip the paragraph's content in place instead.
    """
    parent = para.getparent()
    if parent is None:
        return False

    if parent.tag not in BLOCK_CONTAINERS:
        return True

    remaining = [child for child in block_children(parent) if child is not para]
    if not remaining:
        return False

    if parent.tag == TC_TAG:
        # A cell must contain at least one paragraph and must end with one.
        if not any(child.tag == P_TAG for child in remaining):
            return False
        if remaining[-1].tag != P_TAG:
            return False

    return True


def field_chars_balanced(elem: etree._Element) -> bool:
    """True if every field begins and ends inside ``elem``.

    A ``w:fldChar`` ``begin`` whose ``end`` lives in a later paragraph (a
    multi-paragraph TOC field, say) must not be carried off on its own —
    unbalanced field structure is the kind of damage Word refuses to open.
    """
    depth = 0
    for node in elem.iter(f"{W}fldChar"):
        kind = node.get(f"{W}fldCharType")
        if kind == "begin":
            depth += 1
        elif kind == "end":
            depth -= 1
            if depth < 0:
                return False
    return depth == 0


def orphaned_range_markers(elem: etree._Element) -> list[etree._Element]:
    """Return range markers inside ``elem`` whose partner is outside it.

    Bookmarks and comment ranges are matched on ``w:id``.  An orphan is
    relocated beside the element rather than deleted with it, which keeps the
    range spanning the same content it always did.
    """
    orphans: list[etree._Element] = []
    for start_tag, end_tag in (
        (f"{W}bookmarkStart", f"{W}bookmarkEnd"),
        (f"{W}commentRangeStart", f"{W}commentRangeEnd"),
    ):
        starts = {}
        ends = {}
        for node in elem.iter(start_tag):
            starts.setdefault(node.get(f"{W}id"), []).append(node)
        for node in elem.iter(end_tag):
            ends.setdefault(node.get(f"{W}id"), []).append(node)
        for marker_id, nodes in starts.items():
            if len(ends.get(marker_id, [])) != len(nodes):
                orphans.extend(nodes)
        for marker_id, nodes in ends.items():
            if len(starts.get(marker_id, [])) != len(nodes):
                orphans.extend(nodes)
    return orphans


# ---------------------------------------------------------------------------
# Revisions and comments
# ---------------------------------------------------------------------------

#: Tracked changes whose content goes when the change is accepted.
REVISION_DELETE_TAGS = (f"{W}del", f"{W}moveFrom")

#: Tracked changes whose content stays; only the revision wrapper goes.
REVISION_UNWRAP_TAGS = (f"{W}ins", f"{W}moveTo")

#: Bookkeeping left behind by revision tracking: move ranges and the records
#: of formatting changes.  None of it carries content.
REVISION_MARKER_TAGS = (
    f"{W}moveFromRangeStart",
    f"{W}moveFromRangeEnd",
    f"{W}moveToRangeStart",
    f"{W}moveToRangeEnd",
    f"{W}pPrChange",
    f"{W}rPrChange",
    f"{W}sectPrChange",
    f"{W}tblPrChange",
    f"{W}tblPrExChange",
    f"{W}tcPrChange",
    f"{W}trPrChange",
    f"{W}cellIns",
    f"{W}cellDel",
    f"{W}cellMerge",
)

#: Comment anchors inside the document body.
COMMENT_MARKER_TAGS = (
    f"{W}commentRangeStart",
    f"{W}commentRangeEnd",
    f"{W}commentReference",
)

#: Comment parts of the package, removed together with their anchors.
COMMENT_PARTS = (
    "comments.xml",
    "commentsExtended.xml",
    "commentsIds.xml",
    "commentsExtensible.xml",
)


def unwrap_element(elem: etree._Element) -> None:
    """Replace an element with its children, in place."""
    parent = elem.getparent()
    if parent is None:
        return
    index = parent.index(elem)
    for child in reversed(list(elem)):
        parent.insert(index, child)
    parent.remove(elem)


def _remove_all(root: etree._Element, tags) -> int:
    removed = 0
    for tag in tags:
        for elem in list(root.iter(tag)):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)
                removed += 1
    return removed


def _accept_structural_deletions(root: etree._Element) -> int:
    """Remove table rows and cells that a tracked change marks as deleted.

    A deleted row keeps its text in ordinary ``w:t`` and records the deletion
    as a marker in ``w:trPr``; dropping only the marker would accept the row
    back into the document with its content intact, which is the opposite of
    accepting the deletion.  Cells work the same way through ``w:cellDel``.
    """
    changed = 0
    #: Tables a row or cell was actually removed from.  A pre-existing rowless
    #: table is not this run's doing, and rewriting the document because the
    #: revision checkbox happened to be on -- in a document with no revisions at
    #: all -- would change it for a reason unrelated to what was asked.
    touched: list[etree._Element] = []

    for marker_tag, properties_tag, container_tag in (
        (f"{W}del", f"{W}trPr", f"{W}tr"),
        (f"{W}cellDel", f"{W}tcPr", TC_TAG),
    ):
        for marker in list(root.iter(marker_tag)):
            properties = marker.getparent()
            if properties is None or properties.tag != properties_tag:
                continue
            container = properties.getparent()
            if container is None or container.tag != container_tag:
                continue
            row = container.getparent() if container_tag == TC_TAG else None
            table = _enclosing_table(container)
            if container.getparent() is not None:
                container.getparent().remove(container)
                changed += 1
                if table is not None:
                    touched.append(table)
            # A row emptied of every cell is no longer a row.
            if row is not None and row.getparent() is not None:
                if not any(child.tag == TC_TAG for child in row):
                    row.getparent().remove(row)

    return changed + _remove_rowless_tables(touched)


def _enclosing_table(node: etree._Element) -> etree._Element | None:
    """The nearest ``w:tbl`` ancestor of ``node``, if any."""
    while node is not None:
        if node.tag == TBL_TAG:
            return node
        node = node.getparent()
    return None


def _remove_rowless_tables(candidates: list[etree._Element]) -> int:
    """Remove tables that accepting revisions left with no rows.

    Deleting a table's last row already worked; the ``w:tbl`` around it stayed,
    holding nothing.  Word does not accept a table with no rows, and neither
    lint nor verification could see one — so the package looked clean and the
    file did not open.

    Only tables this acceptance actually took a row or cell from are considered.
    A table that arrived already rowless is the document's own problem, and
    removing it would rewrite a file that had no revisions to accept — the same
    rule the structural comparison follows, that only what this run did is this
    run's doing.

    Nesting needs no special traversal here: a table with no rows has no cells,
    so it can hold no inner table.  It is always a leaf, and the rows above were
    collected into a list before any of them were removed.

    Where the table was the only block its parent had, or the last block of a
    cell, an empty paragraph takes its place — that is the minimum the container
    requires, not a repair of the table.
    """
    changed = 0
    seen: set[int] = set()
    for table in candidates:
        if id(table) in seen:
            continue
        seen.add(id(table))
        if any(child.tag == f"{W}tr" for child in table):
            continue
        parent = table.getparent()
        if parent is None:
            continue
        position = list(parent).index(table)
        parent.remove(table)
        changed += 1

        if parent.tag not in BLOCK_CONTAINERS:
            continue
        remaining = block_children(parent)
        if not remaining:
            parent.insert(position, etree.Element(P_TAG))
        elif parent.tag == TC_TAG and remaining[-1].tag != P_TAG:
            parent.append(etree.Element(P_TAG))
    return changed


def accept_revisions(root: etree._Element) -> int:
    """Accept every tracked change in one XML part.

    Insertions keep their content and lose the revision wrapper; deletions go
    with their ``w:delText``, which no text extractor can see and which
    therefore survives an ordinary clean along with its markup.  Deleted table
    rows and cells go whole, since their text is not marked up at all.  A
    deleted paragraph *mark* is the one thing not acted on — the paragraphs
    stay separate — because merging them would move content the user never
    asked to move.
    """
    changed = _accept_structural_deletions(root)
    changed += _remove_all(root, REVISION_DELETE_TAGS)

    for tag in REVISION_UNWRAP_TAGS:
        for elem in list(root.iter(tag)):
            if elem.getparent() is not None:
                unwrap_element(elem)
                changed += 1

    return changed + _remove_all(root, REVISION_MARKER_TAGS)


def strip_comment_markers(root: etree._Element) -> int:
    """Remove comment anchors from one XML part."""
    return _remove_all(root, COMMENT_MARKER_TAGS)


def remove_comment_parts(unpacked_dir: Path) -> list[str]:
    """Delete the comment parts of a package, with their bookkeeping.

    A part left listed in the relationships or the content types after its
    file is gone is a package Word will not open, so both are updated, along
    with each part's own sidecar ``.rels``.
    """
    word_dir = unpacked_dir / "word"
    removed = [name for name in COMMENT_PARTS if (word_dir / name).exists()]
    if not removed:
        return []

    for name in removed:
        (word_dir / name).unlink()
        # A part's own relationships live in a sidecar named after it — a
        # comment holding an image or a hyperlink has one.  Left behind, it
        # relates to a part that no longer exists, which is an invalid package.
        sidecar = word_dir / "_rels" / f"{name}.rels"
        if sidecar.exists():
            sidecar.unlink()

    rels_path = word_dir / "_rels" / "document.xml.rels"
    if rels_path.exists():
        tree = parse_xml(rels_path)
        rels_root = tree.getroot()
        for rel in list(rels_root):
            target = (rel.get("Target") or "").lstrip("/").rsplit("/", 1)[-1]
            if target in removed:
                rels_root.remove(rel)
        write_xml(tree, rels_path)

    types_path = unpacked_dir / "[Content_Types].xml"
    if types_path.exists():
        tree = parse_xml(types_path)
        types_root = tree.getroot()
        for override in list(types_root):
            part_name = (override.get("PartName") or "").rsplit("/", 1)[-1]
            if part_name in removed:
                types_root.remove(override)
        write_xml(tree, types_path)

    return removed


# ---------------------------------------------------------------------------
# XML parts
# ---------------------------------------------------------------------------

#: XML parts inside ``word/`` that carry document content.  ``processor.py``
#: and ``verify.py`` both walk this list — keep them walking the same one.
def collect_content_parts(word_dir: Path, include_glossary: bool = True) -> list[Path]:
    """Return the content-bearing XML parts of an unpacked DOCX, in order."""
    parts: list[Path] = []

    doc_xml = word_dir / "document.xml"
    if doc_xml.exists():
        parts.append(doc_xml)

    parts.extend(sorted(word_dir.glob("header*.xml")))
    parts.extend(sorted(word_dir.glob("footer*.xml")))

    for extra in ("footnotes.xml", "endnotes.xml"):
        path = word_dir / extra
        if path.exists():
            parts.append(path)

    if include_glossary:
        glossary = word_dir / "glossary" / "document.xml"
        if glossary.exists():
            parts.append(glossary)

    return parts


def parse_xml(path: Path) -> etree._ElementTree:
    """Parse a DOCX XML part, preserving whitespace exactly."""
    parser = etree.XMLParser(remove_blank_text=False)
    return etree.parse(str(path), parser)


def write_xml(tree: etree._ElementTree, path: Path) -> None:
    """Write a DOCX XML part back with the declaration Word expects."""
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

@dataclass
class StyleInfo:
    """One entry from ``word/styles.xml``.

    ``vanish`` is three-valued: True when the style declares hidden, False when
    it declares ``w:val="0"``, and None when it says nothing at all.  The three
    are not interchangeable — ``w:vanish`` is a *toggle* property, so a
    declaration means something different from silence.
    """
    style_id: str
    name: str = ""
    based_on: str | None = None
    vanish: bool | None = None
    style_type: str = ""
    #: ``w:default="1"`` — the style Word applies where none is named.
    default: bool = False
    #: ``w:pPr/w:numPr/w:numId/@w:val`` if the style declares one.  Three-valued
    #: like ``vanish``: a value, the string ``"0"`` meaning *no* numbering, or
    #: None for silence.  ``"0"`` is an override, not a list called zero.
    num_id: str | None = None


def load_styles(word_dir: Path) -> dict[str, StyleInfo]:
    """Read ``word/styles.xml`` into a ``{styleId: StyleInfo}`` map.

    Returns an empty map when the part is missing or unparseable; style-aware
    detection then simply falls back to literal style-ID matching.
    """
    styles_path = word_dir / "styles.xml"
    if not styles_path.exists():
        return {}

    try:
        root = parse_xml(styles_path).getroot()
    except etree.XMLSyntaxError:
        return {}

    styles: dict[str, StyleInfo] = {}
    for style in root.iter(f"{W}style"):
        style_id = style.get(f"{W}styleId")
        if not style_id:
            continue

        name_elem = style.find(f"{W}name")
        based_on_elem = style.find(f"{W}basedOn")
        rpr = style.find(RPR_TAG)
        vanish_elem = rpr.find(f"{W}vanish") if rpr is not None else None
        num_id_elem = style.find(f"{W}pPr/{W}numPr/{W}numId")

        styles[style_id] = StyleInfo(
            style_id=style_id,
            name=(name_elem.get(f"{W}val") if name_elem is not None else "") or "",
            based_on=(based_on_elem.get(f"{W}val") if based_on_elem is not None else None),
            vanish=is_on(vanish_elem) if vanish_elem is not None else None,
            style_type=style.get(f"{W}type") or "",
            default=(style.get(f"{W}default") or "").strip().lower()
            in ("1", "true", "on"),
            num_id=(num_id_elem.get(f"{W}val") if num_id_elem is not None else None),
        )

    return styles


def numbering_id(para: etree._Element, styles: "StyleIndex") -> str | None:
    """The automatic-numbering list this paragraph belongs to, or None.

    Direct ``w:numPr`` first, then the paragraph style chain.  ``w:numId``
    ``"0"`` is Word's way of saying *no* numbering — an explicit override of an
    inherited value, not a list called zero — so it answers None rather than
    counting as participation.

    A ``w:numPr`` carrying only ``w:ilvl`` sets the level and leaves the list to
    the style, so it falls through rather than answering.
    """
    ppr = para.find(f"{W}pPr")
    if ppr is not None:
        direct = ppr.find(f"{W}numPr/{W}numId")
        if direct is not None:
            value = direct.get(f"{W}val")
            return None if value in (None, "0") else value

    # A paragraph naming no style still has one: Word applies the default
    # paragraph style, so a chain that started at None examined nothing and a
    # default style carrying w:numPr was invisible.
    style = paragraph_style(para) or styles.default_paragraph_style()
    for info in styles.chain(style):
        if info.num_id is not None:
            return None if info.num_id == "0" else info.num_id
    return None


def fold_style_name(name: str) -> str:
    """Normalise a style name/ID for comparison ("Specifier Note" == "specifiernote")."""
    return re.sub(r"[\s_-]+", "", name).lower()


def fold_style_names(names) -> frozenset[str]:
    """Fold a collection of configured style names for repeated comparison."""
    return frozenset(fold_style_name(name) for name in names if name)


class StyleIndex:
    """Resolves a paragraph or run style ID against ``word/styles.xml``.

    Comparing configured names to literal style IDs misses most of what firms
    actually ship: a template derives its note style from ``CMT`` under some
    other ID, or names it "Specifier Note" with a space, which no style ID can
    ever equal.  This walks the ``w:basedOn`` chain and matches display names
    as well as IDs.  With no styles part it degrades to folded ID matching,
    which is still an improvement on exact equality.
    """

    def __init__(self, styles: dict[str, StyleInfo] | None = None):
        self.styles = styles or {}

    def default_paragraph_style(self) -> str | None:
        """The style Word applies to a paragraph that names none."""
        for info in self.styles.values():
            if info.default and info.style_type == "paragraph":
                return info.style_id
        return None

    def chain(self, style_id: str | None) -> list[StyleInfo]:
        """The style and everything it is based on, nearest first."""
        resolved: list[StyleInfo] = []
        seen: set[str] = set()
        current = style_id
        while current and current not in seen:
            seen.add(current)
            info = self.styles.get(current)
            if info is None:
                resolved.append(StyleInfo(style_id=current))
                break
            resolved.append(info)
            current = info.based_on
        return resolved

    def matches(self, style_id: str | None, folded_names: frozenset[str]) -> bool:
        """True if the style, or any style it inherits from, is one of ``folded_names``."""
        if not style_id or not folded_names:
            return False
        for info in self.chain(style_id):
            if fold_style_name(info.style_id) in folded_names:
                return True
            if info.name and fold_style_name(info.name) in folded_names:
                return True
        return False

    def is_hidden(self, style_id: str | None) -> bool:
        """Effective hidden state of a style chain, under toggle semantics.

        MasterSpec hides its notes through the style rather than on each run.
        ``w:vanish`` is a toggle property, so declarations along the
        ``w:basedOn`` chain do not simply accumulate: a derived style that
        repeats its base style's ``<w:vanish/>`` switches hidden back *off*,
        and Word renders that text normally.  Treating every declaration as
        "on" would delete it.

        An explicit ``w:val="0"`` anywhere in the chain is taken as off
        outright.  Where the spec leaves room, the reading that keeps text is
        the one to take.
        """
        if not style_id:
            return False

        hidden = False
        for info in self.chain(style_id):
            if info.vanish is None:
                continue
            if info.vanish is False:
                return False
            hidden = not hidden
        return hidden


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

#: Config keys whose values are lists of regular expressions.
PATTERN_KEYS = ("text_patterns", "low_confidence_patterns", "inline_patterns")

#: Config keys that switch behaviour on or off, wherever they appear in a
#: section.  Checked because YAML quietly makes ``'false'`` a *string*, and
#: every reader here asks a plain truthiness question — so a quoted "off"
#: turns the option on.  For ``formatting_only_removal`` that is the one path
#: that removes text on no content evidence at all, switched on by someone
#: writing that it should be off.
BOOLEAN_KEYS = ("enabled", "formatting_only_removal")

#: What Word writes in ``w:color w:val``: six hexadecimal digits, no "#".
_HEX_COLOR = re.compile(r"\A[0-9A-Fa-f]{6}\Z")


def normalise_color(value, where: str) -> str:
    """Validate one editorial colour and return it as Word writes it.

    A leading ``#`` is normalised away rather than rejected.  It has exactly
    one possible meaning, it is how every other tool writes a hex colour, and
    the alternative is a rule that silently matches nothing — ``'#FF0000'``
    never equals the ``FF0000`` Word puts in ``w:val``, so the user's red text
    is simply never recognised and nothing says why.

    Anything else is refused, because there is no defensible guess.  Two
    measured failures this replaces: ``255`` raised ``AttributeError: 'int'
    object has no attribute 'upper'`` per file, naming no section, key or
    line; ``'bright red'`` was accepted and matched nothing, forever.
    """
    if not isinstance(value, str):
        raise ValueError(
            f"{where}: expected a colour like 'FF0000', got "
            f"{type(value).__name__} {value!r}"
        )
    candidate = value[1:] if value.startswith("#") else value
    if not _HEX_COLOR.match(candidate):
        raise ValueError(
            f"{where}: {value!r} is not a colour — expected six hexadecimal "
            "digits as Word writes them, for example 'FF0000' (a leading '#' "
            "is accepted and removed)"
        )
    return candidate.upper()


def compile_patterns(patterns: list[str], where: str) -> list[re.Pattern]:
    """Compile a list of pattern strings, naming the offender on failure."""
    compiled: list[re.Pattern] = []
    for index, pattern in enumerate(patterns):
        if not isinstance(pattern, str):
            raise ValueError(f"{where}[{index}]: expected a string, got {type(pattern).__name__}")
        try:
            compiled.append(re.compile(pattern, re.IGNORECASE | re.DOTALL))
        except re.error as exc:
            raise ValueError(f"{where}[{index}]: invalid regex {pattern!r} — {exc}") from exc
    return compiled


def load_config(config_path: Path) -> dict:
    """Load ``patterns.yaml`` as UTF-8 and validate every regex in it.

    The file is UTF-8 and contains ``©``, ``–`` and ``—``.  Reading it with the
    platform default encoding (cp1252 on Windows) silently mangles those into
    mojibake that matches nothing, so the encoding is pinned here rather than
    left to ``locale.getpreferredencoding()``.

    Raises:
        FileNotFoundError: the file does not exist.
        ValueError: the file is empty, is not a mapping, is not valid YAML, or
            contains a pattern that does not compile.
    """
    config_path = Path(config_path)
    if not config_path.exists():
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    with open(config_path, "r", encoding="utf-8") as handle:
        try:
            config = yaml.safe_load(handle)
        except yaml.YAMLError as exc:
            raise ValueError(f"{config_path.name} is not valid YAML — {exc}") from exc

    if config is None:
        raise ValueError(f"{config_path.name} is empty")
    if not isinstance(config, dict):
        raise ValueError(
            f"{config_path.name} must contain a mapping of sections, "
            f"got {type(config).__name__}"
        )

    for section_name, section in config.items():
        if not isinstance(section, dict):
            continue
        for key in PATTERN_KEYS:
            values = section.get(key)
            if values is None:
                continue
            if not isinstance(values, list):
                raise ValueError(f"{section_name}.{key} must be a list of patterns")
            compile_patterns(values, f"{section_name}.{key}")

        for key in BOOLEAN_KEYS:
            if key in section and not isinstance(section[key], bool):
                raise ValueError(
                    f"{section_name}.{key} must be true or false, got "
                    f"{type(section[key]).__name__} {section[key]!r}. Quoting it "
                    "makes it a string, and a non-empty string counts as true."
                )

        signals = section.get("formatting_signals")
        if isinstance(signals, dict) and signals.get("colors") is not None:
            colors = signals["colors"]
            where = f"{section_name}.formatting_signals.colors"
            if not isinstance(colors, list):
                raise ValueError(f"{where} must be a list of colours")
            # Normalised in place, so every reader downstream compares against
            # one shape and no caller has to remember to do this itself.
            signals["colors"] = [
                normalise_color(color, f"{where}[{index}]")
                for index, color in enumerate(colors)
            ]

    return config
