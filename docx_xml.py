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
from collections.abc import Iterator
from dataclasses import dataclass, field
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
    """One entry from ``word/styles.xml``."""
    style_id: str
    name: str = ""
    based_on: str | None = None
    hidden: bool = False
    style_type: str = ""


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

        styles[style_id] = StyleInfo(
            style_id=style_id,
            name=(name_elem.get(f"{W}val") if name_elem is not None else "") or "",
            based_on=(based_on_elem.get(f"{W}val") if based_on_elem is not None else None),
            hidden=toggle_on(rpr, f"{W}vanish"),
            style_type=style.get(f"{W}type") or "",
        )

    return styles


def fold_style_name(name: str) -> str:
    """Normalise a style name/ID for comparison ("Specifier Note" == "specifiernote")."""
    return re.sub(r"[\s_-]+", "", name).lower()


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

#: Config keys whose values are lists of regular expressions.
PATTERN_KEYS = ("text_patterns", "low_confidence_patterns", "inline_patterns")


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

    return config
