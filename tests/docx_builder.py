"""
Synthetic DOCX builder for the SpecCleanse test suite.

Builds minimal but structurally real .docx files with ``zipfile`` so the
processor, verifier, and structural lint can be exercised end to end without
shipping binary fixtures or adding a dependency.
"""

import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#: Namespace declarations every generated part carries.
#:
#: ``w14`` is declared even though nothing here emits a ``w14:`` element,
#: because ``mc:Ignorable`` names it.  Markup Compatibility (ECMA-376 Part 3)
#: requires every prefix listed there to be a declared namespace prefix, and
#: an undeclared one makes the part non-conformant — so a fixture written that
#: way is not a valid positive control for Word validation, whatever else it
#: proves.  It named ``w14`` without declaring it until §17.1 asked.
NS_DECL = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'mc:Ignorable="w14"'
)

XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

RELS = XML_HEAD + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    '</Relationships>'
)

#: Relationship types for the parts a fixture may carry, keyed by the file name
#: relative to ``word/``.  A relationship is written **only** when its target is
#: actually in the package: a relationship naming a part that is not there is a
#: dangling one, which is exactly what
#: ``test_revisions.test_package_bookkeeping_is_updated`` asserts the *cleaner*
#: must never leave behind.  The builder used to emit the comments relationship
#: unconditionally, so every fixture without comments was invalid in the way the
#: project already treats as breaking.
_RELATIONSHIP_TYPE = {
    "comments.xml": "comments",
    "footnotes.xml": "footnotes",
    "endnotes.xml": "endnotes",
    "styles.xml": "styles",
}

_REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def _document_rels(part_names) -> str:
    """Relationships from ``word/document.xml`` to the parts that exist."""
    relationships = []
    for index, part_name in enumerate(sorted(part_names), start=10):
        inside_word = part_name.removeprefix("word/")
        kind = _RELATIONSHIP_TYPE.get(inside_word)
        if kind is None:
            continue
        relationships.append(
            f'<Relationship Id="rId{index}" Type="{_REL_BASE}{kind}" '
            f'Target="{inside_word}"/>'
        )
    return XML_HEAD + (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(relationships) +
        '</Relationships>'
    )

_CONTENT_TYPE = {
    "document": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "styles": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    "header": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    "footer": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    "footnotes": "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    "endnotes": "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
    "comments": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
}


# ---------------------------------------------------------------------------
# XML fragment helpers
# ---------------------------------------------------------------------------

def _toggle(tag: str, value) -> str:
    """Render a toggle property: True -> bare tag, False -> w:val="0"."""
    if value is None:
        return ""
    return f"<w:{tag}/>" if value else f'<w:{tag} w:val="0"/>'


def run(
    text: str = "",
    italic=None,
    bold=None,
    color: str | None = None,
    vanish=None,
    rstyle: str | None = None,
    inner: str = "",
) -> str:
    """Build a ``w:r``.  ``inner`` is raw XML appended after the text."""
    props = "".join([
        _toggle("i", italic),
        _toggle("b", bold),
        _toggle("vanish", vanish),
        f'<w:color w:val="{color}"/>' if color else "",
        f'<w:rStyle w:val="{rstyle}"/>' if rstyle else "",
    ])
    rpr = f"<w:rPr>{props}</w:rPr>" if props else ""
    body = f"<w:t xml:space=\"preserve\">{escape(text)}</w:t>" if text else ""
    return f"<w:r>{rpr}{body}{inner}</w:r>"


def para(
    *children: str,
    style: str | None = None,
    sect: bool = False,
    ppr_extra: str = "",
) -> str:
    """Build a ``w:p`` from run/marker fragments."""
    props = "".join([
        f'<w:pStyle w:val="{style}"/>' if style else "",
        ppr_extra,
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>' if sect else "",
    ])
    ppr = f"<w:pPr>{props}</w:pPr>" if props else ""
    return f"<w:p>{ppr}{''.join(children)}</w:p>"


def text_para(text: str, **kwargs) -> str:
    """Shorthand for a paragraph holding a single plain run."""
    return para(run(text), **kwargs)


def table(*cells: str) -> str:
    """Build a one-row ``w:tbl``; each argument is the content of one cell."""
    tcs = "".join(
        f'<w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>{cell}</w:tc>'
        for cell in cells
    )
    return f"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid><w:tr>{tcs}</w:tr></w:tbl>"


DRAWING = (
    '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
    '<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>'
    '</wp:inline></w:drawing>'
)


def field_begin() -> str:
    return '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'


def field_end() -> str:
    return '<w:r><w:fldChar w:fldCharType="end"/></w:r>'


def text_box(*paragraphs: str) -> str:
    """Build a run holding a VML text box with nested paragraphs."""
    return (
        '<w:r><w:pict><v:shape><v:textbox><w:txbxContent>'
        + "".join(paragraphs)
        + '</w:txbxContent></v:textbox></v:shape></w:pict></w:r>'
    )


def bookmark_start(bid: str, name: str = "bm") -> str:
    return f'<w:bookmarkStart w:id="{bid}" w:name="{name}"/>'


def bookmark_end(bid: str) -> str:
    return f'<w:bookmarkEnd w:id="{bid}"/>'


def inserted(*children: str, author: str = "Editor") -> str:
    """Wrap runs in a tracked insertion."""
    return (
        f'<w:ins w:id="90" w:author="{author}" w:date="2026-01-01T00:00:00Z">'
        f'{"".join(children)}</w:ins>'
    )


def deleted(text: str, author: str = "Editor") -> str:
    """A tracked deletion — its text lives in w:delText, not w:t."""
    return (
        f'<w:del w:id="91" w:author="{author}" w:date="2026-01-01T00:00:00Z">'
        f'<w:r><w:delText xml:space="preserve">{escape(text)}</w:delText></w:r></w:del>'
    )


def comment_anchor(comment_id: str = "1") -> str:
    """Comment range markers plus the reference run, around nothing."""
    return (
        f'<w:commentRangeStart w:id="{comment_id}"/>'
        f'<w:commentRangeEnd w:id="{comment_id}"/>'
        f'<w:r><w:commentReference w:id="{comment_id}"/></w:r>'
    )


COMMENTS_RELS = XML_HEAD + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
    'Target="media/image1.png"/>'
    '</Relationships>'
)


def deleted_row(*cells: str) -> str:
    """A table row whose deletion is tracked in w:trPr — its text stays plain."""
    tcs = "".join(f"<w:tc><w:tcPr/>{cell}</w:tc>" for cell in cells)
    return (
        '<w:tr><w:trPr>'
        '<w:del w:id="95" w:author="Editor" w:date="2026-01-01T00:00:00Z"/>'
        f'</w:trPr>{tcs}</w:tr>'
    )


def row(*cells: str) -> str:
    """A plain table row."""
    tcs = "".join(f"<w:tc><w:tcPr/>{cell}</w:tc>" for cell in cells)
    return f"<w:tr>{tcs}</w:tr>"


def table_of(*rows: str) -> str:
    """A table built from pre-made rows."""
    return (
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>'
        + "".join(rows)
        + "</w:tbl>"
    )


def comments(*bodies: str) -> str:
    """Build ``word/comments.xml``."""
    entries = "".join(
        f'<w:comment w:id="{i + 1}" w:author="Editor" w:date="2026-01-01T00:00:00Z">'
        f'{body}</w:comment>'
        for i, body in enumerate(bodies)
    )
    return XML_HEAD + f"<w:comments {NS_DECL}>{entries}</w:comments>"


def hyperlink(*children: str, anchor: str = "target") -> str:
    return f'<w:hyperlink w:anchor="{anchor}">{"".join(children)}</w:hyperlink>'


# ---------------------------------------------------------------------------
# Part builders
# ---------------------------------------------------------------------------

def document(*body: str, final_sect: bool = True) -> str:
    """Wrap body XML fragments in a ``w:document``."""
    sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>' if final_sect else ""
    return (
        XML_HEAD
        + f"<w:document {NS_DECL}><w:body>{''.join(body)}{sect}</w:body></w:document>"
    )


def header(*body: str) -> str:
    return XML_HEAD + f"<w:hdr {NS_DECL}>{''.join(body)}</w:hdr>"


def footer(*body: str) -> str:
    return XML_HEAD + f"<w:ftr {NS_DECL}>{''.join(body)}</w:ftr>"


def footnotes(*bodies: str) -> str:
    """Build ``footnotes.xml``; each argument is one footnote's paragraphs."""
    separators = (
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0">'
        '<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    )
    notes = "".join(
        f'<w:footnote w:id="{i + 1}">{body}</w:footnote>'
        for i, body in enumerate(bodies)
    )
    return XML_HEAD + f"<w:footnotes {NS_DECL}>{separators}{notes}</w:footnotes>"


def styles(*style_defs: str) -> str:
    return XML_HEAD + f"<w:styles {NS_DECL}>{''.join(style_defs)}</w:styles>"


def style_def(
    style_id: str,
    name: str | None = None,
    based_on: str | None = None,
    hidden: bool | None = None,
    style_type: str = "paragraph",
) -> str:
    """Build a ``w:style``.  ``hidden`` is three-valued, like ``w:vanish``:
    True declares hidden, False declares ``w:val="0"``, None says nothing."""
    vanish = "" if hidden is None else _toggle("vanish", hidden)
    parts = [
        f'<w:name w:val="{name if name is not None else style_id}"/>',
        f'<w:basedOn w:val="{based_on}"/>' if based_on else "",
        f"<w:rPr>{vanish}</w:rPr>" if vanish else "",
    ]
    return (
        f'<w:style w:type="{style_type}" w:styleId="{style_id}">'
        f'{"".join(parts)}</w:style>'
    )


# ---------------------------------------------------------------------------
# Packaging
# ---------------------------------------------------------------------------

def build_docx(path: Path, document_xml: str, extra_parts: dict[str, str] | None = None) -> Path:
    """Write a .docx containing ``document_xml`` plus any ``extra_parts``.

    ``extra_parts`` maps a path inside the package (e.g. ``word/footer1.xml``)
    to its XML text.
    """
    parts = {"word/document.xml": document_xml}
    parts.update(extra_parts or {})

    overrides = []
    for part_name in parts:
        base = Path(part_name).stem.rstrip("0123456789")
        content_type = _CONTENT_TYPE.get(base)
        if content_type:
            overrides.append(f'<Override PartName="/{part_name}" ContentType="{content_type}"/>')

    content_types = XML_HEAD + (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides) +
        '</Types>'
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/_rels/document.xml.rels", _document_rels(parts))
        for part_name, xml in parts.items():
            zf.writestr(part_name, xml)

    return path


def read_part(path: Path, part_name: str = "word/document.xml") -> str:
    """Read one XML part back out of a .docx."""
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read(part_name).decode("utf-8")
