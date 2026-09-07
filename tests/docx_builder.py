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

NS_DECL = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
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

DOC_RELS = XML_HEAD + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId10" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
    'Target="comments.xml"/>'
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
    hidden: bool = False,
    style_type: str = "paragraph",
) -> str:
    parts = [
        f'<w:name w:val="{name if name is not None else style_id}"/>',
        f'<w:basedOn w:val="{based_on}"/>' if based_on else "",
        "<w:rPr><w:vanish/></w:rPr>" if hidden else "",
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
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
        for part_name, xml in parts.items():
            zf.writestr(part_name, xml)

    return path


def read_part(path: Path, part_name: str = "word/document.xml") -> str:
    """Read one XML part back out of a .docx."""
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read(part_name).decode("utf-8")
