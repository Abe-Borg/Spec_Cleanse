"""Census B — how much of a clean sits inside a referenced bookmark range.

The plan considered protecting the content of every bookmark a `REF` field
points at, and rescoped that to detect-and-report because the effect on
cleaning volume was never bounded. Word bookmarks headings it cross-references,
so "protect referenced ranges" could in principle suppress most of a clean —
or almost none of it. Nobody has measured which.

This measures it: of the paragraphs a clean would remove, how many lie inside
the range of a bookmark some supported internal reference actually names.
A high share means retention needs a much narrower design. A low share means
it is a cheap follow-up.

Supported consumers are simple-field and complex-field ``REF``/``PAGEREF``/
``NOTEREF`` and internal hyperlink anchors, read from field instructions only —
never from visible prose that happens to contain the word REF.

Nothing is written; the clean is a dry run.

    python -m tools.census_references SPEC.docx [MORE.docx ...]
    python -m tools.census_references /path/to/folder

The field-instruction reader here is deliberately simple, because a census can
tolerate a miss that the cleaner could not. W06 needs the same reader for real,
and should promote a stricter version of it into docx_xml.py rather than import
this one.
"""

import re
import shutil
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from tempfile import mkdtemp

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lxml import etree                              # noqa: E402

from apppaths import resolve_config_path            # noqa: E402
from detection import DetectionEngine               # noqa: E402
from docx_xml import (                              # noqa: E402
    P_TAG,
    W,
    collect_content_parts,
    load_config,
    load_styles,
    paragraph_text,
    parse_xml,
)
from processor import DocxProcessor                 # noqa: E402
from tools.census_formatting import collect_paths   # noqa: E402

#: A field instruction naming a bookmark.  The name is the first argument, and
#: may be quoted; switches such as ``\h`` follow it.
_FIELD_TARGET = re.compile(
    r'^\s*(?:REF|PAGEREF|NOTEREF)\s+(?:"([^"]*)"|([^\s\\]+))', re.IGNORECASE
)


def field_target(instruction: str) -> str | None:
    """The bookmark a field instruction names, or None if it names none."""
    match = _FIELD_TARGET.match(instruction or "")
    if match is None:
        return None
    return match.group(1) or match.group(2) or None


def _complex_field_instructions(elements: list[etree._Element]) -> list[str]:
    """Instruction text of every complex field, one string per field.

    A complex field is a ``w:fldChar`` begin, some ``w:instrText`` runs, a
    separate, its cached result, and an end. The instruction can be split
    across runs, so it is reassembled. Fields nest, hence the stack.
    """
    instructions: list[str] = []
    open_fields: list[list[str]] = []

    for element in elements:
        if element.tag == f"{W}fldChar":
            kind = element.get(f"{W}fldCharType")
            if kind == "begin":
                open_fields.append([])
            elif kind == "end" and open_fields:
                instructions.append("".join(open_fields.pop()))
        elif element.tag == f"{W}instrText" and open_fields:
            open_fields[-1].append(element.text or "")

    return instructions


def referenced_bookmarks(elements: list[etree._Element]) -> set[str]:
    """Names that some supported internal consumer in this part points at."""
    names: set[str] = set()

    for element in elements:
        if element.tag == f"{W}fldSimple":
            target = field_target(element.get(f"{W}instr") or "")
            if target:
                names.add(target)
        elif element.tag == f"{W}hyperlink":
            anchor = element.get(f"{W}anchor")
            if anchor:
                names.add(anchor)

    for instruction in _complex_field_instructions(elements):
        target = field_target(instruction)
        if target:
            names.add(target)

    return names


def _owning_paragraph(element: etree._Element) -> etree._Element | None:
    node = element.getparent()
    while node is not None:
        if node.tag == P_TAG:
            return node
        node = node.getparent()
    return None


def bookmark_paragraphs(elements: list[etree._Element]) -> dict[str, set[int]]:
    """Map each bookmark name to the positions in ``elements`` its range covers.

    ``elements`` is the part in document order, materialised once by the
    caller. That is not a convenience: lxml builds element proxies on demand
    and lets them go when nothing holds a reference, so ``id()`` is stable only
    while the list is alive. Two separate ``root.iter()`` passes can hand back
    different proxy objects for one node, and any position map built across
    them is wrong.

    A marker sitting inside a paragraph is credited to that paragraph, so a
    bookmark opened mid-sentence still covers the sentence it opened in. A
    bookmark with no matching end covers nothing: an unterminated range is not
    evidence about any particular paragraph.
    """
    order = {id(element): index for index, element in enumerate(elements)}
    paragraph_positions = [
        index for index, element in enumerate(elements) if element.tag == P_TAG
    ]

    def position(marker: etree._Element) -> int:
        para = _owning_paragraph(marker)
        if para is not None and id(para) in order:
            return order[id(para)]
        return order[id(marker)]

    ends = {
        element.get(f"{W}id"): element
        for element in elements
        if element.tag == f"{W}bookmarkEnd"
    }

    covered: dict[str, set[int]] = {}
    for start in elements:
        if start.tag != f"{W}bookmarkStart":
            continue
        name = start.get(f"{W}name")
        end = ends.get(start.get(f"{W}id"))
        if not name or end is None:
            continue
        first, last = position(start), position(end)
        covered.setdefault(name, set()).update(
            pos for pos in paragraph_positions if first <= pos <= last
        )

    return covered


@dataclass
class ReferenceCensus:
    """What one document's removals owe to referenced bookmark ranges."""

    path: Path
    bookmarks: int = 0
    referenced_bookmarks: int = 0
    removable_paragraphs: int = 0
    removable_inside_referenced_range: int = 0
    examples: list[str] = field(default_factory=list)
    error: str | None = None

    @property
    def share_inside(self) -> float | None:
        """The share of removals that retention would have to give up."""
        if not self.removable_paragraphs:
            return None
        return self.removable_inside_referenced_range / self.removable_paragraphs


def census_one(path: Path, config: dict, sample: int = 5) -> ReferenceCensus:
    """Measure one document."""
    report = ReferenceCensus(path=path)
    temp_dir = Path(mkdtemp(prefix="speccleanse_census_"))
    try:
        with zipfile.ZipFile(path, "r") as archive:
            archive.extractall(temp_dir)

        engine = DetectionEngine(config)
        engine.bind_styles(load_styles(temp_dir / "word"))
        processor = DocxProcessor(engine, dry_run=True)

        for xml_path in collect_content_parts(temp_dir / "word"):
            root = parse_xml(xml_path).getroot()
            # Materialised once and held for the whole part: see
            # bookmark_paragraphs on why id() needs a live reference.
            elements = list(root.iter())

            spans = bookmark_paragraphs(elements)
            referenced = referenced_bookmarks(elements)
            report.bookmarks += len(spans)
            report.referenced_bookmarks += len(referenced & spans.keys())

            protected: set[int] = set()
            for name in referenced & spans.keys():
                protected |= spans[name]

            for index, element in enumerate(elements):
                if element.tag != P_TAG:
                    continue
                detections = processor._process_paragraph(element)
                if not processor._should_remove_paragraph(element, detections):
                    continue
                report.removable_paragraphs += 1
                if index in protected:
                    report.removable_inside_referenced_range += 1
                    if len(report.examples) < sample:
                        report.examples.append(paragraph_text(element).strip())
    except Exception as exc:  # a census must not stop on one bad file
        report.error = str(exc)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    return report


def census(paths: list[Path], config: dict) -> list[ReferenceCensus]:
    return [census_one(path, config) for path in paths]


def format_report(reports: list[ReferenceCensus]) -> str:
    """A readable summary, and the totals the reference-scope decision needs."""
    lines = [
        f"{'document':<40} {'bkmks':>6} {'refd':>6} {'removable':>10} "
        f"{'inside':>7} {'share':>7}",
        "-" * 82,
    ]
    totals = dict(bookmarks=0, referenced=0, removable=0, inside=0)
    failed = []

    for report in reports:
        if report.error:
            failed.append(report)
            continue
        share = report.share_inside
        lines.append(
            f"{report.path.name[:40]:<40} {report.bookmarks:>6} "
            f"{report.referenced_bookmarks:>6} {report.removable_paragraphs:>10} "
            f"{report.removable_inside_referenced_range:>7} "
            f"{('-' if share is None else format(share, '.1%')):>7}"
        )
        totals["bookmarks"] += report.bookmarks
        totals["referenced"] += report.referenced_bookmarks
        totals["removable"] += report.removable_paragraphs
        totals["inside"] += report.removable_inside_referenced_range

    lines.append("-" * 82)
    overall = (totals["inside"] / totals["removable"]) if totals["removable"] else None
    lines.append(
        f"{'TOTAL':<40} {totals['bookmarks']:>6} {totals['referenced']:>6} "
        f"{totals['removable']:>10} {totals['inside']:>7} "
        f"{('-' if overall is None else format(overall, '.1%')):>7}"
    )

    examples = [text for report in reports for text in report.examples][:10]
    if examples:
        lines.append("")
        lines.append("Removable text inside a referenced range:")
        lines.extend(f"  {text[:100]}" for text in examples)

    if failed:
        lines.append("")
        lines.append("Could not be measured:")
        lines.extend(f"  {r.path.name}: {r.error}" for r in failed)

    if not totals["removable"]:
        lines.append("")
        lines.append(
            "No removals measured — the census says nothing about reference scope."
        )
    elif not totals["referenced"]:
        lines.append("")
        lines.append(
            "No referenced bookmarks found; retention would suppress nothing here."
        )

    return "\n".join(lines)


def main(argv: list[str]) -> int:
    if not argv:
        print(__doc__)
        return 2
    paths = collect_paths(argv)
    if not paths:
        print("No .docx files found.")
        return 1
    config = load_config(resolve_config_path())
    print(format_report(census(paths, config)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
