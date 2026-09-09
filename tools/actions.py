"""What a build would actually do to a document, one row per action.

Both censuses and the corpus harness need the same thing: not what the
detectors *found*, but what the processor would *do*. Those are different, and
building on the first is what made the first cut of these tools wrong.

A paragraph can carry several detections and still be one action. A specifier
rule and a copyright rule matching the same note produce two detections and one
removed paragraph; counting detections inflated the removal total and therefore
the very percentage a default decision was to be taken on. A paragraph made of
nothing but a placeholder produces an inline detection and is *removed* whole,
not redacted.

So the decision order here mirrors ``DocxProcessor._process_xml_file`` exactly —
preserve, then whole-paragraph removal, then redaction (or removal, when
cutting the placeholders leaves nothing), then run removal — and it asks the
processor's own methods rather than reimplementing the rules. The tools cannot
drift from the cleaner without this file failing to import.
"""

import hashlib
import shutil
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from tempfile import mkdtemp

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from detection import ContentType, DetectionEngine   # noqa: E402
from docx_xml import (                               # noqa: E402
    P_TAG,
    collect_content_parts,
    load_styles,
    paragraph_text,
    parse_xml,
    run_text,
)
from processor import DocxProcessor                  # noqa: E402

#: A whole paragraph goes.
REMOVED = "removed"
#: Placeholders are cut out and the sentence around them survives.
REDACTED = "redacted"
#: One run inside a surviving paragraph goes.
RUN_REMOVED = "run removed"
#: A preserve rule protects it.
PRESERVED = "preserved"


def content_digest(text: str) -> str:
    """A short stable fingerprint of some text.

    Lets two recordings be compared without either carrying the text. It is a
    fingerprint, not encryption: a digest of a sentence you already suspect can
    be confirmed by hashing it. It stops a recording from being *readable*,
    which is what committing one would otherwise leak.
    """
    return hashlib.sha256(" ".join(text.split()).encode("utf-8")).hexdigest()[:16]


@dataclass(frozen=True)
class Action:
    """One thing a build would do to one piece of content."""

    document: str
    part: str
    action: str
    #: Every category that spoke to this action, sorted — a paragraph can have
    #: more than one, and which of them was decisive is not knowable.
    categories: tuple[str, ...]
    rules: tuple[str, ...]
    text: str

    @property
    def digest(self) -> str:
        return content_digest(self.text)


def _evidence(detections) -> tuple[tuple[str, ...], tuple[str, ...]]:
    """The categories and rules behind an action, deduplicated and ordered."""
    categories = sorted({d.content_type.value for d in detections})
    rules = sorted({d.reason for d in detections if d.reason})
    return tuple(categories), tuple(rules)


def iter_actions(path: Path, config: dict) -> list[Action]:
    """Every action a dry run would take on one document, in document order.

    Raises on an unreadable package; callers decide whether that stops a run.
    """
    engine = DetectionEngine(config)
    actions: list[Action] = []
    temp_dir = Path(mkdtemp(prefix="speccleanse_actions_"))
    try:
        with zipfile.ZipFile(path, "r") as archive:
            archive.extractall(temp_dir)

        engine.bind_styles(load_styles(temp_dir / "word"))
        processor = DocxProcessor(engine, dry_run=True)

        for xml_path in collect_content_parts(temp_dir / "word"):
            root = parse_xml(xml_path).getroot()
            for para in [e for e in root.iter() if e.tag == P_TAG]:
                actions.extend(
                    _paragraph_actions(processor, para, path.name, xml_path.name)
                )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    return actions


def _paragraph_actions(
    processor: DocxProcessor, para, document: str, part: str
) -> list[Action]:
    """Mirror the processor's decision order for one paragraph."""
    detections = processor._process_paragraph(para)
    if not detections:
        return []

    categories, rules = _evidence(detections)
    text = paragraph_text(para).strip()

    if any(d.content_type == ContentType.PRESERVE for d in detections):
        return [Action(document, part, PRESERVED, categories, rules, text)]

    if processor._should_remove_paragraph(para, detections):
        return [Action(document, part, REMOVED, categories, rules, text)]

    spans = processor._redaction_spans(para, detections)
    if spans is not None:
        # An empty span list is the processor's signal that cutting every
        # placeholder leaves nothing behind, so the paragraph goes whole.
        # Recording that as a redaction misstates a deletion.
        action = REDACTED if spans else REMOVED
        return [Action(document, part, action, categories, rules, text)]

    runs = []
    for run, run_detections in processor._group_run_detections(para, detections):
        if not processor.engine.should_remove(run_detections):
            continue
        run_categories, run_rules = _evidence(run_detections)
        runs.append(Action(
            document, part, RUN_REMOVED, run_categories, run_rules,
            run_text(run).strip(),
        ))
    return runs
