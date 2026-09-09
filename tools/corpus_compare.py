"""The corpus harness — record what a build decides, then diff two recordings.

W02 narrows the shipped rules and W03 changes what verification accepts. The
question those packages have to answer is not "do the tests pass" but "what
did this change decide differently, on real documents, and was each difference
intended". That needs a record of the decisions a build makes, taken before
the change and again after.

A recording is one row per decision: which document, which part, where in it,
what was done, and the rule that did it. Diffing two recordings gives the
changed decisions and nothing else.

Sources are never modified and nothing is cleaned to disk — the recording is a
dry run.

    python -m tools.corpus_compare record BASELINE.json SPEC.docx [MORE ...]
    python -m tools.corpus_compare diff BASELINE.json CANDIDATE.json

**Privacy.** Rows carry a truncated preview of the text a decision was made
about, because a diff nobody can read is not reviewable. Specifications are
usually proprietary: keep recordings in a scratch directory, and do not commit
one without the maintainer's say-so. ``--no-text`` omits previews entirely and
leaves only counts and rules, which still diffs usefully.
"""

import json
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from apppaths import resolve_config_path            # noqa: E402
from detection import ContentType, DetectionEngine  # noqa: E402
from docx_xml import load_config                    # noqa: E402
from processor import DocxProcessor                 # noqa: E402
from tools.census_formatting import collect_paths   # noqa: E402

PREVIEW_WIDTH = 100


@dataclass(frozen=True)
class Decision:
    """One thing a build decided to do to one piece of content."""

    document: str
    #: ``removed``, ``redacted`` or ``preserved``.
    action: str
    category: str
    rule: str
    preview: str

    @property
    def identity(self) -> tuple[str, str, str]:
        """What makes two decisions the same decision across two recordings.

        Deliberately not the position: narrowing a rule changes how many
        paragraphs are removed, so every later position shifts and a
        position-keyed diff would report the whole document as changed.
        """
        return (self.document, self.preview, self.category)


def _preview(text: str, keep_text: bool) -> str:
    if not keep_text:
        return ""
    flat = " ".join(text.split())
    return flat if len(flat) <= PREVIEW_WIDTH else flat[:PREVIEW_WIDTH] + "..."


def record_one(path: Path, config: dict, keep_text: bool = True) -> list[Decision]:
    """Every decision a dry run makes about one document."""
    engine = DetectionEngine(config)
    result = DocxProcessor(engine, dry_run=True).process(
        path, path.parent / "unused.docx"
    )
    if result.errors:
        return [Decision(path.name, "error", "processing", "; ".join(result.errors), "")]

    decisions: list[Decision] = []
    for detection in result.detections:
        if detection.content_type == ContentType.PRESERVE:
            action = "preserved"
        elif detection.content_type == ContentType.INLINE_PLACEHOLDER:
            action = "redacted"
        elif engine.should_remove([detection]):
            action = "removed"
        else:
            continue  # detected but not acted on; not a decision
        decisions.append(Decision(
            document=path.name,
            action=action,
            category=detection.content_type.value,
            rule=detection.reason,
            preview=_preview(detection.text, keep_text),
        ))
    return decisions


def record(paths: list[Path], config: dict, keep_text: bool = True) -> list[Decision]:
    return [d for path in paths for d in record_one(path, config, keep_text)]


def save(decisions: list[Decision], destination: Path) -> None:
    destination.write_text(
        json.dumps([asdict(d) for d in decisions], indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def load(source: Path) -> list[Decision]:
    return [Decision(**row) for row in json.loads(source.read_text(encoding="utf-8"))]


@dataclass
class Difference:
    """One decision the candidate makes differently from the baseline."""

    kind: str            # "no longer acted on", "newly acted on", "action changed"
    baseline: Decision | None
    candidate: Decision | None

    def describe(self) -> str:
        subject = self.candidate or self.baseline
        assert subject is not None
        lines = [f"[{self.kind}] {subject.document}: {subject.preview or '(no text)'}"]
        if self.baseline is not None:
            lines.append(f"    was: {self.baseline.action} — {self.baseline.rule}")
        if self.candidate is not None:
            lines.append(f"    now: {self.candidate.action} — {self.candidate.rule}")
        return "\n".join(lines)


def diff(baseline: list[Decision], candidate: list[Decision]) -> list[Difference]:
    """The decisions that changed, and only those."""
    before = {d.identity: d for d in baseline}
    after = {d.identity: d for d in candidate}

    differences: list[Difference] = []
    for identity, old in before.items():
        new = after.get(identity)
        if new is None:
            differences.append(Difference("no longer acted on", old, None))
        elif new.action != old.action or new.rule != old.rule:
            differences.append(Difference("action changed", old, new))
    for identity, new in after.items():
        if identity not in before:
            differences.append(Difference("newly acted on", None, new))

    return differences


def format_diff(differences: list[Difference]) -> str:
    if not differences:
        return "No decisions changed."
    counts: dict[str, int] = {}
    for difference in differences:
        counts[difference.kind] = counts.get(difference.kind, 0) + 1
    header = ", ".join(f"{count} {kind}" for kind, count in sorted(counts.items()))
    body = "\n".join(difference.describe() for difference in differences)
    return (
        f"{len(differences)} changed decision(s): {header}\n\n{body}\n\n"
        "Each one needs a verdict: an intended conservative retention, a "
        "confirmed precision improvement, or unresolved."
    )


def main(argv: list[str]) -> int:
    keep_text = "--no-text" not in argv
    argv = [a for a in argv if a != "--no-text"]

    if len(argv) >= 3 and argv[0] == "record":
        destination, paths = Path(argv[1]), collect_paths(argv[2:])
        if not paths:
            print("No .docx files found.")
            return 1
        decisions = record(paths, load_config(resolve_config_path()), keep_text)
        save(decisions, destination)
        print(f"Recorded {len(decisions)} decision(s) from {len(paths)} document(s) "
              f"to {destination}")
        return 0

    if len(argv) == 3 and argv[0] == "diff":
        print(format_diff(diff(load(Path(argv[1])), load(Path(argv[2])))))
        return 0

    print(__doc__)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
