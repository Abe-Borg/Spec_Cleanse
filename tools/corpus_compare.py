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
one without the maintainer's say-so.

``--no-text`` omits previews and leaves a short content digest in their place.
The digest is what makes two recordings comparable at all — without it, every
decision of one category in a document would collapse to one entry and a
missing decision would diff as no change. It is a fingerprint, not encryption:
a digest of a sentence you already suspect can be confirmed by hashing it. What
it prevents is a recording being *readable*, which is what committing one would
otherwise leak.
"""

import json
import sys
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from apppaths import resolve_config_path            # noqa: E402
from docx_xml import load_config                    # noqa: E402
from tools.actions import content_digest, iter_actions  # noqa: E402
from tools.census_formatting import collect_paths   # noqa: E402

PREVIEW_WIDTH = 100


@dataclass(frozen=True)
class Decision:
    """One thing a build decided to do to one piece of content."""

    document: str
    part: str
    #: ``removed``, ``redacted``, ``run removed``, ``preserved`` or ``error``.
    action: str
    #: Every category behind the action, joined. A paragraph can have more than
    #: one, and recording them separately made one action look like several.
    categories: str
    rules: str
    #: Fingerprint of the content, always present. Identity rests on this, so a
    #: recording taken with --no-text still compares.
    digest: str
    preview: str

    @property
    def content(self) -> tuple[str, str]:
        """Which piece of content this decision is about.

        Deliberately not the position: narrowing a rule changes how many
        paragraphs are removed, so every later position shifts and a
        position-keyed diff would report the whole document as changed.
        """
        return (self.document, self.digest)

    @property
    def detail(self) -> tuple[str, str, str]:
        """What was decided about it."""
        return (self.action, self.categories, self.rules)


def _preview(text: str, keep_text: bool) -> str:
    if not keep_text:
        return ""
    flat = " ".join(text.split())
    return flat if len(flat) <= PREVIEW_WIDTH else flat[:PREVIEW_WIDTH] + "..."


def record_one(path: Path, config: dict, keep_text: bool = True) -> list[Decision]:
    """Every decision a dry run makes about one document.

    One row per *action*, not per detection. A paragraph can carry several
    detections and still be one action; recording them separately meant that
    narrowing one of two rules removed a baseline row and the diff reported
    "no longer acted on" for content the candidate still deletes.
    """
    try:
        actions = iter_actions(path, config)
    except Exception as exc:
        return [Decision(path.name, "", "error", "processing", str(exc),
                         content_digest(str(exc)), "")]

    return [
        Decision(
            document=action.document,
            part=action.part,
            action=action.action,
            categories="; ".join(action.categories),
            rules="; ".join(action.rules),
            digest=action.digest,
            preview=_preview(action.text, keep_text),
        )
        for action in actions
    ]


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
        shown = subject.preview or f"(digest {subject.digest})"
        lines = [f"[{self.kind}] {subject.document}: {shown}"]
        if self.baseline is not None:
            lines.append(f"    was: {self.baseline.action} — {self.baseline.rules}")
        if self.candidate is not None:
            lines.append(f"    now: {self.candidate.action} — {self.candidate.rules}")
        return "\n".join(lines)


def _index(decisions: list[Decision]):
    """Group decisions by the content they are about, keeping multiplicity.

    A specification repeats boilerplate, so the same content carries the same
    decision many times over. Indexing by identity alone collapsed those to one
    entry, and a recording that lost two of three identical removals diffed as
    no change at all.
    """
    counts: dict[tuple[str, str], Counter] = {}
    rows: dict[tuple[tuple[str, str], tuple[str, str, str]], Decision] = {}
    for decision in decisions:
        counts.setdefault(decision.content, Counter())[decision.detail] += 1
        rows.setdefault((decision.content, decision.detail), decision)
    return counts, rows


def diff(baseline: list[Decision], candidate: list[Decision]) -> list[Difference]:
    """The decisions that changed, and only those."""
    before, before_rows = _index(baseline)
    after, after_rows = _index(candidate)

    differences: list[Difference] = []
    for content in sorted(before.keys() | after.keys()):
        was = before.get(content, Counter())
        now = after.get(content, Counter())
        lost = sorted((was - now).elements())
        gained = sorted((now - was).elements())

        # The same content decided differently: pair them off, so a rule change
        # reads as one changed decision rather than a removal and an addition.
        for old_detail, new_detail in zip(lost, gained):
            differences.append(Difference(
                "action changed",
                before_rows[(content, old_detail)],
                after_rows[(content, new_detail)],
            ))
        for old_detail in lost[len(gained):]:
            differences.append(Difference(
                "no longer acted on", before_rows[(content, old_detail)], None
            ))
        for new_detail in gained[len(lost):]:
            differences.append(Difference(
                "newly acted on", None, after_rows[(content, new_detail)]
            ))

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
