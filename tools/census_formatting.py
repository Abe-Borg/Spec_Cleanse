"""Census A — how much of a clean depends on formatting evidence alone.

``specifier_notes.formatting_only_removal`` lets italic-in-an-editorial-colour
remove text with no pattern and no style behind it. Some firms mark their notes
only that way; for others, real specification content is red and italic. The
plan proposes turning it off by default, and that proposal has so far rested on
an asymmetric-cost argument rather than on any measurement of what the change
would cost a real workflow.

This measures it directly, by cleaning each document twice — once with the
switch on, once off — and reporting the paragraphs that would newly survive.

It counts *paragraphs the processor would remove*, not detections. Those are
not the same number: a note matching both a specifier rule and a copyright rule
produces two detections and one removed paragraph, and counting detections
inflated the removal total and so deflated the very percentage this exists to
report. Counting detections that merely carry the formatting-only flag would be
wrong for a second reason: a paragraph whose formatting-only detection sits
alongside a pattern match is removed either way, and flipping the switch
changes nothing for it.

Nothing is written. Both runs are dry runs.

    python -m tools.census_formatting SPEC.docx [MORE.docx ...]
    python -m tools.census_formatting /path/to/folder
"""

import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from apppaths import resolve_config_path            # noqa: E402
from docx_xml import load_config                    # noqa: E402
from tools.actions import (                         # noqa: E402
    PRESERVED,
    REDACTED,
    REMOVED,
    iter_actions,
)


@dataclass
class FormattingCensus:
    """What one document's clean owes to formatting-only evidence."""

    path: Path
    #: Paragraphs removed with the switch on, as it ships today.
    removals_with_switch_on: int = 0
    #: Paragraphs removed with the switch off.
    removals_with_switch_off: int = 0
    #: Inline placeholder redactions; unaffected by the switch, shown for scale.
    inline_redactions: int = 0
    #: Paragraphs protected by a preserve rule; also shown for scale.
    preserved: int = 0
    #: A sample of the text that would newly survive, for eyeballing.
    examples: list[str] = field(default_factory=list)
    error: str | None = None

    @property
    def would_newly_survive(self) -> int:
        """Paragraphs the flip would stop removing — the cost of turning it off."""
        return self.removals_with_switch_on - self.removals_with_switch_off

    @property
    def share_of_removals(self) -> float | None:
        """That cost as a share of everything the clean removes today."""
        if not self.removals_with_switch_on:
            return None
        return self.would_newly_survive / self.removals_with_switch_on


def _with_switch(config: dict, formatting_only: bool) -> dict:
    """A copy of ``config`` with the switch forced one way."""
    copied = {
        key: (dict(value) if isinstance(value, dict) else value)
        for key, value in config.items()
    }
    copied.setdefault("specifier_notes", {})["formatting_only_removal"] = formatting_only
    return copied


def _removed_texts(actions) -> Counter:
    """Text of every paragraph the build would remove, as a multiset.

    A multiset rather than a set because a document repeats paragraphs —
    identical headings, identical boilerplate — and losing that multiplicity
    would understate the count.
    """
    return Counter(action.text for action in actions if action.action == REMOVED)


def census_one(path: Path, config: dict, sample: int = 5) -> FormattingCensus:
    """Measure one document."""
    report = FormattingCensus(path=path)
    try:
        with_switch = iter_actions(path, _with_switch(config, True))
        without_switch = iter_actions(path, _with_switch(config, False))
    except Exception as exc:  # a census must not stop on one bad file
        report.error = str(exc)
        return report

    on = _removed_texts(with_switch)
    off = _removed_texts(without_switch)
    report.removals_with_switch_on = sum(on.values())
    report.removals_with_switch_off = sum(off.values())

    lost = on - off
    report.examples = [text for text, _ in lost.most_common(sample) if text]

    for action in with_switch:
        if action.action == REDACTED:
            report.inline_redactions += 1
        elif action.action == PRESERVED:
            report.preserved += 1

    return report


def census(paths: list[Path], config: dict) -> list[FormattingCensus]:
    return [census_one(path, config) for path in paths]


def format_report(reports: list[FormattingCensus]) -> str:
    """A readable summary, and the totals the W02 decision needs."""
    lines = [
        f"{'document':<44} {'removed':>8} {'newly':>8} {'share':>7}",
        f"{'':<44} {'today':>8} {'survive':>8} {'':>7}",
        "-" * 70,
    ]
    total_on = total_lost = total_inline = total_preserved = 0
    failed = []

    for report in reports:
        if report.error:
            failed.append(report)
            continue
        share = report.share_of_removals
        lines.append(
            f"{report.path.name[:44]:<44} {report.removals_with_switch_on:>8} "
            f"{report.would_newly_survive:>8} "
            f"{('-' if share is None else format(share, '.1%')):>7}"
        )
        total_on += report.removals_with_switch_on
        total_lost += report.would_newly_survive
        total_inline += report.inline_redactions
        total_preserved += report.preserved

    lines.append("-" * 70)
    overall = (total_lost / total_on) if total_on else None
    lines.append(
        f"{'TOTAL':<44} {total_on:>8} {total_lost:>8} "
        f"{('-' if overall is None else format(overall, '.1%')):>7}"
    )
    lines.append("")
    lines.append(f"Inline redactions (unaffected by the switch): {total_inline}")
    lines.append(f"Paragraphs protected by a preserve rule:      {total_preserved}")

    examples = [text for report in reports for text in report.examples][:10]
    if examples:
        lines.append("")
        lines.append("Text that would newly survive:")
        lines.extend(f"  {text[:100]}" for text in examples)

    if failed:
        lines.append("")
        lines.append("Could not be measured:")
        lines.extend(f"  {r.path.name}: {r.error}" for r in failed)

    if not total_on:
        lines.append("")
        lines.append("No removals measured — the census says nothing about the switch.")

    return "\n".join(lines)


def collect_paths(arguments: list[str]) -> list[Path]:
    """Expand folder arguments into the .docx files inside them."""
    paths: list[Path] = []
    for argument in arguments:
        path = Path(argument)
        if path.is_dir():
            paths.extend(sorted(p for p in path.rglob("*.docx")
                                if not p.name.startswith("~$")))
        else:
            paths.append(path)
    return paths


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
