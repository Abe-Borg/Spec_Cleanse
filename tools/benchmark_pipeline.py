"""Timing and work counts for the clean-and-verify pipeline, on synthetic fixtures.

Built to answer §16.1's question — where does the time actually go, *now* —
rather than to reuse a number measured before W03–W07 rewrote the comparison.
A performance decision taken on a stale profile optimises whatever used to be
slow.

This is **not** the §8.4 corpus harness.  That one records decisions about real
documents; this one generates its own fixtures, holds the configuration fixed,
and reports timings.  Conflating them would mean comparing two builds on
documents that also differ.

Every fixture family §16.1 names is here, and family 1 is labelled what it is:
the *adversarial* case, not the representative one.  Family 2 is the closest
thing to a realistic distribution, and §16.4 forbids gating on family 1 alone.

Behaviour is compared through extracted semantic content — a digest of the
output's paragraph text, the field and reference inventories, and the
verification verdict — never through ZIP bytes, which mostly reflect how the
archive recompressed.

Usage::

    python -m tools.benchmark_pipeline                 # default sizes
    python -m tools.benchmark_pipeline --sizes 500,2000
    python -m tools.benchmark_pipeline --repeat 5 --out bench/
    python -m tools.benchmark_pipeline --family adversarial_headings
"""

import argparse
import hashlib
import json
import platform
import shutil
import statistics
import sys
import time
from dataclasses import dataclass, field, asdict
from pathlib import Path
from tempfile import mkdtemp

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from apppaths import resolve_config_path              # noqa: E402
from detection import DetectionEngine                 # noqa: E402
from docx_xml import (                                # noqa: E402
    collect_content_parts,
    field_instructions,
    load_config,
    parse_xml,
    reference_consumers,
)
from processor import DocxProcessor                   # noqa: E402
from verify import extract_paragraphs, verify_clean   # noqa: E402

from tests import docx_builder as db                  # noqa: E402

#: Sizes §16.1 names.  The largest are opt-in: at the adversarial distribution
#: they were the ones that took minutes, which is the finding, not a reason to
#: run them on every invocation.
DEFAULT_SIZES = (500, 2000, 6000)
LARGE_SIZES = (12000, 24000)

#: A paragraph the cleaner removes on text alone, and one it must keep.
NOTE = "Retain or delete manufacturers below."
REQUIREMENT = "Provide listed sprinklers with a 155 deg F temperature rating."
HEADING = "PART 1 - GENERAL"


# =============================================================================
# Fixture families
# =============================================================================

def _numbered(text: str, index: int) -> str:
    """A requirement that is unique in the document."""
    return db.text_para(f"{text} Item {index}.")


def adversarial_headings(size: int) -> list[str]:
    """Family 1: the original repeated-heading distribution.

    §2.4 and §16.1 both insist on the label: this is the *adversarial* case.
    Every third paragraph is the same string, so the differ's alignment has
    nothing unique to anchor on — which is what made it super-quadratic, and
    is not what a real specification looks like.
    """
    body = []
    for index in range(size):
        if index % 3 == 0:
            body.append(db.text_para(HEADING))
        elif index % 3 == 1:
            body.append(db.text_para(NOTE))
        else:
            body.append(db.text_para(REQUIREMENT))
    return body


def realistic_requirements(size: int) -> list[str]:
    """Family 2: mostly unique requirements, sparse editorial removals.

    The controlled contrast that isolated the duplicate effect, and the closest
    thing here to a real document.  §16.4 forbids gating on family 1 alone
    precisely so this one cannot regress unnoticed.
    """
    body = []
    for index in range(size):
        if index % 20 == 0:
            body.append(db.text_para(NOTE))
        else:
            body.append(_numbered(REQUIREMENT, index))
    return body


def repeated_requirements(size: int) -> list[str]:
    """Family 3: repeated identical headings *and* repeated identical requirements.

    Where text equality says least about identity, so the signature reasoning
    W04 added is doing the work.
    """
    body = []
    for index in range(size):
        if index % 4 == 0:
            body.append(db.text_para(HEADING))
        elif index % 4 == 1:
            body.append(db.text_para(REQUIREMENT))
        elif index % 4 == 2:
            body.append(db.text_para(REQUIREMENT))
        else:
            body.append(db.text_para(NOTE))
    return body


def no_anchors(size: int) -> list[str]:
    """Family 4: long runs of changed paragraphs with no unchanged anchors.

    Load-bearing, per §16.1: an anchor-based alignment has nothing to work with
    here, so duplicate-bucket evidence alone does not establish patience/LIS as
    the answer.  Every paragraph carries a placeholder, so every one is
    rewritten and none survives unchanged.
    """
    return [
        db.text_para(f"Provide [Insert quantity {index}] sprinklers at {index} ft.")
        for index in range(size)
    ]


def many_placeholders(size: int) -> list[str]:
    """Family 5: many inline placeholders, including very short survivors."""
    body = []
    for index in range(size):
        if index % 2 == 0:
            body.append(db.text_para(f"Provide [Verify quantity {index}] units."))
        else:
            # A survivor short enough that the similarity path would reject it.
            body.append(db.text_para(f"[Insert a long descriptive clause {index}] X."))
    return body


def long_paragraphs(size: int) -> list[str]:
    """Family 6: very long individual paragraphs.

    Deliberately separate from paragraph-count scaling: this varies characters
    per paragraph, not the number of them, so the two effects cannot be
    confused.  The count is capped for that reason.
    """
    count = max(4, size // 100)
    filler = " ".join(f"clause {n} of the requirement" for n in range(200))
    return [
        db.text_para(f"{REQUIREMENT} {filler} Item {index}.")
        for index in range(count)
    ]


def tables_and_parts(size: int) -> list[str]:
    """Family 7: tables and several content parts (the header comes separately)."""
    body = []
    for index in range(size // 2):
        body.append(db.table_of(
            db.row(db.text_para(f"{REQUIREMENT} Row {index}.")),
            db.row(db.text_para(NOTE)),
        ))
    return body


def damaged_output(size: int) -> list[str]:
    """Family 8: source for a deliberately damaged output.

    The output is built separately, dropping requirements policy never
    authorized losing.  A faster verification that stops reporting these has
    not been optimised, it has been broken.
    """
    return [_numbered(REQUIREMENT, index) for index in range(size)]


#: name -> (builder, needs_extra_parts)
FAMILIES = {
    "adversarial_headings": adversarial_headings,
    "realistic_requirements": realistic_requirements,
    "repeated_requirements": repeated_requirements,
    "no_anchors": no_anchors,
    "many_placeholders": many_placeholders,
    "long_paragraphs": long_paragraphs,
    "tables_and_parts": tables_and_parts,
    "damaged_output": damaged_output,
}


# =============================================================================
# Measurement
# =============================================================================

@dataclass
class Measurement:
    """One family at one size: what it contains and how long it took."""

    family: str
    size: int
    input_paragraphs: int = 0
    output_paragraphs: int = 0
    removed: int = 0
    modified: int = 0
    clean_seconds: list[float] = field(default_factory=list)
    verify_seconds: list[float] = field(default_factory=list)
    #: Semantic content of the output, never its ZIP bytes.
    output_digest: str = ""
    fields: int = 0
    references: int = 0
    passed: bool = False
    categories: list[str] = field(default_factory=list)

    @property
    def clean_median(self) -> float:
        return statistics.median(self.clean_seconds) if self.clean_seconds else 0.0

    @property
    def verify_median(self) -> float:
        return statistics.median(self.verify_seconds) if self.verify_seconds else 0.0


def semantic_digest(path: Path, engine: DetectionEngine) -> str:
    """A digest of what the document *says*, not of the archive holding it.

    §16.1 is explicit that ZIP bytes are not the behaviour oracle: they change
    with compression and ordering while the content is identical, and stay
    identical while a field carrier inside disappears.
    """
    texts = [info.raw_text for info in extract_paragraphs(path, engine, evidence=False)]
    digest = hashlib.sha256()
    for text in texts:
        digest.update(text.encode("utf-8"))
        digest.update(b"\\x00")
    return digest.hexdigest()[:16]


def carrier_counts(path: Path) -> tuple[int, int]:
    """Field carriers and reference consumers surviving in the output."""
    unpacked = Path(mkdtemp(prefix="speccleanse_bench_"))
    try:
        shutil.unpack_archive(str(path), str(unpacked), "zip")
        fields = references = 0
        for part in collect_content_parts(unpacked / "word"):
            root = parse_xml(part).getroot()
            fields += sum(field_instructions(root).values())
            references += len(reference_consumers(root))
        return fields, references
    finally:
        shutil.rmtree(unpacked, ignore_errors=True)


def build_fixture(family: str, size: int, directory: Path) -> tuple[Path, Path | None]:
    """Write one family's source document, and its damaged output if it has one."""
    body = FAMILIES[family](size)
    extra = None
    if family == "tables_and_parts":
        # A second content part, so the per-location comparison is exercised.
        extra = {"word/header1.xml": db.header(db.text_para(HEADING))}

    source = db.build_docx(directory / f"{family}_{size}.docx",
                           db.document(*body), extra)

    if family != "damaged_output":
        return source, None

    # Built independently, never by running the cleaner: the point is an output
    # verification must still reject.
    kept = body[: len(body) // 2]
    damaged = db.build_docx(directory / f"{family}_{size}_damaged.docx",
                            db.document(*kept), extra)
    return source, damaged


def measure(family: str, size: int, engine: DetectionEngine, repeat: int,
            directory: Path) -> Measurement:
    """Time one family at one size, repeating unprofiled runs for a median."""
    result = Measurement(family=family, size=size)
    source, damaged = build_fixture(family, size, directory)

    for _ in range(repeat):
        output = directory / f"{family}_{size}_out.docx"
        processor = DocxProcessor(engine, verbose=False)

        started = time.perf_counter()
        processed = processor.process(source, output)
        result.clean_seconds.append(time.perf_counter() - started)
        if not processed.success:
            raise RuntimeError(f"{family}/{size}: clean failed — {processed.errors}")

        judged = damaged if damaged is not None else output
        started = time.perf_counter()
        verification = verify_clean(source, judged, engine=engine)
        result.verify_seconds.append(time.perf_counter() - started)

    result.input_paragraphs = verification.input_paragraph_count
    result.output_paragraphs = verification.output_paragraph_count
    result.removed = len(verification.removed)
    result.modified = len(verification.modified)
    result.passed = verification.passed
    result.categories = sorted(c.value for c in verification.review_categories())
    result.output_digest = semantic_digest(judged, engine)
    result.fields, result.references = carrier_counts(judged)

    if family == "damaged_output" and verification.passed:
        raise RuntimeError(
            "damaged_output verified clean — verification has been broken, "
            "not optimised"
        )
    return result


def environment() -> dict:
    """What the numbers are only comparable within."""
    return {
        "python": sys.version.split()[0],
        "implementation": platform.python_implementation(),
        "platform": platform.platform(),
        "machine": platform.machine(),
    }


def run(families: list[str], sizes: list[int], repeat: int,
        directory: Path) -> dict:
    engine = DetectionEngine(load_config(resolve_config_path()))
    measurements = [
        measure(family, size, engine, repeat, directory)
        for family in families
        for size in sizes
    ]
    return {
        "environment": environment(),
        "repeat": repeat,
        "measurements": [asdict(m) for m in measurements],
    }


def format_report(report: dict) -> str:
    """A table, plus the environment the numbers belong to."""
    lines = [
        "SpecCleanse pipeline benchmark",
        "",
        f"  python {report['environment']['python']}"
        f" ({report['environment']['implementation']})"
        f" on {report['environment']['platform']}",
        f"  median of {report['repeat']} unprofiled run(s)",
        "",
        f"  {'family':24} {'size':>7} {'paras':>7} {'clean s':>9} {'verify s':>9}"
        f" {'removed':>8} {'modified':>9}  verdict",
        f"  {'-' * 24} {'-' * 7} {'-' * 7} {'-' * 9} {'-' * 9} {'-' * 8} {'-' * 9}  {'-' * 7}",
    ]
    for row in report["measurements"]:
        clean = statistics.median(row["clean_seconds"])
        verify = statistics.median(row["verify_seconds"])
        verdict = "pass" if row["passed"] else ",".join(
            c.split()[0] for c in row["categories"]
        ) or "fail"
        lines.append(
            f"  {row['family']:24} {row['size']:>7} {row['input_paragraphs']:>7}"
            f" {clean:>9.3f} {verify:>9.3f} {row['removed']:>8} {row['modified']:>9}"
            f"  {verdict}"
        )
    lines.append("")
    lines.append("  Timings are comparable only within one machine, interpreter and")
    lines.append("  configuration.  adversarial_headings is the worst case by")
    lines.append("  construction, not the representative one; realistic_requirements")
    lines.append("  is the closest thing here to a real document.")
    return "\n".join(lines)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Time the clean-and-verify pipeline on synthetic fixtures.",
    )
    parser.add_argument("--family", action="append", choices=sorted(FAMILIES),
                        help="Measure only this family; repeatable.")
    parser.add_argument("--sizes", default=",".join(str(s) for s in DEFAULT_SIZES),
                        help="Comma-separated paragraph counts.")
    parser.add_argument("--large", action="store_true",
                        help=f"Also measure {LARGE_SIZES}, which are slow by design.")
    parser.add_argument("--repeat", type=int, default=3,
                        help="Unprofiled runs per point; the median is reported.")
    parser.add_argument("--out", type=Path,
                        help="Directory to write benchmark.json into.")
    args = parser.parse_args(argv)

    sizes = [int(s) for s in args.sizes.split(",") if s.strip()]
    if args.large:
        sizes.extend(LARGE_SIZES)
    families = args.family or sorted(FAMILIES)

    directory = Path(mkdtemp(prefix="speccleanse_bench_"))
    try:
        report = run(families, sizes, args.repeat, directory)
    finally:
        shutil.rmtree(directory, ignore_errors=True)

    print(format_report(report))
    if args.out:
        args.out.mkdir(parents=True, exist_ok=True)
        target = args.out / "benchmark.json"
        target.write_text(json.dumps(report, indent=2), encoding="utf-8")
        print(f"\n  wrote {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
