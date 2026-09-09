#!/usr/bin/env python3
"""
SpecCleanse GUI

Tkinter-based graphical interface for SpecCleanse.
Runs single-pass content removal and verification on one or more DOCX files.
"""

import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from apppaths import resolve_config_path
from batch import BatchItem, BatchPlan, FileOutcome, plan_batch, summarise
from detection import DetectionEngine, ContentType, config_notices
from docx_xml import load_config
from processor import DocxProcessor, ProcessingResult
from verify import verify_clean


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Where patterns.yaml lives depends on how SpecCleanse is running: beside the
# modules from source, beside the .exe or under %APPDATA% when frozen. See
# apppaths.resolve_config_path.
CONFIG_PATH = resolve_config_path()


def build_engine() -> DetectionEngine:
    """Load patterns.yaml and build the detection engine.

    Raises FileNotFoundError or ValueError with a readable message if the
    configuration is missing, empty, malformed, or holds an invalid regex.
    """
    return DetectionEngine(load_config(CONFIG_PATH))


def _shorten(text: str, width: int = 90) -> str:
    """One-line preview of a paragraph, trimmed to ``width`` characters."""
    flat = " ".join(text.split())
    return flat if len(flat) <= width else flat[:width] + "..."


def _preview_one(
    input_path: Path, engine: DetectionEngine, log, strip_revisions: bool = False
) -> bool:
    """Run a dry-run preview on a single file and log detections."""
    processor = DocxProcessor(
        engine, verbose=False, dry_run=True, strip_revisions=strip_revisions
    )

    temp_dir = Path(tempfile.mkdtemp(prefix="speccleanse_preview_"))
    preview_output = temp_dir / f"{input_path.stem}_preview.docx"

    try:
        result: ProcessingResult = processor.process(
            input_path=input_path,
            output_path=preview_output,
        )

        if not result.success:
            for err in result.errors:
                log(f"  ERROR: {err}")
            return False

        removed, redacted, preserved = _group_detections(result.detections)

        log(f"  Would remove {sum(len(v) for v in removed.values())} items,"
            f" redact {len(redacted)} inline placeholder(s),"
            f" preserve {len(preserved)}")

        for category in sorted(removed):
            detections = removed[category]
            log(f"\n  REMOVALS — {category} ({len(detections)}):")
            for d in detections:
                label = f"{d.confidence:.2f}"
                if d.formatting_only:
                    label += ", formatting-only"
                log(f"    [{label}] \"{_shorten(d.text)}\"")

        if redacted:
            log(f"\n  INLINE REDACTIONS ({len(redacted)}):")
            for d in redacted:
                cuts = ", ".join(d.text[start:end] for start, end in d.spans)
                log(f"    cut {_shorten(cuts, 60)!r} from \"{_shorten(d.text)}\"")

        if preserved:
            log(f"\n  PRESERVED ({len(preserved)}):")
            for d in preserved:
                log(f"    [preserve, {d.confidence:.2f}] \"{_shorten(d.text)}\"")

        return True

    except Exception as exc:
        log(f"  FAILED: {exc}")
        return False

    finally:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)


def _group_detections(detections):
    """Split detections into removals by category, redactions, and preserves."""
    removed: dict[str, list] = {}
    redacted = []
    preserved = []

    for d in detections:
        if d.content_type == ContentType.PRESERVE:
            preserved.append(d)
        elif d.content_type == ContentType.INLINE_PLACEHOLDER:
            redacted.append(d)
        else:
            removed.setdefault(d.content_type.value, []).append(d)

    return removed, redacted, preserved


def _clean_one(
    input_path: Path,
    output_path: Path,
    engine: DetectionEngine,
    log,
    strip_revisions: bool = False,
) -> FileOutcome:
    """Run single-pass content removal on a single file.

    Writing the output and verifying it are separate outcomes.  A file whose
    verification reported a preserve violation was written successfully and is
    still not something to hand on unread, so it is neither a success nor a
    failure: it needs review, and the caller is told which.
    """
    processor = DocxProcessor(engine, verbose=False, strip_revisions=strip_revisions)

    log("  Content removal...")
    result: ProcessingResult = processor.process(
        input_path=input_path,
        output_path=output_path,
    )

    if not result.success:
        for err in result.errors:
            log(f"  ERROR: {err}")
        return FileOutcome.FAILED

    removed, redacted, preserved = _group_detections(result.detections)
    log(f"    Removed {sum(len(v) for v in removed.values())} items,"
        f" redacted {len(redacted)} inline placeholder(s),"
        f" preserved {len(preserved)}")

    # Things kept on purpose, where removing them would have gone beyond
    # removing content.  Not failures, and not a reason to withhold the file.
    for warning in result.warnings:
        log(f"    NOTE: {warning}")

    log("  Checking the output against the rules that produced it...")
    try:
        vresult = verify_clean(
            input_path, output_path, engine=engine, strip_revisions=strip_revisions
        )
    except Exception as exc:
        # The file was written; only the check failed.  Saying nothing was
        # produced would be false, and hiding the path would leave an
        # unverified document sitting in the output folder unannounced.
        log(f"  FAILED: the output could not be verified: {exc}")
        if output_path.exists():
            log(f"  The cleaned file was written but is UNVERIFIED: {output_path}")
        return FileOutcome.FAILED

    _log_verification(vresult, log)

    log(f"  Done: {len(vresult.removed)} paragraph(s) removed,"
        f" {vresult.removed_characters:,} characters of text taken out")

    if vresult.passed:
        return FileOutcome.VERIFIED

    log(f"  NEEDS REVIEW — the cleaned file was written: {output_path}")
    return FileOutcome.NEEDS_REVIEW


def _log_verification(vresult, log) -> None:
    """Print the verification report: removals, modifications, structure."""
    n_unexpected = len(vresult.unexpected_removals)
    n_preserve = len(vresult.preserve_violations)

    log(f"    Paragraphs removed: {len(vresult.removed)}"
        f" ({len(vresult.expected_removals)} expected, {n_unexpected} unexpected"
        f", {n_preserve} preserve violations)")

    if vresult.modified:
        log(f"    Paragraphs modified: {len(vresult.modified)}"
            f" ({len(vresult.expected_modifications)} expected,"
            f" {len(vresult.unexpected_modifications)} unexpected)")

    if vresult.passed:
        # What this actually establishes: every difference between input and
        # output was accounted for by a configured rule, and the structural
        # checks found nothing the input did not already have.  It is not a
        # statement that the document is correct, nor that Word will open it.
        log("    PASS — no unexplained text changes or new checked structural "
            "problems were found")
        return

    if vresult.structural:
        log(f"    FAIL — {len(vresult.structural)} structural problem(s) "
            "the output has and the input did not:")
        for violation in vresult.structural:
            log(f"      {violation}")

    if n_preserve:
        log(f"    FAIL — {n_preserve} preserve violation(s) "
            "(content that should NEVER be removed):")
        for r in vresult.preserve_violations:
            log(f"      \"{_shorten(r.text)}\"")
            log(f"        matched: {r.pattern_matched}")

    if n_unexpected:
        log(f"    WARN — {n_unexpected} removal(s) may be real content:")
        for r in vresult.unexpected_removals:
            log(f"      \"{_shorten(r.text)}\"")

    if vresult.numbering:
        # Its own category, and worded as what it is.  Nothing was lost here;
        # what may have changed is the numbers a reader sees.  Saying so
        # separately keeps it from being read as damage.
        log(f"    NOTE — {len(vresult.numbering)} removed paragraph(s) took part "
            "in automatic numbering; displayed numbering or references to it "
            "may change:")
        for notice in vresult.numbering:
            log(f"      {notice}")

    if vresult.unexpected_modifications:
        log(f"    WARN — {len(vresult.unexpected_modifications)} paragraph(s) "
            "lost text no rule accounts for:")
        for m in vresult.unexpected_modifications:
            log(f"      \"{_shorten(m.before)}\"")
            for fragment in m.fragments:
                log(f"        lost: {_shorten(fragment, 60)!r}")

    if vresult.added:
        log(f"    FAIL — {len(vresult.added)} paragraph(s) in the output "
            "were not in the input:")
        for text in vresult.added:
            log(f"      \"{_shorten(text)}\"")


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

# Colours / theme
BG        = "#1e1e2e"
BG_LIGHT  = "#313244"
FG        = "#cdd6f4"
FG_DIM    = "#6c7086"
ACCENT    = "#89b4fa"
GREEN     = "#a6e3a1"
RED       = "#f38ba8"
YELLOW    = "#f9e2af"
SURFACE   = "#45475a"


class SpecCleanseGUI:
    """Main application window."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("SpecCleanse")
        self.root.configure(bg=BG)
        self.root.minsize(720, 520)

        self.files: list[Path] = []
        self.output_dir: Path | None = None
        self._running = False

        # The worker thread writes log lines here; the main loop drains them.
        self._log_queue: queue.Queue[str] = queue.Queue()

        self._build_ui()
        self._drain_log()

    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=FG, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=BG, foreground=ACCENT,
                         font=("Segoe UI", 18, "bold"))
        style.configure("Sub.TLabel", background=BG, foreground=FG_DIM,
                         font=("Segoe UI", 9))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Clean.TButton", font=("Segoe UI", 12, "bold"))

        style.configure("green.Horizontal.TProgressbar",
                         troughcolor=BG_LIGHT, background=GREEN)

        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="SpecCleanse", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Removes editorial notes, specifier comments, copyright boilerplate, "
                 "hidden text, and editing instructions from specification documents.",
            style="Sub.TLabel",
            wraplength=680,
        ).pack(anchor="w", pady=(0, 12))

        file_frame = ttk.Frame(outer)
        file_frame.pack(fill="x", pady=(0, 4))

        self.btn_add = tk.Button(
            file_frame, text="Add Files...", command=self._add_files,
            bg=ACCENT, fg=BG, activebackground=GREEN, activeforeground=BG,
            font=("Segoe UI", 10, "bold"), relief="flat", padx=14, pady=4,
        )
        self.btn_add.pack(side="left")

        self.btn_remove = tk.Button(
            file_frame, text="Remove Selected", command=self._remove_selected,
            bg=SURFACE, fg=FG, activebackground=ACCENT, activeforeground=BG,
            font=("Segoe UI", 10), relief="flat", padx=10, pady=4,
        )
        self.btn_remove.pack(side="left", padx=(8, 0))

        self.btn_clear = tk.Button(
            file_frame, text="Clear", command=self._clear_files,
            bg=SURFACE, fg=FG, activebackground=RED, activeforeground=BG,
            font=("Segoe UI", 10), relief="flat", padx=10, pady=4,
        )
        self.btn_clear.pack(side="left", padx=(8, 0))

        self.lbl_count = ttk.Label(file_frame, text="No files selected", style="Sub.TLabel")
        self.lbl_count.pack(side="left", padx=(12, 0))

        list_frame = ttk.Frame(outer)
        list_frame.pack(fill="both", expand=False, pady=(0, 8))

        self.file_listbox = tk.Listbox(
            list_frame, height=5, selectmode="extended",
            bg=BG_LIGHT, fg=FG, selectbackground=ACCENT, selectforeground=BG,
            font=("Consolas", 9), relief="flat", borderwidth=0,
            highlightthickness=1, highlightcolor=SURFACE, highlightbackground=SURFACE,
        )
        self.file_listbox.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        sb.pack(side="right", fill="y")
        self.file_listbox.configure(yscrollcommand=sb.set)

        out_frame = ttk.Frame(outer)
        out_frame.pack(fill="x", pady=(0, 8))

        self.btn_outdir = tk.Button(
            out_frame, text="Output Folder...", command=self._pick_output_dir,
            bg=SURFACE, fg=FG, activebackground=ACCENT, activeforeground=BG,
            font=("Segoe UI", 10), relief="flat", padx=10, pady=4,
        )
        self.btn_outdir.pack(side="left")

        self.btn_open_outdir = tk.Button(
            out_frame, text="Open Output Folder", command=self._open_output_dir,
            bg=SURFACE, fg=FG, activebackground=ACCENT, activeforeground=BG,
            font=("Segoe UI", 10), relief="flat", padx=10, pady=4,
        )
        self.btn_open_outdir.pack(side="left", padx=(8, 0))

        self.lbl_outdir = ttk.Label(
            out_frame, text="Default: same folder as input, with _cleaned suffix",
            style="Sub.TLabel",
        )
        self.lbl_outdir.pack(side="left", padx=(12, 0))

        option_frame = ttk.Frame(outer)
        option_frame.pack(fill="x", pady=(0, 8))

        self.strip_revisions = tk.BooleanVar(value=False)
        self.chk_revisions = tk.Checkbutton(
            option_frame,
            text="Strip comments and accept tracked changes",
            variable=self.strip_revisions,
            bg=BG, fg=FG, selectcolor=BG_LIGHT, activebackground=BG,
            activeforeground=FG, disabledforeground=FG_DIM,
            font=("Segoe UI", 9), relief="flat", highlightthickness=0,
            anchor="w",
        )
        self.chk_revisions.pack(side="left")

        action_frame = ttk.Frame(outer)
        action_frame.pack(pady=(4, 8))

        self.btn_preview = tk.Button(
            action_frame, text="Preview", command=self._start_preview,
            bg=SURFACE, fg=FG, activebackground=ACCENT, activeforeground=BG,
            font=("Segoe UI", 11, "bold"), relief="flat", padx=18, pady=6,
            disabledforeground=FG_DIM,
        )
        self.btn_preview.pack(side="left", padx=(0, 8))

        self.btn_clean = tk.Button(
            action_frame, text="CLEAN", command=self._start_clean,
            bg=GREEN, fg=BG, activebackground=ACCENT, activeforeground=BG,
            font=("Segoe UI", 14, "bold"), relief="flat", padx=24, pady=6,
            disabledforeground=FG_DIM,
        )
        self.btn_clean.pack(side="left")

        self.progress = ttk.Progressbar(
            outer, mode="determinate", style="green.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x", pady=(0, 4))

        self.lbl_status = ttk.Label(outer, text="Ready", style="Sub.TLabel")
        self.lbl_status.pack(anchor="w")

        log_frame = ttk.Frame(outer)
        log_frame.pack(fill="both", expand=True, pady=(4, 0))

        self.log_text = tk.Text(
            log_frame, height=12, wrap="word",
            bg=BG_LIGHT, fg=FG, insertbackground=FG,
            font=("Consolas", 9), relief="flat", borderwidth=0,
            highlightthickness=1, highlightcolor=SURFACE, highlightbackground=SURFACE,
            state="disabled",
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        log_sb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_sb.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=log_sb.set)

        self._log(f"Patterns: {CONFIG_PATH}")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select DOCX files",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        for p in paths:
            pp = Path(p)
            if pp not in self.files:
                self.files.append(pp)
                self.file_listbox.insert("end", str(pp))
        self._update_count()

    def _remove_selected(self):
        for index in sorted(self.file_listbox.curselection(), reverse=True):
            self.file_listbox.delete(index)
            del self.files[index]
        self._update_count()

    def _clear_files(self):
        self.files.clear()
        self.file_listbox.delete(0, "end")
        self._update_count()

    def _update_count(self):
        n = len(self.files)
        self.lbl_count.configure(
            text=f"{n} file{'s' if n != 1 else ''} selected" if n else "No files selected"
        )

    def _pick_output_dir(self):
        d = filedialog.askdirectory(title="Choose output folder")
        if d:
            self.output_dir = Path(d)
            self.lbl_outdir.configure(text=str(self.output_dir))

    def _open_output_dir(self):
        """Open the folder the cleaned files go to in the system file browser."""
        folder = self.output_dir
        if folder is None and self.files:
            folder = self.files[0].parent
        if folder is None:
            self._log("No output folder yet — add a file or choose one.")
            return
        if not folder.exists():
            self._log(f"Folder does not exist: {folder}")
            return

        try:
            if sys.platform == "win32":
                os.startfile(folder)  # noqa: S606 - the platform's own file browser
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except OSError as exc:
            self._log(f"Could not open {folder}: {exc}")

    def _log(self, text: str):
        """Queue a line for the log; safe to call from the worker thread."""
        self._log_queue.put(text)

    def _drain_log(self):
        """Move queued log lines into the widget, on the main thread."""
        lines: list[str] = []
        while True:
            try:
                lines.append(self._log_queue.get_nowait())
            except queue.Empty:
                break

        if lines:
            self.log_text.configure(state="normal")
            self.log_text.insert("end", "\n".join(lines) + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")

        self.root.after(100, self._drain_log)

    def _set_status(self, text: str):
        self.root.after(0, lambda: self.lbl_status.configure(text=text))

    def _set_progress(self, value: float):
        self.root.after(0, lambda: self.progress.configure(value=value))

    def _disable_controls(self):
        self._running = True
        for button in self._run_controls():
            button.configure(state="disabled")

    def _enable_controls(self):
        self._running = False
        for button in self._run_controls():
            button.configure(state="normal")

    def _run_controls(self) -> list[tk.Widget]:
        """Controls that must not be usable while a run is in flight.

        Everything that feeds a run is included: each run works from the
        selection, destination, and options it started with, so leaving these
        live would only let a mid-run change look like it took effect.
        """
        return [
            self.btn_preview,
            self.btn_clean,
            self.btn_add,
            self.btn_remove,
            self.btn_clear,
            self.btn_outdir,
            self.chk_revisions,
        ]

    def _clear_log(self):
        while True:
            try:
                self._log_queue.get_nowait()
            except queue.Empty:
                break
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _start_clean(self):
        if self._running:
            return
        if not self.files:
            self._log("No files selected. Click 'Add Files...' first.")
            return

        # Snapshot the inputs here, on the main thread: the run should finish
        # against the selection, destination, and options it started with.
        files = list(self.files)
        output_dir = self.output_dir
        strip_revisions = self.strip_revisions.get()

        # Work out every destination and check the whole set before opening a
        # single file.  Two selected documents that share a basename land on one
        # destination when a common output folder is chosen, and the second
        # clean would silently replace the first.
        plan = plan_batch(files, output_dir)
        if not plan.ok:
            self._reject_batch(plan)
            return

        if not self._confirm_overwrite(plan.existing_outputs):
            return

        self._disable_controls()
        self._clear_log()
        threading.Thread(
            target=self._run_clean,
            args=(plan.items, strip_revisions),
            daemon=True,
        ).start()

    def _reject_batch(self, plan: BatchPlan) -> None:
        """Refuse a batch whose destinations collide, before anything is written."""
        self._clear_log()
        self._log("Cannot start: the selected files do not have separate "
                  "destinations.  Nothing was written.")
        for conflict in plan.conflicts:
            self._log("")
            self._log("  " + conflict.describe().replace("\n", "\n  "))
        self._log("")
        self._log("Choose a different output folder, or clean these files in "
                  "separate runs.")
        self._set_status("Batch rejected — destinations collide")
        messagebox.showerror(
            "Conflicting destinations",
            f"{len(plan.conflicts)} destination conflict(s) would cause a "
            "cleaned file to be overwritten or an input destroyed.\n\n"
            "Nothing was written.  See the log for the files involved.",
            parent=self.root,
        )

    def _confirm_overwrite(self, existing: list[Path]) -> bool:
        """Ask before replacing cleaned files from an earlier run.

        Reached only once the batch's own destinations are known to be
        distinct: there is nothing to confirm about a run that will not happen.
        """
        if not existing:
            return True

        listed = "\n".join(f"  {out.name}" for out in existing[:8])
        if len(existing) > 8:
            listed += f"\n  ...and {len(existing) - 8} more"

        confirmed = messagebox.askyesno(
            "Overwrite existing files?",
            f"{len(existing)} cleaned file(s) already exist and will be "
            f"replaced:\n\n{listed}",
            parent=self.root,
        )
        if not confirmed:
            self._log("Cancelled — nothing was written.")
        return confirmed

    def _start_preview(self):
        if self._running:
            return
        if not self.files:
            self._log("No files selected. Click 'Add Files...' first.")
            return

        files = list(self.files)
        strip_revisions = self.strip_revisions.get()

        self._disable_controls()
        self._clear_log()
        threading.Thread(
            target=self._run_preview, args=(files, strip_revisions), daemon=True
        ).start()

    def _load_engine(self) -> DetectionEngine | None:
        """Build the detection engine once per run, reporting config errors.

        Configuration problems used to surface as a stderr traceback — invisible
        under pythonw.exe — while the buttons stayed disabled forever.
        """
        try:
            engine = build_engine()
        except Exception as exc:
            self._log(f"Configuration error in {CONFIG_PATH}: {exc}")
            self._log("Fix the file and try again — nothing was processed.")
            self._set_status("Configuration error")
            return None

        # Every run starts by clearing the log, so the startup line is already
        # gone. A finished run's log has to say which patterns produced it.
        self._log(f"Patterns: {CONFIG_PATH}")

        # An installed patterns.yaml is never overwritten by an update, so a
        # copy made before a rule was narrowed keeps the old, broader rule and
        # nothing would otherwise say so.
        for notice in config_notices(engine.config):
            self._log(f"  NOTE: {notice}")

        return engine

    def _run_preview(self, files: list[Path], strip_revisions: bool = False):
        try:
            engine = self._load_engine()
            if engine is None:
                return

            total = len(files)
            successes = 0
            failures = 0

            for i, fpath in enumerate(files, 1):
                self._set_status(f"Previewing {i}/{total}: {fpath.name}")
                self._set_progress((i - 1) / total * 100)
                self._log(f"[{i}/{total}] {fpath.name} — Preview")

                ok = _preview_one(fpath, engine, self._log, strip_revisions)
                if ok:
                    successes += 1
                else:
                    failures += 1

                self._log("")

            self._set_progress(100)

            summary = f"Preview done: {successes} succeeded"
            if failures:
                summary += f", {failures} failed"
            self._set_status(summary)
            self._log("=" * 50)
            self._log(summary)

        except Exception as exc:
            self._log(f"Preview stopped: {exc}")
            self._set_status("Preview stopped — see log")

        finally:
            self.root.after(0, self._enable_controls)

    def _run_clean(
        self, items: list[BatchItem], strip_revisions: bool = False
    ):
        try:
            engine = self._load_engine()
            if engine is None:
                return

            total = len(items)
            counts = {outcome: 0 for outcome in FileOutcome}

            # The destinations were worked out and validated before this
            # thread started.  They are not recomputed here: the selection and
            # output folder are live widgets the user can change mid-run.
            for i, item in enumerate(items, 1):
                self._set_status(f"Cleaning {i}/{total}: {item.source.name}")
                self._set_progress((i - 1) / total * 100)
                self._log(f"[{i}/{total}] {item.source.name}")

                outcome = _clean_one(
                    item.source, item.destination, engine, self._log, strip_revisions
                )
                counts[outcome] += 1
                if outcome is not FileOutcome.FAILED:
                    self._log(f"  -> {item.destination.name}")

                self._log("")

            self._set_progress(100)

            summary = summarise(counts)
            self._set_status(summary)
            self._log("=" * 50)
            self._log(summary)

        except Exception as exc:
            self._log(f"Clean stopped: {exc}")
            self._set_status("Clean stopped — see log")

        finally:
            self.root.after(0, self._enable_controls)


def main():
    root = tk.Tk()
    SpecCleanseGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
