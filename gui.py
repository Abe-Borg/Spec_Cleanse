#!/usr/bin/env python3
"""
SpecCleanse GUI

Tkinter-based graphical interface for SpecCleanse.
Runs single-pass content removal and verification on one or more DOCX files.
"""

import shutil
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

import yaml

from detection import DetectionEngine, ContentType
from processor import DocxProcessor, ProcessingResult
from verify import verify_clean


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

CONFIG_PATH = Path(__file__).parent / "patterns.yaml"


def load_config() -> dict:
    with open(CONFIG_PATH, "r") as f:
        return yaml.safe_load(f)


def _preview_one(input_path: Path, log) -> bool:
    """Run a dry-run preview on a single file and log detections."""
    config = load_config()
    engine = DetectionEngine(config)
    processor = DocxProcessor(engine, verbose=False, dry_run=True)

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

        removed = [d for d in result.detections if d.content_type != ContentType.PRESERVE]
        preserved = [d for d in result.detections if d.content_type == ContentType.PRESERVE]

        log(f"  Would remove {len(removed)} items, preserve {len(preserved)}")

        if removed:
            log("\n  REMOVALS:")
            for d in removed:
                preview = d.text.replace("\n", " ").strip()
                if len(preview) > 90:
                    preview = preview[:90] + "..."
                log(f"    [{d.content_type.value}, {d.confidence:.2f}] \"{preview}\"")

        if preserved:
            log("\n  PRESERVED:")
            for d in preserved:
                preview = d.text.replace("\n", " ").strip()
                if len(preview) > 90:
                    preview = preview[:90] + "..."
                log(f"    [preserve, {d.confidence:.2f}] \"{preview}\"")

        return True

    except Exception as exc:
        log(f"  FAILED: {exc}")
        return False

    finally:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)


def _clean_one(input_path: Path, output_path: Path, log) -> bool:
    """Run single-pass content removal on a single file."""
    config = load_config()
    engine = DetectionEngine(config)
    processor = DocxProcessor(engine, verbose=False)

    try:
        log("  Content removal...")
        result: ProcessingResult = processor.process(
            input_path=input_path,
            output_path=output_path,
        )

        if not result.success:
            for err in result.errors:
                log(f"  ERROR: {err}")
            return False

        removed = [d for d in result.detections if d.content_type != ContentType.PRESERVE]
        preserved = [d for d in result.detections if d.content_type == ContentType.PRESERVE]
        log(f"    Removed {len(removed)} items, preserved {len(preserved)}")

        original_size = input_path.stat().st_size
        final_size = output_path.stat().st_size
        saved = original_size - final_size
        pct = (saved / original_size * 100) if original_size else 0

        log("  Verifying no spec content was lost...")
        vresult = verify_clean(input_path, output_path)
        n_expected = len(vresult.expected_removals)
        n_unexpected = len(vresult.unexpected_removals)
        n_preserve = len(vresult.preserve_violations)
        log(f"    Paragraphs removed: {len(vresult.removed)}"
            f" ({n_expected} expected, {n_unexpected} unexpected"
            f", {n_preserve} preserve violations)")
        if vresult.passed:
            log("    PASS — all removals match known bloat patterns")
        else:
            if n_preserve:
                log(f"    FAIL — {n_preserve} preserve violation(s) "
                    "(content that should NEVER be removed):")
                for r in vresult.preserve_violations:
                    preview = r.text[:90] + "..." if len(r.text) > 90 else r.text
                    preview = preview.replace("\n", " ")
                    log(f"      \"{preview}\"")
                    log(f"        matched: {r.pattern_matched}")
            if n_unexpected:
                log(f"    WARN — {n_unexpected} removal(s) may be real content:")
            for r in vresult.unexpected_removals:
                preview = r.text[:90] + "..." if len(r.text) > 90 else r.text
                preview = preview.replace("\n", " ")
                log(f"      \"{preview}\"")

        log(f"  Done: {original_size:,} -> {final_size:,} bytes ({pct:.1f}% smaller)")
        return True

    except Exception as exc:
        log(f"  FAILED: {exc}")
        return False


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

        self._build_ui()

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
            list_frame, height=5,
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

        self.lbl_outdir = ttk.Label(
            out_frame, text="Default: same folder as input, with _cleaned suffix",
            style="Sub.TLabel",
        )
        self.lbl_outdir.pack(side="left", padx=(12, 0))

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

        self.log_text = tk.Text(
            outer, height=12, wrap="word",
            bg=BG_LIGHT, fg=FG, insertbackground=FG,
            font=("Consolas", 9), relief="flat", borderwidth=0,
            highlightthickness=1, highlightcolor=SURFACE, highlightbackground=SURFACE,
            state="disabled",
        )
        self.log_text.pack(fill="both", expand=True, pady=(4, 0))

        log_sb = ttk.Scrollbar(outer, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)

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

    def _output_for(self, input_path: Path) -> Path:
        stem = input_path.stem + "_cleaned"
        parent = self.output_dir if self.output_dir else input_path.parent
        return parent / (stem + ".docx")

    def _log(self, text: str):
        def _append():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", text + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(0, _append)

    def _set_status(self, text: str):
        self.root.after(0, lambda: self.lbl_status.configure(text=text))

    def _set_progress(self, value: float):
        self.root.after(0, lambda: self.progress.configure(value=value))

    def _disable_controls(self):
        self._running = True
        self.btn_preview.configure(state="disabled")
        self.btn_clean.configure(state="disabled")
        self.btn_add.configure(state="disabled")
        self.btn_clear.configure(state="disabled")

    def _enable_controls(self):
        self._running = False
        self.btn_preview.configure(state="normal")
        self.btn_clean.configure(state="normal")
        self.btn_add.configure(state="normal")
        self.btn_clear.configure(state="normal")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _start_clean(self):
        if self._running:
            return
        if not self.files:
            self._log("No files selected. Click 'Add Files...' first.")
            return

        self._disable_controls()
        self._clear_log()
        threading.Thread(target=self._run_clean, daemon=True).start()

    def _start_preview(self):
        if self._running:
            return
        if not self.files:
            self._log("No files selected. Click 'Add Files...' first.")
            return

        self._disable_controls()
        self._clear_log()
        threading.Thread(target=self._run_preview, daemon=True).start()

    def _run_preview(self):
        total = len(self.files)
        successes = 0
        failures = 0

        for i, fpath in enumerate(self.files, 1):
            self._set_status(f"Previewing {i}/{total}: {fpath.name}")
            self._set_progress((i - 1) / total * 100)
            self._log(f"[{i}/{total}] {fpath.name} — Preview")

            ok = _preview_one(fpath, self._log)
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

        self.root.after(0, self._enable_controls)

    def _run_clean(self):
        total = len(self.files)
        successes = 0
        failures = 0

        for i, fpath in enumerate(self.files, 1):
            self._set_status(f"Cleaning {i}/{total}: {fpath.name}")
            self._set_progress((i - 1) / total * 100)
            self._log(f"[{i}/{total}] {fpath.name}")

            out = self._output_for(fpath)
            ok = _clean_one(fpath, out, self._log)
            if ok:
                successes += 1
                self._log(f"  -> {out.name}")
            else:
                failures += 1

            self._log("")

        self._set_progress(100)

        summary = f"Done: {successes} succeeded"
        if failures:
            summary += f", {failures} failed"
        self._set_status(summary)
        self._log("=" * 50)
        self._log(summary)

        self.root.after(0, self._enable_controls)


def main():
    root = tk.Tk()
    SpecCleanseGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
