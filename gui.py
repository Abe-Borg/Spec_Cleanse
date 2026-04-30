#!/usr/bin/env python3
"""
SpecCleanse GUI

Tkinter-based graphical interface for SpecCleanse.

Three workflows:
  * Preview       — dry-run report of every detected removal candidate.
  * Inspect Notes — list each specifier note with its location and let the
                    user pick which ones to delete (per-file).
  * CLEAN         — single-pass content removal + verification on one or
                    more DOCX files.
"""

import shutil
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import yaml

from detection import DetectionEngine, ContentType
from notes import (
    SpecifierNote,
    extract_specifier_notes,
    remove_specifier_notes_at_locations,
)
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


class NotesReviewWindow:
    """Modal-ish window listing every detected specifier note with a checkbox.

    The user picks which notes to delete; on confirmation we write a fresh
    DOCX to disk via remove_specifier_notes_at_locations.  Word does not
    need to be open — the file is manipulated as a ZIP archive of XML.
    """

    def __init__(
        self,
        parent: tk.Tk,
        input_path: Path,
        notes: list[SpecifierNote],
        default_output: Path,
        logger,
    ):
        self.parent = parent
        self.input_path = input_path
        self.notes = notes
        self.output_path = default_output
        self.logger = logger
        self._vars: list[tk.BooleanVar] = []

        self.win = tk.Toplevel(parent)
        self.win.title(f"Specifier Notes — {input_path.name}")
        self.win.configure(bg=BG)
        self.win.minsize(780, 560)
        self.win.transient(parent)

        self._build_ui()

    def _build_ui(self):
        outer = ttk.Frame(self.win, padding=14)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text=f"Specifier notes in {self.input_path.name}",
            style="Header.TLabel",
        ).pack(anchor="w")

        ttk.Label(
            outer,
            text=(
                f"{len(self.notes)} note(s) detected. Tick the boxes for "
                "the notes you want to delete, then click Delete Selected. "
                "The original file is never modified — a new .docx is "
                "written to the output path below. Word does not need to "
                "be open (it must NOT be open on Windows, which holds a "
                "write lock on open files)."
            ),
            style="Sub.TLabel",
            wraplength=720,
            justify="left",
        ).pack(anchor="w", pady=(4, 10))

        # Selection toolbar
        bar = ttk.Frame(outer)
        bar.pack(fill="x", pady=(0, 6))
        tk.Button(
            bar, text="Select All", command=self._select_all,
            bg=SURFACE, fg=FG, activebackground=ACCENT, activeforeground=BG,
            font=("Segoe UI", 10), relief="flat", padx=10, pady=2,
        ).pack(side="left")
        tk.Button(
            bar, text="Select None", command=self._select_none,
            bg=SURFACE, fg=FG, activebackground=ACCENT, activeforeground=BG,
            font=("Segoe UI", 10), relief="flat", padx=10, pady=2,
        ).pack(side="left", padx=(6, 0))
        self.lbl_count = ttk.Label(bar, text="", style="Sub.TLabel")
        self.lbl_count.pack(side="right")

        # Scrollable list of checkboxes
        list_container = tk.Frame(
            outer, bg=BG_LIGHT,
            highlightthickness=1, highlightcolor=SURFACE,
            highlightbackground=SURFACE,
        )
        list_container.pack(fill="both", expand=True)

        canvas = tk.Canvas(
            list_container, bg=BG_LIGHT, highlightthickness=0,
        )
        canvas.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(
            list_container, orient="vertical", command=canvas.yview,
        )
        sb.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=sb.set)

        inner = tk.Frame(canvas, bg=BG_LIGHT)
        canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_resize(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_resize)

        # Mouse-wheel scrolling (cross-platform best effort).
        def _on_wheel(event):
            delta = -1 * (event.delta // 120) if event.delta else (
                -1 if getattr(event, "num", 0) == 4 else 1
            )
            canvas.yview_scroll(delta, "units")
        canvas.bind_all("<MouseWheel>", _on_wheel)
        canvas.bind_all("<Button-4>", _on_wheel)
        canvas.bind_all("<Button-5>", _on_wheel)

        for idx, note in enumerate(self.notes):
            var = tk.BooleanVar(value=False)
            var.trace_add("write", lambda *_: self._update_count())
            self._vars.append(var)
            self._build_note_row(inner, idx, note, var)

        # Output path picker
        out_frame = ttk.Frame(outer)
        out_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(out_frame, text="Output file:", style="Sub.TLabel").pack(
            side="left"
        )
        self.lbl_outpath = ttk.Label(
            out_frame, text=str(self.output_path), style="Sub.TLabel",
        )
        self.lbl_outpath.pack(side="left", padx=(6, 6))
        tk.Button(
            out_frame, text="Save As…", command=self._pick_output,
            bg=SURFACE, fg=FG, activebackground=ACCENT, activeforeground=BG,
            font=("Segoe UI", 9), relief="flat", padx=8, pady=2,
        ).pack(side="right")

        # Action buttons
        action_frame = ttk.Frame(outer)
        action_frame.pack(fill="x", pady=(10, 0))
        tk.Button(
            action_frame, text="Cancel", command=self.win.destroy,
            bg=SURFACE, fg=FG, activebackground=RED, activeforeground=BG,
            font=("Segoe UI", 10), relief="flat", padx=14, pady=6,
        ).pack(side="right")
        self.btn_delete = tk.Button(
            action_frame, text="Delete Selected", command=self._delete,
            bg=RED, fg=BG, activebackground=ACCENT, activeforeground=BG,
            font=("Segoe UI", 11, "bold"), relief="flat", padx=18, pady=6,
            disabledforeground=FG_DIM,
        )
        self.btn_delete.pack(side="right", padx=(0, 8))

        self._update_count()

    def _build_note_row(self, parent, idx: int, note: SpecifierNote, var: tk.BooleanVar):
        row = tk.Frame(parent, bg=BG_LIGHT, padx=8, pady=6)
        row.pack(fill="x", anchor="w")

        cb = tk.Checkbutton(
            row, variable=var,
            bg=BG_LIGHT, fg=FG, selectcolor=BG,
            activebackground=BG_LIGHT, activeforeground=FG,
            highlightthickness=0, borderwidth=0,
        )
        cb.pack(side="left", anchor="n")

        text_col = tk.Frame(row, bg=BG_LIGHT)
        text_col.pack(side="left", fill="x", expand=True, padx=(6, 0))

        location_text = (
            f"#{idx + 1} • {note.location_label} • confidence {note.confidence:.2f}"
        )
        tk.Label(
            text_col, text=location_text,
            bg=BG_LIGHT, fg=ACCENT, font=("Segoe UI", 9, "bold"),
            anchor="w", justify="left",
        ).pack(fill="x", anchor="w")

        preview = note.text.replace("\n", " ").strip()
        # No truncation — users want to see the whole note when deciding.
        tk.Label(
            text_col, text=preview,
            bg=BG_LIGHT, fg=FG, font=("Consolas", 9),
            anchor="w", justify="left", wraplength=620,
        ).pack(fill="x", anchor="w", pady=(2, 0))

        if note.reason:
            tk.Label(
                text_col, text=note.reason,
                bg=BG_LIGHT, fg=FG_DIM, font=("Segoe UI", 8, "italic"),
                anchor="w", justify="left", wraplength=620,
            ).pack(fill="x", anchor="w", pady=(2, 0))

    def _select_all(self):
        for v in self._vars:
            v.set(True)

    def _select_none(self):
        for v in self._vars:
            v.set(False)

    def _update_count(self):
        n = sum(1 for v in self._vars if v.get())
        self.lbl_count.configure(text=f"{n} of {len(self._vars)} selected")

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Save cleaned DOCX as…",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
            initialdir=str(self.output_path.parent),
            initialfile=self.output_path.name,
        )
        if path:
            self.output_path = Path(path)
            self.lbl_outpath.configure(text=str(self.output_path))

    def _delete(self):
        selected = [n for n, v in zip(self.notes, self._vars) if v.get()]
        if not selected:
            messagebox.showinfo(
                "Inspect Notes",
                "No notes selected. Tick the boxes for the notes you want "
                "to delete, or click Cancel.",
            )
            return

        if self.output_path.exists():
            ok = messagebox.askyesno(
                "Overwrite?",
                f"{self.output_path.name} already exists.\n\nOverwrite it?",
            )
            if not ok:
                return

        self.btn_delete.configure(state="disabled")

        def _work():
            try:
                report = remove_specifier_notes_at_locations(
                    self.input_path, self.output_path, selected,
                )
            except Exception as exc:
                self.logger(f"  FAILED: {exc}")
                self.parent.after(
                    0,
                    lambda: messagebox.showerror("Inspect Notes", str(exc)),
                )
                self.parent.after(
                    0, lambda: self.btn_delete.configure(state="normal"),
                )
                return

            def _finish():
                if report.success:
                    self.logger(
                        f"  Removed {report.removed_count} note(s) → "
                        f"{self.output_path.name}"
                    )
                    if report.skipped:
                        self.logger(
                            f"  Skipped {len(report.skipped)} item(s):"
                        )
                        for s in report.skipped:
                            self.logger(f"    • {s}")
                    messagebox.showinfo(
                        "Inspect Notes",
                        f"Removed {report.removed_count} note(s).\n\n"
                        f"Wrote: {self.output_path}",
                    )
                    self.win.destroy()
                else:
                    for err in report.errors:
                        self.logger(f"  ERROR: {err}")
                    messagebox.showerror(
                        "Inspect Notes",
                        "Failed to remove notes:\n\n"
                        + "\n".join(report.errors),
                    )
                    self.btn_delete.configure(state="normal")

            self.parent.after(0, _finish)

        threading.Thread(target=_work, daemon=True).start()


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
            text=(
                "Removes editorial noise from spec documents.\n"
                "  • Preview — dry-run report of everything the cleaner would remove.\n"
                "  • Inspect Notes — list every specifier note with its location and "
                "let you choose which ones to delete (one file at a time).\n"
                "  • CLEAN — remove all detected noise (notes, copyright, hidden text, "
                "SpecAgent watermarks, editorial artifacts) in one pass."
            ),
            style="Sub.TLabel",
            wraplength=680,
            justify="left",
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

        self.btn_inspect = tk.Button(
            action_frame, text="Inspect Notes", command=self._start_inspect,
            bg=YELLOW, fg=BG, activebackground=ACCENT, activeforeground=BG,
            font=("Segoe UI", 11, "bold"), relief="flat", padx=18, pady=6,
            disabledforeground=FG_DIM,
        )
        self.btn_inspect.pack(side="left", padx=(0, 8))

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
        self.btn_inspect.configure(state="disabled")
        self.btn_clean.configure(state="disabled")
        self.btn_add.configure(state="disabled")
        self.btn_clear.configure(state="disabled")

    def _enable_controls(self):
        self._running = False
        self.btn_preview.configure(state="normal")
        self.btn_inspect.configure(state="normal")
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

    def _start_inspect(self):
        if self._running:
            return
        if not self.files:
            self._log("No files selected. Click 'Add Files...' first.")
            return
        if len(self.files) != 1:
            messagebox.showinfo(
                "Inspect Notes",
                "Inspect Notes works on one file at a time.\n\n"
                "Use Clear and select a single .docx, or use CLEAN to "
                "process multiple files in batch.",
            )
            return

        target = self.files[0]
        self._disable_controls()
        self._clear_log()
        self._log(f"Scanning {target.name} for specifier notes…")
        threading.Thread(
            target=self._run_inspect, args=(target,), daemon=True
        ).start()

    def _run_inspect(self, input_path: Path):
        try:
            config = load_config()
            notes = extract_specifier_notes(input_path, config)
        except Exception as exc:
            self._log(f"  FAILED: {exc}")
            self.root.after(0, self._enable_controls)
            return

        self._log(f"  Found {len(notes)} specifier note(s).")
        self._set_status(f"Found {len(notes)} note(s) in {input_path.name}")

        # Hand control to the review window on the main thread.
        def _open():
            self._enable_controls()
            if not notes:
                messagebox.showinfo(
                    "Inspect Notes",
                    f"No specifier notes detected in {input_path.name}.",
                )
                return
            NotesReviewWindow(
                parent=self.root,
                input_path=input_path,
                notes=notes,
                default_output=self._output_for(input_path),
                logger=self._log,
            )

        self.root.after(0, _open)

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
