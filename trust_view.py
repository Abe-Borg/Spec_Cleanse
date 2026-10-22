#!/usr/bin/env python3
"""
The two trust dialogs: "Why trust it?" and the dossier it opens.

Purpose and reader: see ``trust.py``, which holds every word shown here. This
module only draws them. If the implementation changes, the words in ``trust.py``
change with it — a trust document that has drifted from the code is worse than
none — and nothing here may add a claim of its own.

Two rules this file keeps, because the dossier claims them:

* It loads nothing: no image, font, script or page from anywhere. The diagram is
  drawn on a canvas by the code below, and the further-reading addresses are
  shown as text to copy — the program opens no browser.
* Dialogs stack. The dossier opens on top of "Why trust it?", each one holds
  the input while it is on top, one Escape closes exactly one of them, and
  closing one returns focus to the button that opened it.

Tk has no dialog roles or accessible labels to set, so the diagram is followed
by a sentence restating it, and every section is reachable by keyboard scroll
as well as from the contents rail.
"""

from __future__ import annotations

import re
import tkinter as tk
from tkinter import ttk

import trust
from theme import ACCENT, BG, BG_LIGHT, FG, FG_DIM, GREEN, MONO_FONT, RED, SURFACE, UI_FONT, YELLOW

#: The dossier is at most this share of the screen's height.
DOSSIER_MAX_HEIGHT = 0.88
#: Below this window width the contents rail is hidden to give the text room.
RAIL_MIN_WIDTH = 900

_BOLD = re.compile(r"\*\*(.+?)\*\*")


# ---------------------------------------------------------------------------
# Stacking
# ---------------------------------------------------------------------------

class _StackedDialog(tk.Toplevel):
    """A dialog that takes the input while it is on top and gives it back."""

    def __init__(self, master: tk.Misc, opener: tk.Widget | None, title: str):
        super().__init__(master)
        self.title(title)
        self.configure(bg=BG)
        self.transient(master.winfo_toplevel())
        self.opener = opener
        # Whoever held the input before us — the dialog beneath, if any.
        self._previous_grab = self.grab_current()
        self.protocol("WM_DELETE_WINDOW", self.close)
        # Bound on the toplevel, which is in the binding path of every widget
        # inside it and of no widget in any other window: Escape pressed here
        # reaches this dialog's handler and nothing beneath it.
        self.bind("<Escape>", self._on_escape)
        self.closed = False

    def _on_escape(self, _event=None):
        self.close()
        return "break"

    def take_input(self, focus: tk.Widget) -> None:
        """Hold the input and put focus inside this dialog."""
        self.update_idletasks()
        try:
            self.grab_set()
        except tk.TclError:
            # Not yet viewable on some window managers; focus still moves.
            pass
        focus.focus_set()

    def close(self) -> None:
        if self.closed:
            return
        self.closed = True
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()
        previous = self._previous_grab
        if previous is not None and previous.winfo_exists():
            try:
                previous.grab_set()
            except tk.TclError:
                pass
        if self.opener is not None and self.opener.winfo_exists():
            self.opener.focus_set()


def _button(parent, text, command, primary=False, **extra) -> tk.Button:
    """A button in the main window's style."""
    if primary:
        colours = dict(bg=ACCENT, fg=BG, activebackground=GREEN, activeforeground=BG,
                       font=(UI_FONT, 10, "bold"))
    else:
        colours = dict(bg=SURFACE, fg=FG, activebackground=ACCENT, activeforeground=BG,
                       font=(UI_FONT, 10))
    return tk.Button(parent, text=text, command=command, relief="flat",
                     padx=12, pady=4, highlightthickness=1,
                     highlightbackground=BG, highlightcolor=ACCENT,
                     **colours, **extra)


def _reading_text(parent) -> tk.Text:
    """A read-only text area: selectable and copyable, not editable."""
    text = tk.Text(
        parent, wrap="word", bg=BG, fg=FG, insertbackground=FG,
        relief="flat", borderwidth=0, highlightthickness=1,
        highlightcolor=ACCENT, highlightbackground=BG,
        padx=18, pady=12, font=(UI_FONT, 10), cursor="arrow",
    )
    text.tag_configure("bold", font=(UI_FONT, 10, "bold"))
    return text


def _insert_marked(text: tk.Text, content: str, tags: tuple = ()) -> None:
    """Insert ``content``, rendering ``**bold**`` spans in bold."""
    position = 0
    for match in _BOLD.finditer(content):
        text.insert("end", content[position:match.start()], tags)
        text.insert("end", match.group(1), tags + ("bold",))
        position = match.end()
    text.insert("end", content[position:], tags)


# ---------------------------------------------------------------------------
# Level 1 — "Why trust it?"
# ---------------------------------------------------------------------------

class WhyTrustDialog(_StackedDialog):
    """The short topic: a lead, the points, and the way into the dossier."""

    def __init__(self, master: tk.Misc, opener: tk.Widget | None = None):
        super().__init__(master, opener, trust.LEVEL1_TITLE)
        self.dossier: DossierDialog | None = None

        screen_h = self.winfo_screenheight()
        height = min(720, int(screen_h * DOSSIER_MAX_HEIGHT))
        self.geometry(f"680x{height}")
        self.minsize(420, 360)

        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=18, pady=(14, 0))
        tk.Label(header, text=trust.LEVEL1_TITLE, bg=BG, fg=ACCENT,
                 font=(UI_FONT, 18, "bold")).pack(side="left")
        self.close_button = _button(header, "Close", self.close)
        self.close_button.pack(side="right")

        # The closing box is packed before the text so it stays on screen
        # however small the window gets; the text takes what is left.
        closing = tk.Frame(self, bg=BG_LIGHT, padx=16, pady=12)
        closing.pack(side="bottom", fill="x", padx=18, pady=(4, 16))
        tk.Label(closing, text=trust.LEVEL1_CLOSING_TITLE, bg=BG_LIGHT, fg=YELLOW,
                 font=(UI_FONT, 12, "bold")).pack(anchor="w")
        closing_text = tk.Label(closing, text=trust.LEVEL1_CLOSING_TEXT, bg=BG_LIGHT,
                                fg=FG, font=(UI_FONT, 10), justify="left",
                                anchor="w", wraplength=600)
        closing_text.pack(anchor="w", pady=(2, 8), fill="x")
        closing.bind("<Configure>", lambda e: closing_text.configure(
            wraplength=max(200, e.width - 40)))
        self.button = _button(closing, trust.LEVEL1_BUTTON, self.open_dossier, primary=True)
        self.button.pack(anchor="w")

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=(0, 18))
        self.text = _reading_text(body)
        scroll = ttk.Scrollbar(body, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)
        self._fill()

        self.take_input(self.text)

    def _fill(self) -> None:
        text = self.text
        text.tag_configure("lead", font=(UI_FONT, 12), foreground=FG,
                           spacing3=10)
        text.tag_configure("point_title", font=(UI_FONT, 11, "bold"),
                           foreground=ACCENT, spacing1=8)
        text.tag_configure("point_body", lmargin1=0, lmargin2=0, spacing3=4)
        _insert_marked(text, trust.LEVEL1_LEAD + "\n", ("lead",))
        for point in trust.LEVEL1_POINTS:
            _insert_marked(text, point.title + "\n", ("point_title",))
            _insert_marked(text, point.body + "\n", ("point_body",))
        text.configure(state="disabled")

    def open_dossier(self) -> "DossierDialog":
        """Open the dossier on top of this dialog."""
        if self.dossier is not None and not self.dossier.closed:
            self.dossier.lift()
            self.dossier.take_input(self.dossier.text)
            return self.dossier
        self.dossier = DossierDialog(self, opener=self.button)
        return self.dossier


# ---------------------------------------------------------------------------
# Level 2 — the dossier
# ---------------------------------------------------------------------------

class DossierDialog(_StackedDialog):
    """Every section in one scroll, with a contents rail that follows it."""

    def __init__(self, master: tk.Misc, opener: tk.Widget | None = None):
        super().__init__(master, opener, f"{trust.LEVEL1_TITLE} — {trust.DOSSIER_TITLE}")

        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
        self.max_height = int(screen_h * DOSSIER_MAX_HEIGHT)
        width = min(1180, int(screen_w * 0.92))
        self.geometry(f"{width}x{self.max_height}")
        self.minsize(520, 360)
        self.maxsize(screen_w, self.max_height)

        top = tk.Frame(self, bg=BG)
        top.pack(fill="x", padx=18, pady=(14, 6))
        self.close_button = _button(top, "Close", self.close)
        self.close_button.pack(side="right", anchor="n")
        tk.Label(top, text=trust.DOSSIER_TITLE, bg=BG, fg=ACCENT,
                 font=(UI_FONT, 18, "bold")).pack(anchor="w")
        subtitle = tk.Label(top, text=trust.DOSSIER_SUBTITLE, bg=BG, fg=FG_DIM,
                            font=(UI_FONT, 9), justify="left", wraplength=900)
        subtitle.pack(anchor="w", fill="x")
        top.bind("<Configure>", lambda e: subtitle.configure(
            wraplength=max(240, e.width - 120)))

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=(18, 18), pady=(0, 16))

        # A fixed width, so the rail does not shift as the highlight moves.
        self.rail = tk.Frame(body, bg=BG, width=230)
        self.rail.pack_propagate(False)
        self.rail_buttons: dict[str, tk.Button] = {}
        tk.Label(self.rail, text="Contents", bg=BG, fg=FG_DIM,
                 font=(UI_FONT, 9, "bold")).pack(anchor="w", pady=(0, 4))
        for section in trust.DOSSIER:
            button = tk.Button(
                self.rail, text=section.title, anchor="w", justify="left",
                wraplength=200, relief="flat", bg=BG, fg=FG,
                highlightthickness=1, highlightbackground=BG,
                highlightcolor=ACCENT,
                activebackground=SURFACE, activeforeground=FG,
                font=(UI_FONT, 9), padx=8, pady=3,
                command=lambda s=section.id: self.go_to(s),
            )
            button.pack(fill="x")
            self.rail_buttons[section.id] = button

        self.text_frame = tk.Frame(body, bg=BG)
        self.text_frame.pack(side="right", fill="both", expand=True)
        self.text = _reading_text(self.text_frame)
        self.scroll = ttk.Scrollbar(self.text_frame, orient="vertical",
                                    command=self.text.yview)
        self.text.configure(yscrollcommand=self._on_scroll)
        self.scroll.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)

        self.rail_visible = False
        self.current_section: str | None = None
        self._resizers: list = []
        self._fill()

        self.bind("<Configure>", self._on_resize, add="+")
        self.text.bind("<Configure>", self._on_text_resize, add="+")
        self._show_rail(width >= RAIL_MIN_WIDTH)
        self._highlight(trust.DOSSIER[0].id)

        self.take_input(self.text)

    # -- navigation ---------------------------------------------------------

    @staticmethod
    def mark_for(section_id: str) -> str:
        return f"section-{section_id}"

    def go_to(self, section_id: str) -> None:
        """Scroll so the section starts at the top.  The rail only navigates."""
        self.text.yview(self.mark_for(section_id))
        self._highlight(section_id)

    def _on_scroll(self, first, last) -> None:
        self.scroll.set(first, last)
        self._highlight(self._section_at_top())

    def _section_at_top(self) -> str:
        top = self.text.index("@0,0")
        current = trust.DOSSIER[0].id
        for section in trust.DOSSIER:
            if self.text.compare(self.mark_for(section.id), "<=", top):
                current = section.id
        return current

    def _highlight(self, section_id: str) -> None:
        if section_id == self.current_section:
            return
        self.current_section = section_id
        for sid, button in self.rail_buttons.items():
            on = sid == section_id
            button.configure(bg=SURFACE if on else BG, fg=ACCENT if on else FG,
                             font=(UI_FONT, 9, "bold" if on else "normal"))

    def _show_rail(self, show: bool) -> None:
        if show == self.rail_visible:
            return
        self.rail_visible = show
        if show:
            self.rail.pack(side="left", fill="y", padx=(0, 12), before=self.text_frame)
        else:
            self.rail.pack_forget()

    def _on_resize(self, event) -> None:
        if event.widget is self:
            self._show_rail(event.width >= RAIL_MIN_WIDTH)

    def _on_text_resize(self, event) -> None:
        usable = max(320, event.width - 60)
        for resize in self._resizers:
            resize(usable)

    # -- content ------------------------------------------------------------

    def _fill(self) -> None:
        text = self.text
        hang = 150
        text.tag_configure("kicker", font=(UI_FONT, 9, "bold"), foreground=YELLOW,
                           spacing1=26)
        text.tag_configure("title", font=(UI_FONT, 16, "bold"), foreground=ACCENT,
                           spacing3=8)
        text.tag_configure("body", spacing3=6)
        text.tag_configure("lead", font=(UI_FONT, 10, "bold"), foreground=FG,
                           spacing1=6)
        text.tag_configure("bullet", lmargin1=6, lmargin2=24, tabs=(24,), spacing3=4)
        text.tag_configure("record_head", font=(UI_FONT, 10, "bold"),
                           foreground=ACCENT, spacing1=10)
        text.tag_configure("record", lmargin1=16, lmargin2=16 + hang,
                           tabs=(16 + hang,), spacing3=3)
        text.tag_configure("label", foreground=FG_DIM, font=(UI_FONT, 9, "bold"))
        text.tag_configure("box", background=BG_LIGHT, lmargin1=14, lmargin2=14,
                           rmargin=14, lmargincolor=BG_LIGHT, rmargincolor=BG_LIGHT)
        text.tag_configure("box_title", font=(UI_FONT, 10, "bold"),
                           foreground=YELLOW, spacing1=8)
        text.tag_configure("claim_title", font=(UI_FONT, 10, "bold"),
                           foreground=GREEN, spacing1=8)
        text.tag_configure("box_end", font=(UI_FONT, 4))
        text.tag_configure("card_title", font=(UI_FONT, 11, "bold"),
                           foreground=ACCENT, spacing1=14, background=BG_LIGHT,
                           lmargin1=14, lmargincolor=BG_LIGHT)
        text.tag_configure("card_row", background=BG_LIGHT, lmargin1=14,
                           lmargin2=14 + hang, tabs=(14 + hang,), rmargin=14,
                           lmargincolor=BG_LIGHT, rmargincolor=BG_LIGHT, spacing3=3)
        text.tag_configure("none", foreground=GREEN, font=(UI_FONT, 10, "bold"))
        text.tag_configure("url", font=(MONO_FONT, 9), foreground=FG)
        text.tag_configure("caption", foreground=FG_DIM, font=(UI_FONT, 9),
                           spacing3=10)

        card_number = 0
        for section in trust.DOSSIER:
            text.mark_set(self.mark_for(section.id), "end-1c")
            text.mark_gravity(self.mark_for(section.id), "left")
            text.insert("end", section.kicker.upper() + "\n", ("kicker",))
            text.insert("end", section.title + "\n", ("title",))
            for block in section.blocks:
                if isinstance(block, trust.RuntimeCard):
                    card_number += 1
                    self._card(block, card_number)
                else:
                    self._block(block)
        text.configure(state="disabled")

    def _block(self, block) -> None:
        text = self.text
        if isinstance(block, trust.Point):
            _insert_marked(text, block.title + "\n", ("lead",))
            _insert_marked(text, block.body + "\n", ("body",))
        elif isinstance(block, trust.Para):
            _insert_marked(text, block.text + "\n", ("body",))
        elif isinstance(block, trust.Bullets):
            for item in block.items:
                _insert_marked(text, "•\t" + item + "\n", ("bullet",))
        elif isinstance(block, trust.Table):
            for row in block.rows:
                text.insert("end", row[0] + "\n", ("record_head",))
                for column, cell in zip(block.columns[1:], row[1:]):
                    text.insert("end", column + "\t", ("record", "label"))
                    _insert_marked(text, cell + "\n", ("record",))
            text.insert("end", "\n", ("body",))
        elif isinstance(block, trust.Note):
            self._box(block.title, block.text, "box_title")
        elif isinstance(block, trust.Term):
            text.insert("end", block.word + "\n", ("record_head",))
            self._box("The exact claim", block.claim, "claim_title", gap=False)
            text.insert("end", "What it does not claim\t", ("record", "label"))
            _insert_marked(text, block.not_claimed + "\n", ("record",))
        elif isinstance(block, trust.Contrast):
            self._contrast(block)
        elif isinstance(block, trust.Diagram):
            self._diagram(block)
        elif isinstance(block, trust.LinkBox):
            text.insert("end", block.title + "\n", ("box", "box_title"))
            for link in block.links:
                text.insert("end", link.label + "\n", ("box", "bold"))
                text.insert("end", link.url + "\n", ("box", "url"))
            text.insert("end", "\n", ("box", "box_end"))
        else:  # pragma: no cover - trust.py and this file grow together
            raise TypeError(f"no drawing for {block!r}")

    def _box(self, title: str, body: str, title_tag: str, gap: bool = True) -> None:
        text = self.text
        text.insert("end", title + "\n", ("box", title_tag))
        _insert_marked(text, body + "\n", ("box",))
        text.insert("end", "\n", ("box", "box_end"))
        if gap:
            text.insert("end", "\n", ("box_end",))

    def _card(self, card: trust.RuntimeCard, number: int) -> None:
        text = self.text
        text.insert("end", f"{number}. {card.title}\n", ("card_title",))
        for field, label in trust.CARD_ROWS:
            value = getattr(card, field)
            text.insert("end", label + "\t", ("card_row", "label"))
            none = field == "ai_involved" and value == "None."
            tags = ("card_row", "none") if none else ("card_row",)
            _insert_marked(text, value + "\n", tags)
        text.insert("end", "\n", ("card_row", "box_end"))
        text.insert("end", "\n", ("box_end",))

    def _contrast(self, block: trust.Contrast) -> None:
        frame = tk.Frame(self.text, bg=BG)
        labels: list[tk.Label] = []
        for column, (title, items, colour) in enumerate((
            (block.left_title, block.left, RED),
            (block.right_title, block.right, GREEN),
        )):
            panel = tk.Frame(frame, bg=BG_LIGHT, padx=12, pady=10)
            panel.grid(row=0, column=column, sticky="nsew", padx=(0, 10) if column == 0 else 0)
            frame.columnconfigure(column, weight=1, uniform="contrast")
            tk.Label(panel, text=title, bg=BG_LIGHT, fg=colour,
                     font=(UI_FONT, 10, "bold")).pack(anchor="w", pady=(0, 4))
            for item in items:
                label = tk.Label(panel, text="•  " + item, bg=BG_LIGHT, fg=FG,
                                 font=(UI_FONT, 9), justify="left", anchor="w")
                label.pack(anchor="w", fill="x", pady=1)
                labels.append(label)

        def resize(width: int) -> None:
            frame.configure(width=width)
            for label in labels:
                label.configure(wraplength=max(140, width // 2 - 40))
        resize(760)
        self._resizers.append(resize)
        self.text.window_create("end", window=frame, padx=0, pady=6)
        self.text.insert("end", "\n", ("body",))

    def _diagram(self, block: trust.Diagram) -> None:
        canvas = tk.Canvas(self.text, height=300, bg=BG, highlightthickness=0)
        self.diagram = canvas

        def resize(width: int) -> None:
            canvas.configure(width=width)
            draw_boundary_diagram(canvas, width)
        resize(760)
        self._resizers.append(resize)
        self.text.window_create("end", window=canvas, pady=6)
        self.text.insert("end", "\n", ("body",))
        # The words a screen reader or a reader in a hurry would need.
        self.text.insert("end", block.description + "\n", ("caption",))


def draw_boundary_diagram(canvas: tk.Canvas, width: int) -> None:
    """Draw where data goes: everything inside one box, nothing out of it.

    Solid lines are what SpecCleanse does; dashed lines are traffic Windows may
    make that a firewall could show, neither drawn from SpecCleanse.
    """
    canvas.delete("all")
    width = max(width, 520)
    height = 300
    font = (UI_FONT, 9)
    bold = (UI_FONT, 9, "bold")
    inside_right = int(width * 0.66)

    def box(x0, y0, x1, y1, label, outline=ACCENT, fill=BG_LIGHT, dash=None, colour=FG):
        canvas.create_rectangle(x0, y0, x1, y1, outline=outline, fill=fill,
                                width=2, dash=dash)
        canvas.create_text((x0 + x1) / 2, (y0 + y1) / 2, text=label, fill=colour,
                           font=font, width=max(40, x1 - x0 - 10), justify="center")

    def arrow(x0, y0, x1, y1, dash=None, colour=FG, both=False):
        canvas.create_line(x0, y0, x1, y1, fill=colour, width=2, dash=dash,
                           arrow="both" if both else "last")

    # Your computer
    canvas.create_rectangle(6, 6, inside_right, height - 6, outline=GREEN, width=2)
    canvas.create_text(18, 18, text="Your computer", anchor="w", fill=GREEN, font=bold)

    span = inside_right - 30
    w = span / 3 - 16
    x_orig, x_app, x_out = 18, 18 + span / 3, 18 + 2 * span / 3
    box(x_orig, 50, x_orig + w, 100, "Your original .docx (read only)")
    box(x_app, 50, x_app + w, 130, "SpecCleanse: cleaner, then the check", fill=SURFACE)
    box(x_out, 50, x_out + w, 100, trust.OUTPUT_NAME)
    box(x_orig, 170, x_orig + w, 220, trust.CONFIG_NAME)
    box(x_app, 190, x_app + w, 250, "Temporary folder (deleted after each file)")
    arrow(x_orig + w, 75, x_app, 75)
    arrow(x_app + w, 75, x_out, 75)
    arrow(x_orig + w, 195, x_app, 115)
    arrow(x_app + w / 2, 130, x_app + w / 2, 190, both=True)

    # Outside
    x_far = inside_right + 26
    far_w = width - x_far - 8
    box(x_far, 30, x_far + far_w, 80, "The internet", outline=RED, fill=BG)
    canvas.create_text(x_far + far_w / 2, 96, text="no connection from SpecCleanse",
                       fill=RED, font=font, width=far_w)
    box(x_far, 120, x_far + far_w, 186,
        "Windows SmartScreen (Windows itself, first launch)",
        outline=FG_DIM, fill=BG, dash=(4, 3))
    box(x_far, 206, x_far + far_w, 272,
        "Network drive or synced folder (only if your files live there)",
        outline=FG_DIM, fill=BG, dash=(4, 3))
    arrow(inside_right, 153, x_far, 153, dash=(4, 3), colour=FG_DIM)
    arrow(inside_right, 239, x_far, 239, dash=(4, 3), colour=FG_DIM, both=True)


def open_why_trust(master: tk.Misc, opener: tk.Widget | None = None) -> WhyTrustDialog:
    """Open "Why trust it?" on top of ``master``."""
    return WhyTrustDialog(master, opener=opener)
