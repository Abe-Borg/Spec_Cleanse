"""The window's colours and fonts, shared by every dialog.

Kept apart from ``gui.py`` so ``trust_view.py`` can draw in the same palette
without importing the main window, which imports it.
"""

BG        = "#1e1e2e"
BG_LIGHT  = "#313244"
FG        = "#cdd6f4"
FG_DIM    = "#6c7086"
ACCENT    = "#89b4fa"
GREEN     = "#a6e3a1"
RED       = "#f38ba8"
YELLOW    = "#f9e2af"
SURFACE   = "#45475a"

UI_FONT = "Segoe UI"
MONO_FONT = "Consolas"
