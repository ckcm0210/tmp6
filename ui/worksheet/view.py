
# -*- coding: utf-8 -*-
"""
Worksheet View Module

This module contains the WorksheetView class, which is responsible for
creating and managing all UI widgets for a single worksheet pane.
"""

import tkinter as tk
from tkinter import ttk
from tkinter import font

from ui.worksheet_ui import create_ui_widgets, bind_ui_commands, _set_placeholder, _on_focus_in, _on_mouse_click, _on_focus_out

class WorksheetView(ttk.Frame):
    """Manages the UI for a single worksheet pane."""
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # --- UI Initialization ---
        create_ui_widgets(self)

    def bind_ui_commands(self):
        bind_ui_commands(self)

    # Placeholder methods for font management, still needed by bind_ui_commands
    def _set_placeholder(self):
        return _set_placeholder(self)

    def _on_focus_in(self, event):
        return _on_focus_in(self, event)

    def _on_mouse_click(self, event):
        return _on_mouse_click(self, event)

    def _on_focus_out(self, event):
        return _on_focus_out(self, event)
