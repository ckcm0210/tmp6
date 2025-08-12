# -*- coding: utf-8 -*-
"""
Worksheet Controller Module

This module contains the WorksheetController class, which manages the state
and business logic for a single worksheet pane.
"""

import tkinter as tk
from ui.worksheet.tab_manager import TabManager
from ui.worksheet.view import WorksheetView
from ui.summary_window import SummaryWindow

class WorksheetController:
    """Manages the state and logic for a single worksheet pane."""
    def __init__(self, parent_frame, root_app, pane_name):
        self.root = root_app
        self.pane_name = pane_name

        # Data and state attributes
        self.xl = None
        self.workbook = None
        self.worksheet = None
        self.all_formulas = []
        self.cell_addresses = {}
        self.use_openpyxl = tk.BooleanVar(value=True)
        self.show_formula = tk.BooleanVar(value=True)
        self.show_local_link = tk.BooleanVar(value=True)
        self.show_external_link = tk.BooleanVar(value=True)
        self.sort_directions = {col: 1 for col in ("type", "address", "formula", "result", "display_value")}
        self.current_sort_column = None
        self.last_workbook_path = None
        self.last_worksheet_name = None

        # Scan-related attributes
        self.scanning_selected_range = False
        self.selected_scan_address = None
        self.selected_scan_count = None
        self.original_user_selection = None
        self.original_user_count = None

        # Placeholder attributes for UI
        self.placeholder_text = "e.g. A, A:A, A:C, Z:A, 10, 10:10, 10:20, 88:17, A1:C3, D40:B5"
        self.placeholder_color = 'grey'
        self.default_fg_color = 'black'
        self.default_font = None
        self.placeholder_font = None

        # Create the view and tab manager
        self.view = WorksheetView(parent_frame, self)
        self.tab_manager = TabManager(self.view.detail_notebook)
        self.view.pack(fill='both', expand=True)
        # Now that all components are initialized, bind the commands
        self.view.bind_ui_commands()

    def clear_filter_inputs(self):
        """Clear all filter input fields when starting a new scan."""
        if hasattr(self.view, 'filter_entries') and self.view.filter_entries:
            for col_id, entry in self.view.filter_entries.items():
                if entry and hasattr(entry, 'delete'):
                    entry.delete(0, 'end')
            if 'address' in self.view.filter_entries:
                self.view._set_placeholder()

    def _filter_results_to_original_selection(self):
        """Filter scan results to show only the original user-selected cell"""
        if not hasattr(self, 'original_user_selection') or not self.original_user_selection:
            return
        
        original_address = self.original_user_selection.replace('$', '')
        # Filter all_formulas to only include the original user selection
        filtered_formulas = []
        for formula_data in self.all_formulas:
            if len(formula_data) >= 2:
                formula_type, cell_address, formula, display_val, cell_text = formula_data
                # Remove $ signs for comparison
                clean_cell_address = cell_address.replace('$', '')
                if clean_cell_address == original_address:
                    filtered_formulas.append(formula_data)
        
        # Update the formulas list
        self.all_formulas = filtered_formulas
