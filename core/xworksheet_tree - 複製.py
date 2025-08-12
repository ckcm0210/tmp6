# -*- coding: utf-8 -*-
"""Module for the main formula list Treeview and its interactions.

This module is responsible for filtering, sorting, and handling user
interactions on the main formula list.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import re

# Refactored imports
from core import navigation_manager
from core import details_manager
from utils.range_optimizer import parse_excel_address
from openpyxl.utils import get_column_letter, column_index_from_string


_last_range_threshold = 5
_last_max_depth = 10

def apply_filter(controller, event=None):
    """Filters the formula list based on user input."""
    controller.view.result_tree.delete(*controller.view.result_tree.get_children())
    controller.cell_addresses.clear()
    address_filter_str = controller.view.filter_entries['address'].get().strip()
    parsed_address_filters = []
    if address_filter_str and address_filter_str != controller.placeholder_text:
        address_tokens = [token.strip() for token in address_filter_str.split(',') if token.strip()]
        if address_tokens:
            try:
                for token in address_tokens:
                    parsed_address_filters.append(parse_excel_address(token))
            except Exception as e:
                messagebox.showerror("Invalid Excel Address", str(e))
                return
    other_filters = {
        'type': (controller.show_formula.get(), controller.show_local_link.get(), controller.show_external_link.get()),
        'formula': controller.view.filter_entries['formula'].get().lower(),
        'result': controller.view.filter_entries['result'].get().lower(),
        'display_value': controller.view.filter_entries['display_value'].get().lower()
    }
    filtered_formulas = []
    for formula_data in controller.all_formulas:
        if len(formula_data) < 5: continue
        formula_type, address, formula_content, result_val, display_val = formula_data
        type_map = {'formula': other_filters['type'][0], 'local link': other_filters['type'][1], 'external link': other_filters['type'][2]}
        if not type_map.get(formula_type, True): continue
        if other_filters['formula'] and other_filters['formula'] not in str(formula_content).lower(): continue
        if other_filters['result'] and other_filters['result'] not in str(result_val).lower(): continue
        if other_filters['display_value'] and other_filters['display_value'] not in str(display_val).lower(): continue
        if parsed_address_filters:
            addr_upper = address.replace("$", "").upper()
            current_cell_match = re.match(r"([A-Z]+)([0-9]+)", addr_upper)
            if not current_cell_match: continue
            cell_col_str, cell_row_str = current_cell_match.groups()
            cell_col_idx = column_index_from_string(cell_col_str)
            cell_row_idx = int(cell_row_str)
            is_match = False
            for f_type, f_val in parsed_address_filters:
                if f_type == 'cell' and addr_upper == f_val:
                    is_match = True; break
                elif f_type == 'row_range':
                    start_r, end_r = map(int, f_val.split(':'))
                    if start_r <= cell_row_idx <= end_r:
                        is_match = True; break
                elif f_type == 'col_range':
                    start_c, end_c = f_val.split(':')
                    if column_index_from_string(start_c) <= cell_col_idx <= column_index_from_string(end_c):
                        is_match = True; break
                elif f_type == 'range':
                    start_cell, end_cell = f_val.split(':')
                    sc_str, sr_str = re.match(r"([A-Z]+)([0-9]+)", start_cell).groups()
                    ec_str, er_str = re.match(r"([A-Z]+)([0-9]+)", end_cell).groups()
                    if (column_index_from_string(sc_str) <= cell_col_idx <= column_index_from_string(ec_str) and
                        int(sr_str) <= cell_row_idx <= int(er_str)):
                        is_match = True; break
            if not is_match: continue
        filtered_formulas.append(formula_data)
    if controller.current_sort_column:
        col_index = controller.view.tree_columns.index(controller.current_sort_column)
        sort_dir = controller.sort_directions[controller.current_sort_column]
        filtered_formulas.sort(key=lambda x: str(x[col_index]), reverse=(sort_dir == -1))
    count = len(filtered_formulas)
    controller.view.formula_list_label.config(text=f"Formula List ({count} records):")
    for i, data in enumerate(filtered_formulas):
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        item_id = controller.view.result_tree.insert("", "end", values=data, tags=(tag,))
        address_index = controller.view.tree_columns.index("address")
        if address_index < len(data):
            controller.cell_addresses[item_id] = data[address_index]

def sort_column(controller, col_id):
    """Handles sorting of the formula list view."""
    controller.current_sort_column = col_id
    controller.sort_directions[col_id] *= -1
    apply_filter(controller)
    for column in controller.view.tree_columns:
        original_text = controller.view.result_tree.heading(column, "text").split(' ')[0]
        controller.view.result_tree.heading(column, text=original_text, image='')
    current_direction = " \u2191" if controller.sort_directions[col_id] == 1 else " \u2193"
    current_text = controller.view.result_tree.heading(col_id, "text").split(' ')[0]
    controller.view.result_tree.heading(col_id, text=current_text + current_direction)

def on_select(controller, event):
    """Handles the selection event on the formula list by delegating to the details manager."""
    selected_item = controller.view.result_tree.selection()
    details_manager.populate_details_panel(controller, selected_item[0] if selected_item else None)
        
def on_double_click(controller, event):
    """Handles double-click event on the formula list to navigate to the cell."""
    selected_item = controller.view.result_tree.selection()
    if not selected_item:
        return
    item_id = selected_item[0]
    cell_address = controller.cell_addresses.get(item_id)
    if cell_address:
        navigation_manager.navigate_in_active_sheet(controller, cell_address)
