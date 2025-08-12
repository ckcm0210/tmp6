# -*- coding: utf-8 -*-
"""Module for managing the details panel UI and logic.

This module is responsible for populating the details text widget when a user
selects an item from the main formula list.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import re
from urllib.parse import unquote

from core import navigation_manager
from core.link_analyzer import get_referenced_cell_values
from utils.excel_io import find_matching_sheet, read_external_cell_value
from core.excel_connector import find_external_workbook_path
from ui.dependency_exploder_view import DependencyExploderView

def populate_details_panel(controller, selected_item_id):
    """Populates the details panel based on the selected formula item."""
    try:
        main_tab_info = controller.tab_manager.detail_tabs.get("Main") or controller.tab_manager.detail_tabs.get("Tab_0")
        if not main_tab_info:
            current_detail_text = controller.tab_manager.get_current_detail_text()
        else:
            current_detail_text = main_tab_info["text_widget"]
            controller.tab_manager.detail_notebook.select(main_tab_info["frame"])
    except (AttributeError, KeyError):
        current_detail_text = controller.tab_manager.get_current_detail_text()

    if not selected_item_id:
        current_detail_text.delete(1.0, 'end')
        return
        
    values = controller.view.result_tree.item(selected_item_id, "values")
    
    if len(values) < 5:
        current_detail_text.delete(1.0, 'end')
        current_detail_text.insert(1.0, "Selected item has incomplete data.")
        return
        
    formula_type, cell_address, formula, result, display_value = values
    
    current_detail_text.delete(1.0, 'end')
    current_detail_text.insert('end', "Type: ", "label")
    current_detail_text.insert('end', f"{formula_type} / ", "value")
    current_detail_text.insert('end', "Cell Address: ", "label")
    current_detail_text.insert('end', f"{cell_address}\n", "value")
    current_detail_text.insert('end', "Calculated Result: ", "label")
    current_detail_text.insert('end', f"{result} / ", "result_value")
    current_detail_text.insert('end', "Displayed Value: ", "label")
    current_detail_text.insert('end', f"{display_value}\n\n", "value")
    current_detail_text.insert('end', "Formula Content:\n", "label")
    current_detail_text.insert('end', f"{formula}  ", "formula_content")
    
    try:
        def build_explode_handler():
            def handler():
                if hasattr(controller, 'workbook') and controller.workbook:
                    current_workbook_path = controller.workbook.FullName
                    current_sheet_name = controller.worksheet.Name if hasattr(controller, 'worksheet') and controller.worksheet else "Unknown"
                    selected_item = controller.view.result_tree.selection()
                    if selected_item:
                        item_id = selected_item[0]
                        current_cell_address = controller.cell_addresses.get(item_id, "A1")
                        DependencyExploderView(controller.view, controller, current_workbook_path, current_sheet_name, current_cell_address, f'{current_sheet_name}!{current_cell_address}')
                    else:
                        messagebox.showwarning("No Selection", "Please select a cell first.")
                else:
                    messagebox.showerror("Excel Not Connected", "Excel connection not available for dependency analysis.")
            return handler
        
        explode_btn = tk.Button(current_detail_text, text="Explode", font=("Arial", 8, "bold"), cursor="hand2", bg="#ffeb3b", command=build_explode_handler())
        current_detail_text.window_create('end', window=explode_btn)
    except Exception as e:
        print(f"Could not create Explode button: {e}")
    
    current_detail_text.insert('end', "\n\n")
    
    referenced_values = None
    excel_connected = controller.xl and controller.worksheet
    
    if excel_connected:
        try:
            read_func = read_external_cell_value
            referenced_values = get_referenced_cell_values(
                formula, controller.worksheet, controller.workbook.FullName, read_func,
                lambda name, obj: find_matching_sheet(controller.workbook, name)
            )
        except Exception as e:
            print(f"Warning: Could not get referenced values: {e}")
            referenced_values = None
    
    current_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
    
    if referenced_values:
        for ref_addr, ref_val in referenced_values.items():
            display_text = ref_addr.split('|', 1)[1] if '|' in ref_addr else ref_addr
            current_detail_text.insert('end', f"  {display_text}: {ref_val}  ", "referenced_value")
            # Navigation button logic will be added here in a future step
            current_detail_text.insert('end', "\n")
    else:
        current_detail_text.insert('end', "  No individual cell references found or accessible.\n", "info_text")
    
    if not excel_connected:
        current_detail_text.insert('end', "\nNote: Excel connection not active. Values shown as 'N/A'.\n", "info_text")
