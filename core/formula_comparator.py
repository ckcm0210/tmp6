# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 09:10:23 2025

@author: kccheng
"""

import tkinter as tk
from tkinter import ttk, messagebox
import win32com.client
import re

from ui.worksheet.controller import WorksheetController
from ui.worksheet.view import WorksheetView
from core.excel_scanner import refresh_data
import time
from core.worksheet_tree import apply_filter

class ExcelFormulaComparator:
    def __init__(self, parent_frame, main_window):
        self.root = parent_frame
        self.main_window = main_window
        self.left_controller = None
        self.right_controller = None
        self.right_frame_placeholder = None
        self.paned_window = None
        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.configure("Large.TButton", font=("Arial", 10))

        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(0, 10))

        first_row = ttk.Frame(button_frame)
        first_row.grid(row=0, column=0, columnspan=4, sticky="ew", pady=2)
        
        # Worksheet section
        ttk.Label(first_row, text="Worksheet:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        self.worksheet_var = tk.StringVar(value="sheet1")
        self.sheet1_radio = ttk.Radiobutton(first_row, text="Sheet1", variable=self.worksheet_var, value="sheet1")
        self.sheet1_radio.pack(side=tk.LEFT, padx=2)
        self.sheet2_radio = ttk.Radiobutton(first_row, text="Sheet2", variable=self.worksheet_var, value="sheet2")
        self.sheet2_radio.pack(side=tk.LEFT, padx=2)
        
        # Separator
        ttk.Label(first_row, text=" | ", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        
        # Mode Button section (new)
        self.current_mode = "quick"
        self.mode_button = ttk.Button(
            first_row, 
            text="Mode: Quick", 
            command=self.toggle_scan_mode
        )
        self.mode_button.pack(side=tk.LEFT, padx=5)
        
        # Separator
        ttk.Label(first_row, text=" | ", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        
        # Target section
        ttk.Label(first_row, text="Target:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(5, 5))
        self.scan_full_button = ttk.Button(first_row, text="Full Worksheet", command=self.scan_worksheet_full, style="Large.TButton")
        self.scan_full_button.pack(side=tk.LEFT, padx=2)
        self.scan_selected_button = ttk.Button(first_row, text="Selected Range", command=self.scan_worksheet_selected, style="Large.TButton")
        self.scan_selected_button.pack(side=tk.LEFT, padx=2)
        
        # Remove second row with Selection info
        
        # Remove third row with central sync buttons

        self.paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill='both', expand=True)

        # Create left pane using MVC
        left_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(left_frame, weight=1)
        self.left_controller = WorksheetController(left_frame, self.root, "Worksheet1")
        
        # Configure left sync button (Sync 1 -> 2)
        self.left_controller.view.sync_button.config(
            text="Sync 1 -> 2",
            command=self.sync_1_to_2,
            state="normal"
        )

        self.right_frame_placeholder = ttk.Frame(self.paned_window)
        self.paned_window.add(self.right_frame_placeholder, weight=1)
        self.paned_window.forget(self.right_frame_placeholder)

    def toggle_scan_mode(self):
        """Toggle between Quick and Full scan modes"""
        if self.current_mode == "quick":
            self.current_mode = "full"
            self.mode_button.config(text="Mode: Full")
        else:
            self.current_mode = "quick"
            self.mode_button.config(text="Mode: Quick")
        print(f"Scan mode changed to: {self.current_mode}")

    def scan_worksheet_full(self):
        mode = self.current_mode  # Use current_mode instead of mode_var
        controller = self._get_active_controller()
        self.update_selection_info(controller)
        refresh_data(controller, self.scan_full_button, scan_mode=mode)

    def scan_worksheet_selected(self):
        mode = self.current_mode  # Use current_mode instead of mode_var
        controller = self._get_active_controller()
        
        try:
            if not hasattr(controller, 'xl') or not controller.xl:
                try:
                    controller.xl = win32com.client.GetActiveObject("Excel.Application")
                except:
                    print("Excel not found")
                    return
            
            if hasattr(controller, 'xl') and controller.xl:
                selection = controller.xl.Selection
                selected_address = selection.Address.replace('$', '')
                cell_count = selection.Count
                
                original_selected_address = selected_address
                original_cell_count = cell_count
                
                if cell_count == 1:
                    try:
                        match = re.match(r'([A-Z]+)(\d+)', selected_address)
                        if match:
                            col_letters = match.group(1)
                            row_num = int(match.group(2))
                            expanded_address = f"{col_letters}{row_num}:{col_letters}{row_num + 1}"
                            selected_address = expanded_address
                            cell_count = 2
                    except Exception as e:
                        pass
                
                controller.selected_scan_address = selected_address
                controller.selected_scan_count = cell_count
                controller.original_user_selection = original_selected_address
                controller.original_user_count = original_cell_count
                controller.scanning_selected_range = True
                
                cell_word = "cell" if original_cell_count == 1 else "cells"
                print(f"Selected: {original_selected_address} ({original_cell_count} {cell_word})")
                
                import time
                time.sleep(0.1)
                
                refresh_data(controller, self.scan_selected_button, scan_mode=mode)
            else:
                self.selection_label.config(text="Not connected")
        except Exception as e:
            print("Error getting selection")
    
    def update_selection_info(self, controller):
        """Update selection info (now just prints to console since UI element removed)"""
        try:
            if hasattr(controller, 'xl') and controller.xl and hasattr(controller, 'worksheet') and controller.worksheet:
                selection = controller.xl.Selection
                address = selection.Address
                cell_count = selection.Count
                clean_address = address.replace('$', '')
                print(f"Selection: {clean_address} ({cell_count} cells)")
            else:
                print("Excel not connected")
        except Exception as e:
            print("Selection info: A1 (1 cell)")

    def sync_1_to_2(self):
        source_controller = self._get_active_controller("sheet1")
        target_controller = self._get_active_controller("sheet2")
        if not target_controller:
            messagebox.showwarning("Warning", "Worksheet2 has not been scanned yet.")
            return
        self.sync_formulas(source_controller, target_controller, "Worksheet1", "Worksheet2")

    def sync_2_to_1(self):
        source_controller = self._get_active_controller("sheet2")
        target_controller = self._get_active_controller("sheet1")
        if not source_controller:
            messagebox.showwarning("Warning", "Worksheet2 has not been scanned yet.")
            return
        self.sync_formulas(source_controller, target_controller, "Worksheet2", "Worksheet1")

    def sync_formulas(self, source, target, source_name, target_name):
        if not source.all_formulas:
            messagebox.showwarning("Warning", f"No formulas found in {source_name}. Please scan first.")
            return

        if not target.all_formulas:
            messagebox.showwarning("Warning", f"No formulas found in {target_name}. Please scan first.")
            return

        target.all_formulas = source.all_formulas.copy()
        target.cell_addresses = source.cell_addresses.copy()
        # We need to call apply_filter on the target's view
        from core.worksheet_tree import apply_filter
        apply_filter(target) # apply_filter now takes controller as argument
        messagebox.showinfo("Success", f"Synced {len(source.all_formulas)} formulas from {source_name} to {target_name}")

    def _get_active_controller(self, sheet=None):
        if sheet is None:
            sheet = self.worksheet_var.get()

        if sheet == "sheet1":
            return self.left_controller
        else:
            if self.right_controller is None:
                current_width = self.main_window.winfo_width()
                current_height = self.main_window.winfo_height()
                self.main_window.geometry(f"{current_width * 2}x{current_height}")
                
                self.paned_window.add(self.right_frame_placeholder, weight=1)

                self.right_controller = WorksheetController(self.right_frame_placeholder, self.root, "Worksheet2")
                
                # Configure right sync button (Sync 2 -> 1)
                self.right_controller.view.sync_button.config(
                    text="Sync 2 -> 1",
                    command=self.sync_2_to_1,
                    state="normal"
                )
                
                self.main_window.update_idletasks()
            return self.right_controller
    
    def hide_worksheet2_interface(self):
        """Hide worksheet2 interface and reset window to single worksheet view"""
        if self.right_controller is not None:
            # Remove the right pane from the paned window
            self.paned_window.forget(self.right_frame_placeholder)
            
            # Destroy the right controller to free memory
            if hasattr(self.right_controller, 'view'):
                self.right_controller.view.destroy()
            self.right_controller = None
            
            # Reset window size to single worksheet size
            current_height = self.main_window.winfo_height()
            self.main_window.geometry(f"900x{current_height}")
            
            print("Hidden worksheet2 interface and reset window size")
        else:
            print("Worksheet2 interface is already hidden")