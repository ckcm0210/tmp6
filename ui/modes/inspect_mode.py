# -*- coding: utf-8 -*-
"""
Inspect Mode UI Module

This module creates a simplified version of the worksheet functionality
for Inspect Mode, reusing existing components but hiding unnecessary elements.
"""

import tkinter as tk
from tkinter import ttk
from ui.worksheet.controller import WorksheetController
import traceback
import win32com.client
from tkinter import messagebox
import re
import time
from core.excel_scanner import refresh_data
from core.worksheet_tree import on_select

class InspectModeView:
    """Simplified worksheet view for Inspect Mode"""
    
    def __init__(self, parent_frame, root_app):
        self.parent = parent_frame
        self.root = root_app
        
        # Create dual pane layout
        self.setup_dual_pane_layout()
    
    def setup_dual_pane_layout(self):
        """Setup dual-pane layout with simplified worksheet controllers"""
        # Create main container
        main_frame = ttk.Frame(self.parent)
        main_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create control frame for toggle button
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill='x', pady=(0, 5))
        
        # Add toggle button for right pane
        self.right_pane_visible = tk.BooleanVar(value=False)  # Default hide right pane
        self.toggle_btn = ttk.Button(
            control_frame, 
            text="Show Right Pane", 
            command=self.toggle_right_pane
        )
        self.toggle_btn.pack(side=tk.LEFT, padx=5)
        
        # Create PanedWindow for resizable panes
        self.paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill='both', expand=True)
        
        # Create left pane 
        self.left_frame = ttk.LabelFrame(self.paned_window, text="Left Pane", padding=5)
        self.paned_window.add(self.left_frame, weight=1)  # In single mode, left takes full space
        
        # Create right pane (initially hidden)
        self.right_frame = ttk.LabelFrame(self.paned_window, text="Right Pane", padding=5)
        
        # Create empty spacer frame (only used in dual pane mode)
        self.spacer_frame = ttk.Frame(self.paned_window)
        # Don't add spacer initially - single pane mode doesn't need it
        
        # Create simplified worksheet controllers
        self.left_controller = SimplifiedWorksheetController(self.left_frame, self.root, "Left")
        self.right_controller = SimplifiedWorksheetController(self.right_frame, self.root, "Right")
        
        # Set initial single pane mode and window size
        self.update_layout_mode()
        
        # Set initial window size to single-pane mode (half width, reduced height)
        try:
            # Reduce height by 1/5: 800 * 0.8 = 640
            self.root.geometry("600x640")
        except:
            pass
    
    def toggle_right_pane(self):
        """Toggle the visibility of the right pane"""
        self.right_pane_visible.set(not self.right_pane_visible.get())
        self.update_layout_mode()
    
    def update_layout_mode(self):
        """Update the layout based on right pane visibility"""
        if self.right_pane_visible.get():
            # Show right pane - dual pane mode (restore original full width)
            self.paned_window.add(self.right_frame, weight=1)
            self.toggle_btn.config(text="Hide Right Pane")
            
            # Get current window geometry to preserve height and position
            try:
                current_geometry = self.root.geometry()
                if 'x' in current_geometry and '+' in current_geometry:
                    # Format: "widthxheight+x+y"
                    width_height, position = current_geometry.split('+', 1)
                    if 'x' in width_height:
                        width, height = width_height.split('x')
                        # Double the width to restore original dual-pane size
                        new_width = int(width) * 2
                        self.root.geometry(f"{new_width}x{height}+{position}")
                    else:
                        # Fallback if parsing fails - reduced height
                        self.root.geometry("1200x640")
                else:
                    # Simple format: "widthxheight"
                    if 'x' in current_geometry:
                        width, height = current_geometry.split('x')
                        new_width = int(width) * 2
                        self.root.geometry(f"{new_width}x{height}")
                    else:
                        self.root.geometry("1200x640")
            except Exception as e:
                print(f"Error adjusting window size: {e}")
                # Fallback to reasonable dual-pane size - reduced height
                self.root.geometry("1200x640")
                
        else:
            # Hide right pane - single pane mode (half width)
            try:
                self.paned_window.remove(self.right_frame)
            except:
                pass
            self.toggle_btn.config(text="Show Right Pane")
            
            # Get current window geometry to preserve height and position  
            try:
                current_geometry = self.root.geometry()
                if 'x' in current_geometry and '+' in current_geometry:
                    # Format: "widthxheight+x+y"
                    width_height, position = current_geometry.split('+', 1)
                    if 'x' in width_height:
                        width, height = width_height.split('x')
                        # Halve the width for single-pane mode
                        new_width = int(width) // 2
                        self.root.geometry(f"{new_width}x{height}+{position}")
                    else:
                        # Fallback if parsing fails - reduced height
                        self.root.geometry("600x640")
                else:
                    # Simple format: "widthxheight"
                    if 'x' in current_geometry:
                        width, height = current_geometry.split('x')
                        new_width = int(width) // 2
                        self.root.geometry(f"{new_width}x{height}")
                    else:
                        self.root.geometry("600x640")
            except Exception as e:
                print(f"Error adjusting window size: {e}")
                # Fallback to reasonable single-pane size - reduced height
                self.root.geometry("600x640")

class SimplifiedWorksheetController(WorksheetController):
    """Simplified version of WorksheetController for Inspect Mode"""
    
    def __init__(self, parent_frame, root_app, pane_name):
        # Initialize with modified pane name for Inspect Mode
        super().__init__(parent_frame, root_app, f"Inspect-{pane_name}")
        
        # Use after_idle to ensure UI is fully created before hiding elements
        self.view.after_idle(self.setup_inspect_mode_ui)
    
    def setup_inspect_mode_ui(self):
        """Setup Inspect Mode UI after the view is fully initialized"""
        self.hide_unnecessary_elements()
        self.modify_layout_for_inspect_mode()
    
    def hide_unnecessary_elements(self):
        """Hide UI elements that are not needed in Inspect Mode"""
        try:
            # Hide progress frame (progress bar and label) - use grid_forget since it uses grid
            if hasattr(self.view, 'progress_frame'):
                self.view.progress_frame.grid_forget()
                print(f"Actually hidden progress frame in {self.pane_name}")
            
            # Find and hide all unwanted widgets by checking all children recursively
            self._hide_widgets_recursively(self.view)
            
        except Exception as e:
            print(f"Warning: Could not hide some UI elements in {self.pane_name}: {e}")
    
    def _hide_widgets_recursively(self, parent_widget):
        """Recursively find and hide unwanted widgets"""
        try:
            for widget in parent_widget.winfo_children():
                # Check LabelFrame for Filters
                if isinstance(widget, ttk.LabelFrame):
                    try:
                        widget_text = str(widget.cget('text')).lower()
                        if 'filter' in widget_text:
                            widget.grid_forget()
                            widget.pack_forget()
                            print(f"Actually hidden filter frame in {self.pane_name}")
                            continue
                    except:
                        pass
                
                # Check Frame for unwanted buttons
                if isinstance(widget, ttk.Frame):
                    try:
                        has_unwanted_buttons = False
                        for child in widget.winfo_children():
                            if isinstance(child, ttk.Button):
                                button_text = str(child.cget('text')).lower()
                                unwanted_keywords = ['summarize', 'export', 'import', 'reconnect']
                                if any(keyword in button_text for keyword in unwanted_keywords):
                                    has_unwanted_buttons = True
                                    break
                        
                        if has_unwanted_buttons:
                            widget.grid_forget()
                            widget.pack_forget()
                            print(f"Actually hidden summary buttons frame in {self.pane_name}")
                            continue
                    except:
                        pass
                
                # Recursively check children
                self._hide_widgets_recursively(widget)
                
        except Exception as e:
            print(f"Warning in recursive hide for {self.pane_name}: {e}")
    
    def modify_layout_for_inspect_mode(self):
        """Modify the layout for Inspect Mode requirements"""
        try:
            # Fix layout spacing issues - remove unnecessary padding/spacing
            self._fix_layout_spacing()
            
            # Adjust formula list height to show one result row (not just column headers)
            if hasattr(self.view, 'result_tree'):
                # Height=2 means 1 header row + 1 data row
                self.view.result_tree.configure(height=2)
                print(f"Modified formula list height to show one result row in {self.pane_name}")
            
            # Add scan button for current Excel selection
            self.add_scan_current_selection_button()
            
        except Exception as e:
            print(f"Warning: Could not modify layout in {self.pane_name}: {e}")
    
    def _fix_layout_spacing(self):
        """Fix layout spacing issues - remove gray spaces and make components tight"""
        try:
            # Find and configure all widgets to remove spacing
            self._configure_tight_spacing(self.view, level=0)
            
            # Find the details area (notebook) and configure it to expand
            self._configure_details_expansion(self.view)
            
            print(f"Fixed layout spacing in {self.pane_name}")
            
        except Exception as e:
            print(f"Warning: Could not fix layout spacing in {self.pane_name}: {e}")
    
    def _configure_tight_spacing(self, widget, level=0):
        """Recursively configure widgets for tight spacing"""
        try:
            # Configure current widget
            grid_info = widget.grid_info()
            if grid_info:
                # Remove padding for most widgets, but keep minimal for readability
                if level == 0:  # Top level - keep some padding
                    widget.grid_configure(pady=2, padx=2)
                else:  # Nested widgets - remove padding
                    widget.grid_configure(pady=0, padx=0)
            
            # Recursively configure children
            for child in widget.winfo_children():
                self._configure_tight_spacing(child, level + 1)
                
        except Exception as e:
            pass  # Ignore errors for individual widgets
    
    def _configure_details_expansion(self, widget):
        """Find and configure the details notebook to expand properly"""
        try:
            for child in widget.winfo_children():
                # Check if this is a notebook (details area)
                if isinstance(child, ttk.Notebook):
                    # Configure the notebook to expand
                    child.grid_configure(sticky='nsew')
                    
                    # Configure the parent to give weight to this row/column
                    parent = child.master
                    grid_info = child.grid_info()
                    if grid_info:
                        row = grid_info.get('row', 0)
                        col = grid_info.get('column', 0)
                        parent.grid_rowconfigure(row, weight=1)
                        parent.grid_columnconfigure(col, weight=1)
                    
                    print(f"Configured details notebook expansion in {self.pane_name}")
                    return True
                
                # Recursively search in children
                if self._configure_details_expansion(child):
                    return True
                    
        except Exception as e:
            pass
        
        return False
    
    def add_scan_current_selection_button(self):
        """Add a button to scan the currently selected cell in Excel"""
        try:
            # Create a frame for the scan button using grid (since WorksheetView uses grid)
            scan_frame = ttk.Frame(self.view)
            
            # Use grid to place it at the top (row 0)
            scan_frame.grid(row=0, column=0, columnspan=10, sticky='ew', pady=2, padx=5)
            
            # Add scan button (similar to Selected Range functionality in Normal Mode)
            scan_btn = ttk.Button(
                scan_frame,
                text="Scan Selected Cell",
                command=self.scan_selected_cell
            )
            scan_btn.pack(side=tk.LEFT, padx=5)
            
            # Add Close All Tabs button (same as Normal Mode)
            close_tabs_btn = ttk.Button(
                scan_frame,
                text="Close All Tabs",
                command=self.close_all_tabs
            )
            close_tabs_btn.pack(side=tk.LEFT, padx=5)
            
            # Shift all other widgets down by updating their row numbers
            self._shift_existing_widgets_down()
            
            # Configure main view grid weights after adding scan frame
            self._configure_main_view_grid()
            
            print(f"Successfully added scan button in {self.pane_name}")
            
        except Exception as e:
            print(f"Warning: Could not add scan button in {self.pane_name}: {e}")
            import traceback
            traceback.print_exc()
    
    def _shift_existing_widgets_down(self):
        """Shift existing widgets down to make room for the scan button"""
        try:
            # Get all widgets and their grid info
            for widget in self.view.winfo_children():
                if widget != self.view.winfo_children()[-1]:  # Skip the scan frame we just added
                    try:
                        grid_info = widget.grid_info()
                        if grid_info and 'row' in grid_info:
                            current_row = int(grid_info['row'])
                            # Move everything down by 1 row
                            widget.grid_configure(row=current_row + 1)
                    except:
                        pass
        except Exception as e:
            print(f"Warning: Could not shift widgets in {self.pane_name}: {e}")
    
    def _configure_main_view_grid(self):
        """Configure main view grid weights to ensure proper expansion"""
        try:
            # Configure column to expand
            self.view.grid_columnconfigure(0, weight=1)
            
            # Find the details area row and configure it to expand
            max_row = 0
            details_row = None
            
            for widget in self.view.winfo_children():
                try:
                    grid_info = widget.grid_info()
                    if grid_info and 'row' in grid_info:
                        row = int(grid_info['row'])
                        max_row = max(max_row, row)
                        
                        # Check if this widget contains the details notebook
                        if self._widget_contains_notebook(widget):
                            details_row = row
                            
                except:
                    pass
            
            # Configure the details row to expand, or use the last row if not found
            if details_row is not None:
                self.view.grid_rowconfigure(details_row, weight=1)
                print(f"Configured row {details_row} for details expansion in {self.pane_name}")
            else:
                # Fallback: configure the last row to expand
                self.view.grid_rowconfigure(max_row, weight=1)
                print(f"Configured fallback row {max_row} for expansion in {self.pane_name}")
                
        except Exception as e:
            print(f"Warning: Could not configure main view grid in {self.pane_name}: {e}")
    
    def _widget_contains_notebook(self, widget):
        """Check if a widget contains a notebook (details area)"""
        try:
            if isinstance(widget, ttk.Notebook):
                return True
            
            for child in widget.winfo_children():
                if self._widget_contains_notebook(child):
                    return True
                    
        except:
            pass
        
        return False
    
    def scan_selected_cell(self):
        """Scan the currently selected cell in Excel (similar to Selected Range in Normal Mode)"""
        try:
            import win32com.client
            from tkinter import messagebox
            
            # Try to connect to Excel
            try:
                self.xl = win32com.client.GetActiveObject("Excel.Application")
            except:
                try:
                    self.xl = win32com.client.Dispatch("Excel.Application")
                    self.xl.Visible = True
                except Exception as e:
                    messagebox.showerror("Excel Error", f"Could not connect to Excel: {e}")
                    return
            
            # Get active workbook and worksheet
            try:
                self.workbook = self.xl.ActiveWorkbook
                self.worksheet = self.xl.ActiveSheet
                
                if not self.workbook or not self.worksheet:
                    messagebox.showerror("Excel Error", "No active workbook or worksheet found.")
                    return
                
                # Store connection state for Go to Reference functionality
                self.last_workbook_path = self.workbook.FullName
                self.last_worksheet_name = self.worksheet.Name
                
                # Update UI labels
                self.view.file_label.config(text=self.workbook.Name, foreground="black")
                self.view.path_label.config(text=self.workbook.Path, foreground="black")
                self.view.sheet_label.config(text=self.worksheet.Name, foreground="black")
                
            except Exception as e:
                messagebox.showerror("Excel Error", f"Could not access Excel workbook: {e}")
                return
            
            # Get the currently selected cell
            try:
                selected_range = self.xl.Selection
                if selected_range:
                    # Use EXACTLY the same logic as Normal Mode's scan_worksheet_selected
                    selected_address = selected_range.Address.replace('$', '')
                    cell_count = selected_range.Count
                    
                    original_selected_address = selected_address
                    original_cell_count = cell_count
                    
                    # Apply the same single cell expansion logic as Normal Mode
                    if cell_count == 1:
                        try:
                            import re
                            match = re.match(r'([A-Z]+)(\d+)', selected_address)
                            if match:
                                col_letters = match.group(1)
                                row_num = int(match.group(2))
                                expanded_address = f"{col_letters}{row_num}:{col_letters}{row_num + 1}"
                                selected_address = expanded_address
                                cell_count = 2
                        except Exception as e:
                            pass
                    
                    # Set up scanning parameters EXACTLY like Normal Mode
                    self.selected_scan_address = selected_address
                    self.selected_scan_count = cell_count
                    self.original_user_selection = original_selected_address
                    self.original_user_count = original_cell_count
                    self.scanning_selected_range = True
                    
                    # Update UI display
                    cell_word = "cell" if original_cell_count == 1 else "cells"
                    self.view.range_label.config(text=f"Selected Cell ({original_selected_address}) ({original_cell_count} {cell_word})", foreground="black")
                    
                    # Add the same small delay as Normal Mode
                    import time
                    time.sleep(0.1)
                    
                    # Use the same refresh_data call with the same mode as Normal Mode
                    from core.excel_scanner import refresh_data
                    
                    # Create a temporary button reference (same as Normal Mode)
                    temp_button = ttk.Button(self.view, text="Scanning...")
                    temp_button.pack_forget()  # Hide it immediately
                    
                    # Call refresh_data with "quick" mode (same default as Normal Mode)
                    refresh_data(self, temp_button, scan_mode="quick")
                    
                    print(f"Scanned selected cell {original_selected_address} in {self.pane_name}")
                    
                    # Auto-select the first result to show details in main tab (Inspect Mode feature)
                    self.view.after(100, self.auto_select_first_result)
                    
                else:
                    messagebox.showwarning("No Selection", "Please select a cell in Excel first.")
                    
            except Exception as e:
                messagebox.showerror("Scan Error", f"Could not scan selected cell: {e}")
                
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Connection Error", f"Could not connect to Excel: {e}")
    
    def auto_select_first_result(self):
        """Auto-select the first result in the tree to show details in main tab"""
        try:
            if hasattr(self.view, 'result_tree') and self.view.result_tree:
                # Get all items in the tree
                items = self.view.result_tree.get_children()
                if items:
                    # Select the first item
                    first_item = items[0]
                    self.view.result_tree.selection_set(first_item)
                    self.view.result_tree.focus(first_item)
                    
                    # Trigger the selection event to show details
                    # This will call the on_select function from worksheet_tree.py
                    # which now includes our optimized Go to Reference functionality
                    from core.worksheet_tree import on_select
                    
                    # Create a mock event object
                    class MockEvent:
                        pass
                    
                    mock_event = MockEvent()
                    on_select(self, mock_event)
                    
                    print(f"Auto-selected first result in {self.pane_name}")
                else:
                    print(f"No results to auto-select in {self.pane_name}")
            else:
                print(f"No result tree found in {self.pane_name}")
        except Exception as e:
            print(f"Warning: Could not auto-select first result in {self.pane_name}: {e}")
    
    def close_all_tabs(self):
        """Close all tabs except the main tab (same as Normal Mode)"""
        try:
            if hasattr(self, 'tab_manager') and self.tab_manager:
                # Use the same method as Normal Mode
                self.tab_manager.close_all_tabs_except_main()
                print(f"Closed all tabs in {self.pane_name}")
            else:
                print(f"No tab manager found in {self.pane_name}")
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Close Tabs Error", f"Could not close tabs: {e}")

# Create the main Inspect Mode class
class InspectMode:
    """Main Inspect Mode class that creates the dual-pane simplified interface"""
    
    def __init__(self, parent_frame, root_app):
        self.view = InspectModeView(parent_frame, root_app)
    
    def get_left_controller(self):
        """Get the left pane controller"""
        return self.view.left_controller
    
    def get_right_controller(self):
        """Get the right pane controller"""
        return self.view.right_controller