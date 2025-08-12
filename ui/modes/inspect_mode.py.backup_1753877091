# -*- coding: utf-8 -*-
"""
Inspect Mode UI Module

This module creates a simplified version of the worksheet functionality
for Inspect Mode, reusing existing components but hiding unnecessary elements.
"""

import tkinter as tk
from tkinter import ttk
from ui.worksheet.controller import WorksheetController

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
        
        # Create PanedWindow for resizable panes
        self.paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill='both', expand=True)
        
        # Create left pane
        left_frame = ttk.LabelFrame(self.paned_window, text="Left Pane", padding=5)
        self.paned_window.add(left_frame, weight=1)
        
        # Create right pane  
        right_frame = ttk.LabelFrame(self.paned_window, text="Right Pane", padding=5)
        self.paned_window.add(right_frame, weight=1)
        
        # Create simplified worksheet controllers
        self.left_controller = SimplifiedWorksheetController(left_frame, self.root, "Left")
        self.right_controller = SimplifiedWorksheetController(right_frame, self.root, "Right")

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
            # Adjust formula list height to show one result row (not just column headers)
            if hasattr(self.view, 'result_tree'):
                # Height=2 means 1 header row + 1 data row
                self.view.result_tree.configure(height=2)
                print(f"Modified formula list height to show one result row in {self.pane_name}")
            
            # Add scan button for current Excel selection
            self.add_scan_current_selection_button()
            
        except Exception as e:
            print(f"Warning: Could not modify layout in {self.pane_name}: {e}")
    
    def add_scan_current_selection_button(self):
        """Add a button to scan the currently selected cell in Excel"""
        try:
            # Create a frame for the scan button using grid (since WorksheetView uses grid)
            scan_frame = ttk.Frame(self.view)
            
            # Use grid to place it at the top (row 0)
            scan_frame.grid(row=0, column=0, columnspan=10, sticky='ew', pady=5, padx=5)
            
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
                    # Get the address of the selected cell/range
                    selected_address = selected_range.Address
                    
                    # For Inspect Mode, we focus on single cells
                    # If a range is selected, use the first cell
                    if ':' in selected_address:
                        # Take only the first cell from the range
                        first_cell = selected_address.split(':')[0]
                        selected_address = first_cell
                    
                    # Remove $ signs for cleaner display
                    clean_address = selected_address.replace('$', '')
                    
                    # Set up for selected range scanning (reuse existing functionality)
                    self.selected_scan_address = selected_address
                    self.selected_scan_count = 1
                    self.scanning_selected_range = True
                    self.original_user_selection = selected_address
                    self.original_user_count = 1
                    
                    # Update range label
                    self.view.range_label.config(text=f"Selected Cell ({clean_address})", foreground="black")
                    
                    # Use the same logic as the "Selected Range" button in Normal Mode
                    # Found that "Selected Range" button uses command=self.scan_worksheet_selected
                    
                    # Set up selected range scanning parameters
                    self.selected_scan_address = selected_address
                    self.selected_scan_count = 1
                    self.scanning_selected_range = True
                    self.original_user_selection = selected_address
                    self.original_user_count = 1
                    
                    # Update range label to show what we're scanning
                    self.view.range_label.config(text=f"Selected Cell ({clean_address})", foreground="black")
                    
                    # Call the same function that "Selected Range" button uses
                    # But we need to call refresh_data directly since scan_worksheet_selected is a class method
                    from formula_comparator import refresh_data
                    
                    # Create a temporary button reference for the refresh_data function
                    temp_button = ttk.Button(self.view, text="Scanning...")
                    temp_button.pack_forget()  # Hide it immediately
                    
                    # Call refresh_data with selected mode (same as scan_worksheet_selected does)
                    refresh_data(self, temp_button, scan_mode="selected")
                    
                    print(f"Scanned selected cell {clean_address} in {self.pane_name}")
                    
                else:
                    messagebox.showwarning("No Selection", "Please select a cell in Excel first.")
                    
            except Exception as e:
                messagebox.showerror("Scan Error", f"Could not scan selected cell: {e}")
                
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Connection Error", f"Could not connect to Excel: {e}")
    
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