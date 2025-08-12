# -*- coding: utf-8 -*-
"""
Dual Pane Controller Module

This module manages the dual-pane layout for Inspect Mode.
Each pane can independently scan and analyze a single cell from different Excel files.
"""

import tkinter as tk
from tkinter import ttk
from ui.worksheet.controller import WorksheetController
import win32com.client

class DualPaneController:
    """Manages dual-pane layout for Inspect Mode"""
    
    def __init__(self, parent_frame, root_app):
        """
        Initialize the Dual Pane Controller
        
        Args:
            parent_frame: The parent tkinter frame
            root_app: The main application window
        """
        self.parent = parent_frame
        self.root = root_app
        
        # Create main container
        self.main_frame = ttk.Frame(self.parent)
        self.main_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create left and right pane frames
        self.left_frame = None
        self.right_frame = None
        
        # Pane controllers
        self.left_controller = None
        self.right_controller = None
        
        self.setup_dual_pane_layout()
    
    def setup_dual_pane_layout(self):
        """Setup the dual-pane layout"""
        # Create a PanedWindow for resizable panes
        self.paned_window = ttk.PanedWindow(self.main_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill='both', expand=True)
        
        # Create left pane
        self.left_frame = ttk.LabelFrame(self.paned_window, text="Left Pane", padding=5)
        self.paned_window.add(self.left_frame, weight=1)
        
        # Create right pane
        self.right_frame = ttk.LabelFrame(self.paned_window, text="Right Pane", padding=5)
        self.paned_window.add(self.right_frame, weight=1)
        
        # Initialize pane controllers
        self.left_controller = InspectPaneController(self.left_frame, self.root, "Left")
        self.right_controller = InspectPaneController(self.right_frame, self.root, "Right")
    
    def get_left_controller(self):
        """Get the left pane controller"""
        return self.left_controller
    
    def get_right_controller(self):
        """Get the right pane controller"""
        return self.right_controller
    
    def reset_both_panes(self):
        """Reset both panes to initial state"""
        if self.left_controller:
            self.left_controller.reset_pane()
        if self.right_controller:
            self.right_controller.reset_pane()

class InspectPaneController:
    """Controller for a single pane in Inspect Mode"""
    
    def __init__(self, parent_frame, root_app, pane_name):
        """
        Initialize the Inspect Pane Controller
        
        Args:
            parent_frame: The parent tkinter frame
            root_app: The main application window
            pane_name: Name of the pane ("Left" or "Right")
        """
        self.parent = parent_frame
        self.root = root_app
        self.pane_name = pane_name
        
        # Excel connection
        self.xl = None
        self.workbook = None
        self.worksheet = None
        
        # Current scan data
        self.current_cell_address = None
        self.current_cell_data = None
        
        self.setup_pane_ui()
    
    def setup_pane_ui(self):
        """Setup the UI for this pane"""
        # Connection info frame
        info_frame = ttk.LabelFrame(self.parent, text="Connection Info", padding=5)
        info_frame.pack(fill='x', pady=(0, 5))
        
        # File path
        ttk.Label(info_frame, text="File:").grid(row=0, column=0, sticky='w', padx=(0, 5))
        self.file_label = ttk.Label(info_frame, text="Not Connected", foreground="red")
        self.file_label.grid(row=0, column=1, sticky='w')
        
        # Worksheet
        ttk.Label(info_frame, text="Sheet:").grid(row=1, column=0, sticky='w', padx=(0, 5))
        self.sheet_label = ttk.Label(info_frame, text="Not Connected", foreground="red")
        self.sheet_label.grid(row=1, column=1, sticky='w')
        
        # Cell address
        ttk.Label(info_frame, text="Cell:").grid(row=2, column=0, sticky='w', padx=(0, 5))
        self.cell_label = ttk.Label(info_frame, text="Not Scanned", foreground="red")
        self.cell_label.grid(row=2, column=1, sticky='w')
        
        # Scan controls frame
        scan_frame = ttk.LabelFrame(self.parent, text="Scan Controls", padding=5)
        scan_frame.pack(fill='x', pady=(0, 5))
        
        # Cell address input
        ttk.Label(scan_frame, text="Cell Address:").grid(row=0, column=0, sticky='w', padx=(0, 5))
        self.cell_entry = ttk.Entry(scan_frame, width=10)
        self.cell_entry.grid(row=0, column=1, padx=(0, 5))
        self.cell_entry.insert(0, "A1")  # Default value
        
        # Scan button
        self.scan_btn = ttk.Button(scan_frame, text="Scan Cell", command=self.scan_current_cell)
        self.scan_btn.grid(row=0, column=2, padx=5)
        
        # Connect button
        self.connect_btn = ttk.Button(scan_frame, text="Connect to Excel", command=self.connect_to_excel)
        self.connect_btn.grid(row=1, column=0, columnspan=3, pady=(5, 0), sticky='ew')
        
        # Results frame
        results_frame = ttk.LabelFrame(self.parent, text="Cell Analysis", padding=5)
        results_frame.pack(fill='both', expand=True)
        
        # Create text widget for results
        self.results_text = tk.Text(results_frame, wrap=tk.WORD, height=10)
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        self.results_text.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar.pack(side=tk.RIGHT, fill='y')
        
        # Configure text tags for formatting
        self.results_text.tag_configure("header", font=("Arial", 10, "bold"))
        self.results_text.tag_configure("label", font=("Arial", 9, "bold"))
        self.results_text.tag_configure("value", font=("Consolas", 9))
        self.results_text.tag_configure("formula", font=("Consolas", 9), foreground="blue")
        self.results_text.tag_configure("error", foreground="red")
    
    def connect_to_excel(self):
        """Connect to Excel and get active workbook/worksheet"""
        try:
            import win32com.client
            
            # Try to connect to Excel
            try:
                self.xl = win32com.client.GetActiveObject("Excel.Application")
            except:
                self.xl = win32com.client.Dispatch("Excel.Application")
                self.xl.Visible = True
            
            # Get active workbook and worksheet
            self.workbook = self.xl.ActiveWorkbook
            self.worksheet = self.xl.ActiveSheet
            
            if self.workbook and self.worksheet:
                # Update UI
                self.file_label.config(text=self.workbook.Name, foreground="black")
                self.sheet_label.config(text=self.worksheet.Name, foreground="black")
                self.connect_btn.config(text="Connected", state="disabled")
                
                # Enable scan button
                self.scan_btn.config(state="normal")
                
                self.show_message("✅ Connected to Excel successfully!")
            else:
                self.show_error("❌ No active workbook or worksheet found")
                
        except Exception as e:
            self.show_error(f"❌ Connection failed: {str(e)}")
    
    def scan_current_cell(self):
        """Scan the specified cell"""
        if not self.xl or not self.workbook or not self.worksheet:
            self.show_error("❌ Please connect to Excel first")
            return
        
        cell_address = self.cell_entry.get().strip().upper()
        if not cell_address:
            self.show_error("❌ Please enter a cell address")
            return
        
        try:
            # Get the cell
            cell = self.worksheet.Range(cell_address)
            
            # Extract cell information
            cell_data = {
                'address': cell_address,
                'value': cell.Value,
                'formula': cell.Formula if hasattr(cell, 'Formula') else None,
                'display_text': cell.Text if hasattr(cell, 'Text') else str(cell.Value),
                'has_formula': bool(cell.Formula and str(cell.Formula).startswith('=')),
                'workbook_name': self.workbook.Name,
                'worksheet_name': self.worksheet.Name
            }
            
            # Store current data
            self.current_cell_address = cell_address
            self.current_cell_data = cell_data
            
            # Update UI
            self.cell_label.config(text=cell_address, foreground="black")
            
            # Display results
            self.display_cell_analysis(cell_data)
            
            self.show_message(f"✅ Cell {cell_address} scanned successfully!")
            
        except Exception as e:
            self.show_error(f"❌ Scan failed: {str(e)}")
    
    def display_cell_analysis(self, cell_data):
        """Display detailed cell analysis"""
        # Clear previous results
        self.results_text.delete(1.0, tk.END)
        
        # Header
        self.results_text.insert(tk.END, f"Cell Analysis: {cell_data['address']}\n", "header")
        self.results_text.insert(tk.END, "=" * 40 + "\n\n")
        
        # Basic information
        self.results_text.insert(tk.END, "Basic Information:\n", "label")
        self.results_text.insert(tk.END, f"  Workbook: {cell_data['workbook_name']}\n", "value")
        self.results_text.insert(tk.END, f"  Worksheet: {cell_data['worksheet_name']}\n", "value")
        self.results_text.insert(tk.END, f"  Cell Address: {cell_data['address']}\n", "value")
        self.results_text.insert(tk.END, f"  Display Text: {cell_data['display_text']}\n", "value")
        self.results_text.insert(tk.END, f"  Calculated Value: {cell_data['value']}\n", "value")
        self.results_text.insert(tk.END, "\n")
        
        # Formula information
        if cell_data['has_formula']:
            self.results_text.insert(tk.END, "Formula Analysis:\n", "label")
            self.results_text.insert(tk.END, f"  Formula: {cell_data['formula']}\n", "formula")
            
            # Analyze formula type
            formula = str(cell_data['formula'])
            if '[' in formula and ']' in formula:
                self.results_text.insert(tk.END, "  Type: External Link\n", "value")
            elif '!' in formula:
                self.results_text.insert(tk.END, "  Type: Local Link\n", "value")
            else:
                self.results_text.insert(tk.END, "  Type: Formula\n", "value")
        else:
            self.results_text.insert(tk.END, "Formula Analysis:\n", "label")
            self.results_text.insert(tk.END, "  Type: Value (No Formula)\n", "value")
        
        self.results_text.insert(tk.END, "\n")
        
        # TODO: Add more detailed analysis
        # - Referenced cells
        # - External links
        # - Dependencies
        self.results_text.insert(tk.END, "Detailed Analysis:\n", "label")
        self.results_text.insert(tk.END, "  [More analysis features coming soon...]\n", "value")
    
    def show_message(self, message):
        """Show a message in the results area"""
        self.results_text.insert(tk.END, f"\n{message}\n", "value")
        self.results_text.see(tk.END)
    
    def show_error(self, error_message):
        """Show an error message in the results area"""
        self.results_text.insert(tk.END, f"\n{error_message}\n", "error")
        self.results_text.see(tk.END)
    
    def reset_pane(self):
        """Reset this pane to initial state"""
        # Reset connection
        self.xl = None
        self.workbook = None
        self.worksheet = None
        
        # Reset data
        self.current_cell_address = None
        self.current_cell_data = None
        
        # Reset UI
        self.file_label.config(text="Not Connected", foreground="red")
        self.sheet_label.config(text="Not Connected", foreground="red")
        self.cell_label.config(text="Not Scanned", foreground="red")
        self.cell_entry.delete(0, tk.END)
        self.cell_entry.insert(0, "A1")
        self.connect_btn.config(text="Connect to Excel", state="normal")
        self.scan_btn.config(state="disabled")
        self.results_text.delete(1.0, tk.END)
        
        self.show_message("Pane reset to initial state")