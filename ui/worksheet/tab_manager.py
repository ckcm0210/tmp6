# -*- coding: utf-8 -*-
"""
Tab Manager Module

This module handles all tab-related functionality including creating,
closing, and managing detail tabs in the worksheet interface.
"""

import tkinter as tk
from tkinter import ttk


class TabManager:
    """Manages detail tabs for worksheet references"""
    
    def __init__(self, detail_notebook):
        self.detail_notebook = detail_notebook
        self.detail_tabs = {}
        self.tab_counter = 0
        self._double_click_bound = False
        self._middle_click_bound = False
        self.create_detail_tab("Main", is_main=True)
    
    def create_detail_tab(self, tab_name, is_main=False):
        """Create a new detail tab with text widget and scrollbar"""
        # Create frame for this tab
        tab_frame = ttk.Frame(self.detail_notebook)
        tab_frame.columnconfigure(0, weight=1)
        tab_frame.rowconfigure(0, weight=1)
        
        # Create text widget with scrollbar
        text_widget = tk.Text(tab_frame, height=23, wrap=tk.WORD, font=("Consolas", 10))
        text_widget.grid(row=0, column=0, sticky="nsew")
        
        scrollbar = ttk.Scrollbar(tab_frame, command=text_widget.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # Configure text tags for formatting
        text_widget.tag_configure("label", font=("Consolas", 10, "bold"), foreground="navy")
        text_widget.tag_configure("value", font=("Consolas", 10), foreground="black")
        text_widget.tag_configure("formula_content", font=("Consolas", 10, "italic"), foreground="darkgreen")
        text_widget.tag_configure("result_value", font=("Consolas", 10), foreground="darkblue")
        text_widget.tag_configure("referenced_value", font=("Consolas", 10), foreground="purple")
        text_widget.tag_configure("info_text", font=("Consolas", 10, "italic"), foreground="grey")
        
        # Add tab to notebook
        self.detail_notebook.add(tab_frame, text=tab_name)
        
        # Store tab reference
        tab_id = f"tab_{self.tab_counter}"
        self.tab_counter += 1
        
        self.detail_tabs[tab_name] = {
            "id": tab_id,
            "frame": tab_frame,
            "text_widget": text_widget,
            "scrollbar": scrollbar,
            "is_main": is_main
        }
        
        # Add tooltip for non-main tabs to show full path info
        if not is_main:
            self._add_tab_tooltip(tab_frame, tab_name)
        
        # If not main tab, add close functionality
        if not is_main:
            self._setup_tab_close_functionality(tab_name, tab_frame)
        
        # Select the new tab
        self.detail_notebook.select(tab_frame)
        
        return text_widget
    
    def _add_tab_tooltip(self, tab_frame, tab_name):
        """Add tooltip to show full tab information"""
        def show_tooltip(event):
            try:
                # Create tooltip window
                tooltip = tk.Toplevel()
                tooltip.wm_overrideredirect(True)
                tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
                
                # Parse tab name to show detailed info
                if '|' in tab_name and '!' in tab_name:
                    parts = tab_name.split('|')
                    file_part = parts[0]
                    sheet_cell_part = parts[1]
                    if '!' in sheet_cell_part:
                        sheet_name, cell_address = sheet_cell_part.split('!', 1)
                        tooltip_text = f"File: {file_part}\nWorksheet: {sheet_name}\nCell: {cell_address}\n\nRight-click to close tab\nDouble-click to close tab\nMiddle-click to close tab"
                    else:
                        tooltip_text = f"Reference: {tab_name}\n\nRight-click to close tab\nDouble-click to close tab\nMiddle-click to close tab"
                else:
                    tooltip_text = f"Reference: {tab_name}\n\nRight-click to close tab\nDouble-click to close tab\nMiddle-click to close tab"
                
                label = tk.Label(tooltip, text=tooltip_text, background="lightyellow", 
                               relief="solid", borderwidth=1, font=("Arial", 9))
                label.pack()
                
                # Store tooltip reference to hide it later
                tab_frame.tooltip = tooltip
                
                # Auto-hide after 3 seconds
                tooltip.after(3000, lambda: tooltip.destroy() if tooltip.winfo_exists() else None)
                
            except Exception:
                pass
        
        def hide_tooltip(event):
            try:
                if hasattr(tab_frame, 'tooltip') and tab_frame.tooltip.winfo_exists():
                    tab_frame.tooltip.destroy()
            except Exception:
                pass
        
        # Bind hover events to tab frame
        tab_frame.bind("<Enter>", show_tooltip)
        tab_frame.bind("<Leave>", hide_tooltip)
    
    def _setup_tab_close_functionality(self, tab_name, tab_frame):
        """Setup multiple ways to close tabs"""
        
        # Method 1: Right-click context menu
        def close_tab():
            self.close_detail_tab(tab_name)
        
        def close_all_other_tabs():
            self.close_all_other_tabs(tab_name)
        
        def on_right_click(event):
            context_menu = tk.Menu(self.detail_notebook, tearoff=0)
            context_menu.add_command(label="Close Tab", command=close_tab)
            context_menu.add_separator()
            context_menu.add_command(label="Close All Other Tabs", command=close_all_other_tabs)
            context_menu.add_separator()
            context_menu.add_command(label="Close All Tabs (except Main)", command=self.close_all_tabs_except_main)
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()
        
        # Bind right-click to tab frame
        tab_frame.bind("<Button-3>", on_right_click)
        
        # Method 2: Double-click on tab to close
        def on_tab_double_click(event):
            try:
                # Get the tab that was clicked
                tab_id = self.detail_notebook.tk.call(self.detail_notebook._w, "identify", "tab", event.x, event.y)
                if tab_id != "":
                    clicked_tab = self.detail_notebook.tabs()[int(tab_id)]
                    # Find the tab name
                    for name, info in self.detail_tabs.items():
                        if str(info["frame"]) == clicked_tab:
                            if not info["is_main"]:
                                self.close_detail_tab(name)
                            break
            except:
                pass
        
        # Bind double-click to notebook (only bind once)
        if not self._double_click_bound:
            self.detail_notebook.bind("<Double-Button-1>", on_tab_double_click)
            self._double_click_bound = True
        
        # Method 3: Middle-click to close
        def on_middle_click(event):
            try:
                tab_id = self.detail_notebook.tk.call(self.detail_notebook._w, "identify", "tab", event.x, event.y)
                if tab_id != "":
                    clicked_tab = self.detail_notebook.tabs()[int(tab_id)]
                    for name, info in self.detail_tabs.items():
                        if str(info["frame"]) == clicked_tab:
                            if not info["is_main"]:
                                self.close_detail_tab(name)
                            break
            except:
                pass
        
        # Bind middle-click to notebook (only bind once)
        if not self._middle_click_bound:
            self.detail_notebook.bind("<Button-2>", on_middle_click)
            self._middle_click_bound = True
    
    def close_detail_tab(self, tab_name):
        """Close a detail tab (except main tab)"""
        if tab_name in self.detail_tabs and not self.detail_tabs[tab_name]["is_main"]:
            tab_frame = self.detail_tabs[tab_name]["frame"]
            self.detail_notebook.forget(tab_frame)
            del self.detail_tabs[tab_name]
            
            # Select main tab if no other tabs
            if len(self.detail_tabs) == 1:  # Only main tab left
                main_frame = self.detail_tabs["Main"]["frame"]
                self.detail_notebook.select(main_frame)
    
    def close_all_other_tabs(self, keep_tab_name):
        """Close all tabs except the specified one and main tab"""
        tabs_to_close = []
        for name, info in self.detail_tabs.items():
            if name != keep_tab_name and not info["is_main"]:
                tabs_to_close.append(name)
        
        for tab_name in tabs_to_close:
            self.close_detail_tab(tab_name)
    
    def close_all_tabs_except_main(self):
        """Close all tabs except the main tab"""
        tabs_to_close = []
        for name, info in self.detail_tabs.items():
            if not info["is_main"]:
                tabs_to_close.append(name)
        
        for tab_name in tabs_to_close:
            self.close_detail_tab(tab_name)
    
    def get_current_detail_text(self):
        """Get the currently active detail text widget"""
        current_tab = self.detail_notebook.select()
        if current_tab:
            for tab_name, tab_info in self.detail_tabs.items():
                if str(tab_info["frame"]) == current_tab:
                    return tab_info["text_widget"]
        return self.detail_tabs["Main"]["text_widget"]  # Fallback to main