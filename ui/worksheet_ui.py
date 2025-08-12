
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 09:11:40 2025

@author: kccheng
"""

import tkinter as tk
from tkinter import font
from tkinter import ttk

# Import core functions directly
from core.excel_connector import reconnect_to_excel
from core.worksheet_export import export_formulas_to_excel, import_and_update_formulas
from core.worksheet_summary import summarize_external_links
from core.worksheet_tree import apply_filter, sort_column, on_select, on_double_click

def create_ui_widgets(self):
    """Creates and places all UI widgets without binding commands."""
    self.columnconfigure(0, weight=1)
    self.rowconfigure(5, weight=1)
    self.rowconfigure(7, weight=3)
    default_content_font = ("Consolas", 10)
    main_label_font = ("Arial", 12, "bold")
    filter_label_font = ("Arial", 9, "bold")
    style = ttk.Style()
    style.configure("Treeview", font=default_content_font)
    style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
    style.configure("evenrow", background="#F0F0F0")
    style.configure("oddrow", background="#FFFFFF")
    style.configure("Toolbutton.TButton", font=("Arial", 8))

    info_frame = ttk.Frame(self)
    info_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
    info_frame.columnconfigure(1, weight=1)
    ttk.Label(info_frame, text="File Path:", font=main_label_font).grid(row=0, column=0, sticky=tk.W)
    self.path_label = ttk.Label(info_frame, text="Not Connected", foreground="red", wraplength=400)
    self.path_label.grid(row=0, column=1, sticky=tk.W)
    ttk.Label(info_frame, text="File Name:", font=main_label_font).grid(row=1, column=0, sticky=tk.W)
    self.file_label = ttk.Label(info_frame, text="Not Connected", foreground="red")
    self.file_label.grid(row=1, column=1, sticky=tk.W)
    ttk.Label(info_frame, text="Worksheet:", font=main_label_font).grid(row=2, column=0, sticky=tk.W)
    self.sheet_label = ttk.Label(info_frame, text="Not Connected", foreground="red")
    self.sheet_label.grid(row=2, column=1, sticky=tk.W)
    ttk.Label(info_frame, text="Data Range:", font=main_label_font).grid(row=3, column=0, sticky=tk.W)
    self.range_label = ttk.Label(info_frame, text="Not Connected", foreground="red")
    self.range_label.grid(row=3, column=1, sticky=tk.W)

    self.progress_frame = ttk.Frame(self)
    self.progress_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
    self.progress_frame.columnconfigure(0, weight=1)
    self.progress_label = ttk.Label(self.progress_frame, text="")
    self.progress_label.pack(fill=tk.X)
    self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
    self.progress_bar.pack(fill=tk.X, pady=(2, 0))

    filter_main_frame = ttk.LabelFrame(self, text="Filters", borderwidth=2, relief=tk.GROOVE, padding=10)
    filter_main_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
    filter_main_frame.columnconfigure(0, weight=1)
    filter_checkbox_frame = ttk.Frame(filter_main_frame)
    filter_checkbox_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
    ttk.Label(filter_checkbox_frame, text="Type:", font=filter_label_font).pack(side=tk.LEFT, padx=(0, 5))
    self.show_formula_check = ttk.Checkbutton(filter_checkbox_frame, text="Formula", variable=self.controller.show_formula)
    self.show_formula_check.pack(side=tk.LEFT, padx=5)
    self.show_local_link_check = ttk.Checkbutton(filter_checkbox_frame, text="Local Link", variable=self.controller.show_local_link)
    self.show_local_link_check.pack(side=tk.LEFT, padx=5)
    self.show_external_link_check = ttk.Checkbutton(filter_checkbox_frame, text="External Link", variable=self.controller.show_external_link)
    self.show_external_link_check.pack(side=tk.LEFT, padx=5)
    self.openpyxl_check = ttk.Checkbutton(filter_checkbox_frame, text="Enable Non-GUI File Reading for Cell Results", variable=self.controller.use_openpyxl)
    self.openpyxl_check.pack(side=tk.LEFT, padx=15)

    filter_entry_frame = ttk.Frame(filter_main_frame)
    filter_entry_frame.pack(side=tk.TOP, fill=tk.X)
    filter_entry_frame.columnconfigure(1, weight=1)
    filter_entry_frame.columnconfigure(2, weight=0)
    self.tree_columns = ("type", "address", "formula", "result", "display_value")
    self.columns_with_entries = ("address", "formula", "result", "display_value")
    self.filter_entries = {}
    column_display_names = {"address": "Address", "formula": "Formula", "result": "Result", "display_value": "Display Value"}
    row_idx = 0
    for col_id in self.columns_with_entries:
        ttk.Label(filter_entry_frame, text=f"{column_display_names[col_id]}:", font=filter_label_font).grid(row=row_idx, column=0, sticky=tk.W, padx=(5,0), pady=2)
        entry = ttk.Entry(filter_entry_frame, font=("Consolas", 10))
        entry.grid(row=row_idx, column=1, sticky=(tk.W, tk.E), padx=(0,5), pady=2)
        self.filter_entries[col_id] = entry
        btn = ttk.Button(filter_entry_frame, text="⏎", width=3)
        btn.grid(row=row_idx, column=2, padx=(0,5), pady=2)
        if col_id == 'address':
            self.controller.default_fg_color = entry.cget("foreground")
            base_font = font.Font(font=entry.cget("font"))
            self.controller.default_font = base_font
            self.controller.placeholder_font = font.Font(family=base_font.cget("family"), size=base_font.cget("size"), slant="italic")
        row_idx += 1
    self._set_placeholder()

    summary_frame = ttk.Frame(self)
    summary_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5, padx=10)
    self.summarize_button = ttk.Button(summary_frame, text="Summarize External Links")
    self.summarize_button.pack(side=tk.LEFT, padx=(0, 5))
    self.export_button = ttk.Button(summary_frame, text="Export and Open List")
    self.export_button.pack(side=tk.LEFT, padx=5)
    self.import_button = ttk.Button(summary_frame, text="Import and Update Formulas")
    self.import_button.pack(side=tk.LEFT, padx=5)
    self.reconnect_button = ttk.Button(summary_frame, text="Reconnect")
    self.reconnect_button.pack(side=tk.LEFT, padx=5)
    
    # Add sync button next to reconnect (will be configured by comparator)
    self.sync_button = ttk.Button(summary_frame, text="Sync", state="disabled")
    self.sync_button.pack(side=tk.LEFT, padx=5)

    self.formula_list_label = ttk.Label(self, text="Formula List:", font=main_label_font)
    self.formula_list_label.grid(row=4, column=0, sticky=tk.W, pady=(10, 0))
    tree_frame = ttk.Frame(self)
    tree_frame.grid(row=5, column=0, sticky="nsew")
    tree_frame.columnconfigure(0, weight=1)
    tree_frame.rowconfigure(0, weight=1)
    self.result_tree = ttk.Treeview(tree_frame, columns=self.tree_columns, show="headings", height=12)
    headings = {"type": "Type", "address": "Address", "formula": "Formula Content", "result": "Result", "display_value": "Display Value"}
    widths = {"type": 70, "address": 70, "formula": 400, "result": 120, "display_value": 120}
    for col_id, text in headings.items():
        self.result_tree.heading(col_id, text=text)
        self.result_tree.column(col_id, width=widths[col_id], minwidth=60)
    self.result_tree.grid(row=0, column=0, sticky="nsew")
    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
    scrollbar.grid(row=0, column=1, sticky="ns")
    self.result_tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.configure(command=self.result_tree.yview)

    detail_header_frame = ttk.Frame(self)
    detail_header_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
    ttk.Label(detail_header_frame, text="Details:", font=main_label_font).pack(side=tk.LEFT)
    self.close_tabs_button = ttk.Button(detail_header_frame, text="Close All Tabs", style="Toolbutton.TButton")
    self.close_tabs_button.pack(side=tk.RIGHT)

    detail_frame = ttk.Frame(self)
    detail_frame.grid(row=7, column=0, sticky="nsew")
    detail_frame.columnconfigure(0, weight=1)
    detail_frame.rowconfigure(0, weight=1)
    self.detail_notebook = ttk.Notebook(detail_frame)
    self.detail_notebook.grid(row=0, column=0, sticky="nsew")
    
    self.ui_initialized = True

def bind_ui_commands(self):
    """Binds all commands and events to the UI widgets."""
    self.show_formula_check.config(command=lambda: apply_filter(self.controller))
    self.show_local_link_check.config(command=lambda: apply_filter(self.controller))
    self.show_external_link_check.config(command=lambda: apply_filter(self.controller))
    self.openpyxl_check.config(command=lambda: on_select(self.controller, event=None))

    for col_id, entry in self.filter_entries.items():
        entry.bind("<Return>", lambda event, s=self.controller: apply_filter(s, event))
        if col_id == 'address':
            entry.bind("<FocusIn>", self._on_focus_in)
            entry.bind("<FocusOut>", self._on_focus_out)
        entry.master.children['!button'].config(command=lambda s=self.controller: apply_filter(s))

    self.summarize_button.config(command=lambda: summarize_external_links(self.controller))
    self.export_button.config(command=lambda: export_formulas_to_excel(self.controller))
    self.import_button.config(command=lambda: import_and_update_formulas(self.controller))
    self.reconnect_button.config(command=lambda: reconnect_to_excel(self.controller))

    for col_id in self.tree_columns:
        self.result_tree.heading(col_id, command=lambda c=col_id, s=self.controller: sort_column(s, c))
    
    self.result_tree.bind("<Double-Button-1>", lambda event, s=self.controller: on_double_click(s, event))
    self.result_tree.bind("<<TreeviewSelect>>", lambda event, s=self.controller: on_select(s, event))

    self.close_tabs_button.config(command=self.controller.tab_manager.close_all_tabs_except_main)

def _set_placeholder(self):
    entry = self.filter_entries.get('address')
    if entry:
        entry.delete(0, tk.END)
        entry.insert(0, self.controller.placeholder_text)
        entry.config(foreground=self.controller.placeholder_color, font=self.controller.placeholder_font)
        entry.icursor(0)

def _on_focus_in(self, event):
    entry = event.widget
    if entry.get() == self.controller.placeholder_text:
        entry.delete(0, tk.END)
        entry.config(foreground=self.controller.default_fg_color, font=self.controller.default_font)

def _on_mouse_click(self, event):
    entry = event.widget
    if entry.get() == self.controller.placeholder_text:
        entry.delete(0, tk.END)
        entry.config(foreground=self.controller.default_fg_color, font=self.controller.default_font)
        return "break"

def _on_focus_out(self, event):
    entry = event.widget
    if not entry.get():
        entry.insert(0, self.controller.placeholder_text)
        entry.config(foreground=self.controller.placeholder_color, font=self.controller.placeholder_font)
