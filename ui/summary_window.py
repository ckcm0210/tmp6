import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import re
import collections
from ui.visualizer import show_visual_chart
from utils.excel_helpers import select_ranges_in_excel, replace_links_in_excel
from utils.range_optimizer import smart_range_display

class SummaryWindow(tk.Toplevel):
    def __init__(self, parent, pane, formulas_to_summarize, is_filtered):
        super().__init__(parent)
        self.parent = parent
        self.pane = pane
        self.formulas_to_summarize = formulas_to_summarize
        self.is_filtered = is_filtered

        self.did_replace = False
        self.transient(parent)
        self.grab_set()
        full_workbook_path = pane.workbook.FullName if pane.workbook and hasattr(pane.workbook, 'FullName') else 'N/A'
        self.title(f"External Link Summary for {full_workbook_path}!{pane.worksheet.Name} ({pane.pane_name})")
        self.geometry("900x700")
        self.resizable(True, True)

        self.main_frame = ttk.Frame(self, padding=10)
        self.main_frame.pack(fill='both', expand=True)
        self.main_frame.rowconfigure(1, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        self.top_frame.columnconfigure(0, weight=1)
        self.top_frame.columnconfigure(1, weight=0)

        self.button_frame = ttk.Frame(self.top_frame)
        self.button_frame.grid(row=0, column=0, sticky="ew")
        
        self.options_frame = ttk.Frame(self.top_frame)
        self.options_frame.grid(row=0, column=1, sticky="e", padx=(10, 0))
        
        self.rescan_var = tk.BooleanVar(value=True)
        rescan_check = ttk.Checkbutton(self.options_frame, text="Refresh worksheet on exit after making replacements", variable=self.rescan_var)
        rescan_check.pack(side='left')

        self.affected_cells_summary = ""
        if self.is_filtered:
            affected_addresses = []
            address_idx = self.pane.view.tree_columns.index("address")
            for formula_data in self.formulas_to_summarize:
                if len(formula_data) > address_idx:
                    affected_addresses.append(formula_data[address_idx])
            if affected_addresses:
                self.affected_cells_summary = f" - Current View ({smart_range_display(affected_addresses)})";

        self.heading_text = "External Link Path"
        if self.is_filtered:
            self.heading_text += self.affected_cells_summary

        self.tree_frame = ttk.LabelFrame(self.main_frame, text="Found External Links")
        self.tree_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        self.tree_frame.rowconfigure(0, weight=1)
        self.tree_frame.columnconfigure(0, weight=1)

        self.summary_tree = ttk.Treeview(self.tree_frame, columns=("link",), show="headings")
        self.summary_tree.heading("link", text=self.heading_text)
        self.summary_tree.column("link", width=800)
        self.summary_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.summary_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.summary_tree.configure(yscrollcommand=scrollbar.set)
        
        # Optimized external link extraction with progress indication
        self.external_path_pattern = re.compile(r"'([^']+\\[^\]]+\.(?:xlsx|xls|xlsm|xlsb)\][^']*?)'", re.IGNORECASE)
        
        # Show progress dialog for large datasets
        total_formulas = len(self.formulas_to_summarize)
        show_progress = total_formulas > 20  # Show progress for more than 20 formulas
        
        if show_progress:
            progress_window = tk.Toplevel(self.parent)
            progress_window.title("Processing External Links...")
            progress_window.geometry("400x100")
            progress_window.resizable(False, False)
            progress_window.transient(self.parent)
            progress_window.grab_set()
            
            # Center the progress window
            progress_window.update_idletasks()
            x = (progress_window.winfo_screenwidth() // 2) - (200)
            y = (progress_window.winfo_screenheight() // 2) - (50)
            progress_window.geometry(f"400x100+{x}+{y}")
            
            frame = ttk.Frame(progress_window, padding=20)
            frame.pack(fill='both', expand=True)
            
            ttk.Label(frame, text="Processing external links...").pack(pady=(0, 10))
            progress_bar = ttk.Progressbar(frame, mode='determinate', maximum=total_formulas)
            progress_bar.pack(fill='x')
            status_label = ttk.Label(frame, text=f"0 / {total_formulas}")
            status_label.pack(pady=(5, 0))
        
        unique_full_paths = set()
        formula_idx = self.pane.view.tree_columns.index("formula")
        
        # Process in batches to avoid UI freezing
        batch_size = 25  # Process 25 formulas at a time
        processed = 0
        
        for i in range(0, total_formulas, batch_size):
            batch = self.formulas_to_summarize[i:i + batch_size]
            
            for formula_data in batch:
                if len(formula_data) > formula_idx:
                    formula_content = formula_data[formula_idx]
                    
                    # Skip empty or very short formulas
                    if not formula_content or len(str(formula_content)) < 10:
                        processed += 1
                        continue
                    
                    # Convert to string once
                    formula_str = str(formula_content)
                    
                    # Quick check: if no external link indicators, skip expensive regex
                    if "'" not in formula_str or "[" not in formula_str or "]" not in formula_str:
                        processed += 1
                        continue
                    
                    # Only run expensive regex if likely to contain external links
                    try:
                        matches = self.external_path_pattern.findall(formula_str)
                        if matches:
                            unique_full_paths.update(matches)
                    except Exception as e:
                        # Skip problematic formulas
                        print(f"Warning: Could not process formula: {e}")
                
                processed += 1
                
                # Update progress
                if show_progress:
                    progress_bar['value'] = processed
                    status_label.config(text=f"{processed} / {total_formulas}")
                    progress_window.update_idletasks()
            
            # Allow UI to update between batches
            self.update_idletasks()
        
        # Close progress window
        if show_progress:
            progress_window.destroy()
        
        self.sorted_full_paths = sorted(list(unique_full_paths))
        self.current_mode = "worksheet"

        self.btn_by_sheet = ttk.Button(self.button_frame, text="Summarize by Path\\[File]Worksheet", command=self.show_summary_by_worksheet)
        self.btn_by_sheet.pack(side='left', padx=5)
        self.btn_by_workbook = ttk.Button(self.button_frame, text="Summarize by Path\\[File] only", command=self.show_summary_by_workbook)
        self.btn_by_workbook.pack(side='left', padx=5)

        self.replace_frame = ttk.LabelFrame(self.main_frame, text="Replace Tool", padding=10)
        self.replace_frame.grid(row=2, column=0, sticky="ew")
        self.replace_frame.columnconfigure(1, weight=1)

        ttk.Label(self.replace_frame, text="Old Link (Selected):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.old_link_var = tk.StringVar(value="<No selection>")
        old_link_entry = ttk.Entry(self.replace_frame, textvariable=self.old_link_var, state="readonly")
        old_link_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=2)

        ttk.Label(self.replace_frame, text="New Link:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.new_link_entry = ttk.Entry(self.replace_frame)
        self.new_link_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        browse_button = ttk.Button(self.replace_frame, text="...", command=self.browse_for_new_link, width=4)
        browse_button.grid(row=1, column=2, sticky="w", padx=(2,5), pady=2)

        self.link_to_addresses_cache = collections.defaultdict(list)
        address_idx = self.pane.view.tree_columns.index("address")
        formula_idx = self.pane.view.tree_columns.index("formula")
        for formula_data in self.formulas_to_summarize:
            if len(formula_data) > formula_idx:
                formula_content = formula_data[formula_idx]
                matches = self.external_path_pattern.findall(str(formula_content))
                for match in matches:
                    if len(formula_data) > address_idx:
                        self.link_to_addresses_cache[match].append(formula_data[address_idx])

        self.summary_tree.bind("<<TreeviewSelect>>", self.on_link_select)

        goto_button = ttk.Button(self.replace_frame, text="Go to Excel and Select Affected Ranges", 
                                 command=lambda: select_ranges_in_excel(self, self.summary_tree, self.pane, self.link_to_addresses_cache))
        goto_button.grid(row=2, column=0, sticky="w", padx=5, pady=10)
        
        visual_button = ttk.Button(self.replace_frame, text="Show Visual Chart", command=lambda: show_visual_chart(self, self.summary_tree, self.pane, self.formulas_to_summarize))
        visual_button.grid(row=3, column=0, sticky="w", padx=5, pady=(0, 10))
        
        self.replace_button = ttk.Button(self.replace_frame, text="Perform Replacement in Excel", command=lambda: replace_links_in_excel(
            self, self.replace_frame, self.pane, self.summary_tree, self.old_link_var, self.new_link_entry, self.rescan_var,
            self.formulas_to_summarize, self.link_to_addresses_cache, self.external_path_pattern,
            self.show_summary_by_workbook, self.show_summary_by_worksheet, self.current_mode,
            self.sorted_full_paths, self.btn_by_sheet, self.btn_by_workbook, browse_button, self.replace_button
        ))
        self.replace_button.grid(row=2, column=1, columnspan=2, sticky="e", padx=5, pady=10)

        self.show_summary_by_worksheet()

        self.protocol("WM_DELETE_WINDOW", self.on_summary_close)

    def show_summary_by_worksheet(self):
        self.current_mode = "worksheet"
        self.summary_tree.delete(*self.summary_tree.get_children())
        for path in self.sorted_full_paths:
            self.summary_tree.insert("", "end", values=(path,))
        self.tree_frame.config(text=f"Found External Links (by Worksheet)")

    def show_summary_by_workbook(self):
        self.current_mode = "workbook"
        self.summary_tree.delete(*self.summary_tree.get_children())
        unique_workbook_paths = set()
        workbook_only_pattern = re.compile(r"^(.*\\\[[^\]]+\.(?:xlsx|xls|xlsm|xlsb)\])")
        for full_path in self.sorted_full_paths:
            match = workbook_only_pattern.match(full_path)
            if match:
                unique_workbook_paths.add(match.group(1))
        
        sorted_workbook_paths = sorted(list(unique_workbook_paths))
        for path in sorted_workbook_paths:
            self.summary_tree.insert("", "end", values=(path,))
            
        lf_text = "Found External Links (by Workbook)"
        if self.is_filtered:
            lf_text += " - Filtered View"
        self.tree_frame.config(text=lf_text)

    def browse_for_new_link(self):
        file_path = filedialog.askopenfilename(title="Select the new Excel file", filetypes=[("Excel Workbooks", "*.xlsx *.xls *.xlsm *.xlsb"), ("All Files", "*.*" )], parent=self)
        if not file_path:
            return
        dir_name = os.path.dirname(file_path).replace('/', '\\')
        file_name = os.path.basename(file_path)
        new_base_path = f"{dir_name}\\[{file_name}]"
        old_link = self.old_link_var.get()
        worksheet_part = ""
        if old_link != "<No selection>" and ']' in old_link:
            try:
                worksheet_part = old_link.split(']', 1)[1]
            except IndexError:
                worksheet_part = ""
        final_path = new_base_path + worksheet_part
        self.new_link_entry.delete(0, 'end')
        self.new_link_entry.insert(0, final_path)

    def on_link_select(self, event):
        selected_items = self.summary_tree.selection()
        if selected_items:
            selected_link = self.summary_tree.item(selected_items[0], "values")[0]
            self.old_link_var.set(selected_link)
            affected_addresses_for_selected = self.link_to_addresses_cache.get(selected_link, [])
            if affected_addresses_for_selected:
                selected_summary = f" - Selected Link ({smart_range_display(affected_addresses_for_selected)})";
                self.summary_tree.heading("link", text="External Link Path" + selected_summary)
            else:
                self.summary_tree.heading("link", text=self.heading_text)
        else:
            self.old_link_var.set("<No selection>")
            self.summary_tree.heading("link", text=self.heading_text)

    def on_summary_close(self):
        if hasattr(self, "did_replace") and self.did_replace and self.rescan_var.get():
            app = self.parent.winfo_toplevel().app if hasattr(self.parent.winfo_toplevel(), "app") else None
            if app:
                if self.pane.pane_name == "Worksheet1":
                    app.scan_left_quick()
                elif self.pane.pane_name == "Worksheet2":
                    app.scan_right_quick()
        self.destroy()
