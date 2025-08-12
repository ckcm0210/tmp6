import tkinter as tk
from tkinter import ttk, messagebox
import os
import re
import json
from urllib.parse import unquote

from utils.dependency_converter import convert_tree_to_graph_data
from core.graph_generator import GraphGenerator
from utils.progress_enhanced_exploder import explode_cell_dependencies_with_progress, ProgressCallback

# This will be updated to import from a new navigation_manager in a future step
from core.navigation_manager import go_to_reference_new_tab

class DependencyExploderView(tk.Toplevel):
    def __init__(self, parent, controller, workbook_path, sheet_name, cell_address, reference_display):
        super().__init__(parent)
        self.parent = parent
        self.controller = controller
        self.workbook_path = workbook_path
        self.sheet_name = sheet_name
        self.cell_address = cell_address
        self.reference_display = reference_display

        self.title(f"Dependency Explosion: {self.reference_display}")
        self.geometry("1200x800")
        self.resizable(True, True)
        # 允許最大化：移除 transient，保留 grab_set 以便模態行為
        # self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self._setup_variables()
        self._setup_ui()

        self.progress_callback.update_progress(f"Ready. Range Threshold: {self.range_threshold_var.get()}, Max Depth: {self.max_depth_var.get()}")

    def on_closing(self):
        """Handle window closing event."""
        if messagebox.askokcancel("Quit Analysis", "The analysis is not saved. Are you sure you want to close this window?"):        
            self.destroy()

    def _setup_variables(self):
        self.analysis_running = tk.BooleanVar(value=False)
        self.analysis_cancelled = tk.BooleanVar(value=False)
        self.log_panel_visible = tk.BooleanVar(value=True)
        self.show_full_address_var = tk.BooleanVar(value=False)
        self.show_full_formula_var = tk.BooleanVar(value=False)
        
        saved_range_threshold = getattr(self.controller, '_saved_range_threshold', 5)
        self.range_threshold_var = tk.IntVar(value=saved_range_threshold)
        
        saved_max_depth = getattr(self.controller, '_saved_max_depth', 8)
        self.max_depth_var = tk.IntVar(value=saved_max_depth)

        self.tree_data = None

    def _setup_ui(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        info_frame = ttk.LabelFrame(main_frame, text="Analysis Info", padding=5)
        info_frame.pack(fill='x', pady=(0, 10))

        control_frame = ttk.Frame(info_frame)
        control_frame.pack(fill='x')

        button_frame = ttk.Frame(control_frame)
        button_frame.pack(side=tk.LEFT)

        self.analyze_btn = ttk.Button(button_frame, text="Start Analysis", command=self.start_analysis)
        self.analyze_btn.pack(side=tk.LEFT, padx=5)

        self.cancel_btn = ttk.Button(button_frame, text="Cancel", state='disabled', command=self.cancel_analysis)
        self.cancel_btn.pack(side=tk.LEFT, padx=5)

        self.toggle_log_btn = ttk.Button(button_frame, text="Hide Log", command=self.toggle_log_panel)
        self.toggle_log_btn.pack(side=tk.LEFT, padx=5)

        progress_display_frame = ttk.Frame(control_frame)
        progress_display_frame.pack(side=tk.LEFT, fill='x', expand=True, padx=10)

        self.progress_var = tk.StringVar(value="Ready to analyze...")
        progress_label = ttk.Label(progress_display_frame, textvariable=self.progress_var, foreground="blue")
        progress_label.pack(side=tk.LEFT)

        progressbar_frame = ttk.Frame(control_frame)
        progressbar_frame.pack(side=tk.RIGHT)

        self.progress_bar = ttk.Progressbar(progressbar_frame, mode='indeterminate', length=200)
        self.progress_bar.pack(side=tk.RIGHT, padx=10)

        options_frame = ttk.LabelFrame(info_frame, text="Display Options", padding=5)
        options_frame.pack(fill='x', pady=(5, 0))

        options_control_frame = ttk.Frame(options_frame)
        options_control_frame.pack(fill='x')

        show_full_address_cb = ttk.Checkbutton(options_control_frame, text="Show Full Cell Address Paths", variable=self.show_full_address_var, command=self.refresh_tree_display)
        show_full_address_cb.pack(side=tk.LEFT, padx=5)

        show_full_formula_cb = ttk.Checkbutton(options_control_frame, text="Show Full Formula Paths", variable=self.show_full_formula_var, command=self.refresh_tree_display)
        show_full_formula_cb.pack(side=tk.LEFT, padx=5)

        params_frame = ttk.Frame(options_frame)
        params_frame.pack(fill='x', pady=(5, 0))

        ttk.Label(params_frame, text="Range Expansion:").pack(side=tk.LEFT, padx=5)
        ttk.Label(params_frame, text="Expand ranges with").pack(side=tk.LEFT, padx=2)
        range_threshold_spinbox = ttk.Spinbox(params_frame, from_=1, to=50, width=5, textvariable=self.range_threshold_var)
        range_threshold_spinbox.pack(side=tk.LEFT, padx=2)
        ttk.Label(params_frame, text="cells or fewer").pack(side=tk.LEFT, padx=2)

        ttk.Separator(params_frame, orient='vertical').pack(side=tk.LEFT, fill='y', padx=10)

        ttk.Label(params_frame, text="Max Depth:").pack(side=tk.LEFT, padx=5)
        max_depth_spinbox = ttk.Spinbox(params_frame, from_=1, to=20, width=5, textvariable=self.max_depth_var)
        max_depth_spinbox.pack(side=tk.LEFT, padx=2)
        ttk.Label(params_frame, text="levels deep").pack(side=tk.LEFT, padx=2)

        self.range_threshold_var.trace('w', lambda *args: self.update_params_preview())
        self.max_depth_var.trace('w', lambda *args: self.update_params_preview())

        graph_btn = ttk.Button(options_control_frame, text="Generate Graph", command=self.handle_generate_graph)
        graph_btn.pack(side=tk.RIGHT, padx=5)

        self.content_paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        self.content_paned.pack(fill='both', expand=True)

        tree_frame = ttk.LabelFrame(self.content_paned, text="Dependency Tree", padding=5)
        self.content_paned.add(tree_frame, weight=3)

        self.log_frame = ttk.LabelFrame(self.content_paned, text="Progress Log", padding=5)
        self.content_paned.add(self.log_frame, weight=1)

        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.dependency_tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set)
        self.dependency_tree.pack(fill='both', expand=True)
        tree_scroll.config(command=self.dependency_tree.yview)

        self.dependency_tree['columns'] = ('formula', 'resolved', 'value', 'type', 'depth')
        self.dependency_tree.column('#0', width=300, minwidth=200)
        self.dependency_tree.column('formula', width=350, minwidth=200)
        self.dependency_tree.column('resolved', width=350, minwidth=200)
        self.dependency_tree.column('value', width=150, minwidth=100)
        self.dependency_tree.column('type', width=100, minwidth=80)
        self.dependency_tree.column('depth', width=80, minwidth=60)

        self.dependency_tree.heading('#0', text='Cell Address', anchor=tk.W)
        self.dependency_tree.heading('formula', text='Formula', anchor=tk.W)
        self.dependency_tree.heading('resolved', text='Resolved', anchor=tk.W)
        self.dependency_tree.heading('value', text='Value', anchor=tk.W)
        self.dependency_tree.heading('type', text='Type', anchor=tk.W)
        self.dependency_tree.heading('depth', text='Depth', anchor=tk.W)

        log_scroll = ttk.Scrollbar(self.log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text = tk.Text(self.log_frame, yscrollcommand=log_scroll.set, wrap=tk.WORD, font=("Consolas", 9), bg="#f8f8f8", state='disabled')
        self.log_text.pack(fill='both', expand=True)
        log_scroll.config(command=self.log_text.yview)

        clear_log_btn = ttk.Button(self.log_frame, text="Clear Log", command=self.clear_log)
        clear_log_btn.pack(pady=2)

        summary_frame = ttk.LabelFrame(main_frame, text="Analysis Summary", padding=5)
        summary_frame.pack(fill='x', pady=(10, 0))
        self.summary_text = tk.Text(summary_frame, height=4, wrap=tk.WORD)
        self.summary_text.pack(fill='x')

        self.dependency_tree.bind("<Double-1>", self.on_tree_double_click)
        self.dependency_tree.bind("<Button-3>", self.show_context_menu)

        self.progress_callback = ProgressCallback(self.progress_var, self, self.log_text)

    def update_params_preview(self):
        range_val = self.range_threshold_var.get()
        depth_val = self.max_depth_var.get()
        self.progress_var.set(f"Ready to analyze with Range Threshold: {range_val}, Max Depth: {depth_val}. Click 'Start Analysis' to begin.")

    def handle_generate_graph(self):
        if not self.tree_data:
            messagebox.showwarning("No Data", "Please run the analysis first to generate data for the graph.", parent=self)
            return
        try:
            self.progress_var.set("Generating graph...")
            self.update()
            nodes_data, edges_data = convert_tree_to_graph_data(self.tree_data)
            if not nodes_data:
                messagebox.showinfo("Empty Graph", "The analysis result is empty, nothing to graph.", parent=self)
                return
            graph_gen = GraphGenerator(nodes_data, edges_data)
            graph_gen.generate_graph()
            self.progress_var.set("Graph generated successfully and opened in browser.")
        except Exception as e:
            messagebox.showerror("Graph Generation Error", f"Failed to generate graph: {e}", parent=self)
            self.progress_var.set(f"Graph generation failed: {e}")

    def clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def toggle_log_panel(self):
        if self.log_panel_visible.get():
            self.content_paned.remove(self.log_frame)
            self.toggle_log_btn.config(text="Show Log")
            self.log_panel_visible.set(False)
        else:
            self.content_paned.add(self.log_frame, weight=1)
            self.toggle_log_btn.config(text="Hide Log")
            self.log_panel_visible.set(True)

    def cancel_analysis(self):
        self.analysis_cancelled.set(True)
        self.cancel_btn.config(state='disabled')
        self.progress_bar.stop()
        self.progress_var.set("Analysis cancelled by user.")

    def start_analysis(self):
        try:
            self.analysis_running.set(True)
            self.analysis_cancelled.set(False)
            self.analyze_btn.config(state='disabled')
            self.cancel_btn.config(state='normal')
            self.progress_bar.start(10)
            self.progress_var.set("Initializing analysis...")
            self.update()
            for item in self.dependency_tree.get_children():
                self.dependency_tree.delete(item)
            self.clear_log()

            self.controller._saved_range_threshold = self.range_threshold_var.get()
            self.controller._saved_max_depth = self.max_depth_var.get()

            self.tree_data, summary = explode_cell_dependencies_with_progress(
                self.workbook_path, self.sheet_name, self.cell_address,
                max_depth=self.max_depth_var.get(),
                range_expand_threshold=self.range_threshold_var.get(),
                progress_callback=self.progress_callback
            )

            if self.analysis_cancelled.get():
                self.progress_var.set("Analysis was cancelled.")
                return

            self.progress_var.set("Populating tree view...")
            self.update()
            self.populate_tree(self.tree_data)
            self.show_summary(summary)
            self.progress_var.set(f"Analysis complete! Found {summary['total_nodes']} nodes, max depth: {summary['max_depth']}")
        except Exception as e:
            if "cancelled" in str(e).lower():
                self.progress_var.set("Analysis cancelled by user.")
            else:
                messagebox.showerror("Analysis Error", f"Could not analyze dependencies:\n{str(e)}", parent=self)
                self.progress_var.set(f"Analysis failed: {str(e)}")
        finally:
            self.analysis_running.set(False)
            self.analyze_btn.config(state='normal')
            self.cancel_btn.config(state='disabled')
            self.progress_bar.stop()

    def format_address_display(self, address, node):
        # 顯示正規化：將字串中的雙反斜線顯示為單反斜線（僅限UI顯示，不影響內部邏輯）
        try:
            address = address.replace('\\\\', '\\')
        except Exception:
            pass
        if not self.show_full_address_var.get():
            return node.get('short_address', address)
        return node.get('full_address', address)

    def populate_tree(self, node, parent=''):
        try:
            raw_address = node.get('address', 'Unknown')
            address = self.format_address_display(raw_address, node)

            # Get the correct formula version based on the checkbox
            if self.show_full_formula_var.get():
                formula = node.get('full_formula', node.get('formula', ''))
            else:
                formula = node.get('short_formula', node.get('formula', ''))

            # Get the correct resolved formula version
            resolved_formula = ""
            if node.get('has_resolved', False):
                if self.show_full_formula_var.get():
                    resolved_formula = node.get('full_resolved_formula', node.get('resolved_formula', ''))
                else:
                    resolved_formula = node.get('short_resolved_formula', node.get('resolved_formula', ''))

            value = str(node.get('value', ''))
            if len(value) > 20:
                value = value[:17] + "..."
            node_type = node.get('type', 'unknown')
            depth = node.get('depth', 0)
            icon = {'formula': "📊", 'value': "🔢", 'error': "❌", 'circular_ref': "🔄", 'limit_reached': "⚠️", 'range': "📋"}.get(node_type, "📄")
            item_id = self.dependency_tree.insert(parent, 'end', text=f"{icon} {address}", values=(formula, resolved_formula, value, node_type, depth))
            
            raw_formula = node.get('full_formula', node.get('formula', ''))

            node_details = {
                'workbook_path': node.get('workbook_path', self.workbook_path),
                'sheet_name': node.get('sheet_name', ''),
                'cell_address': node.get('cell_address', ''),
                'original_formula': raw_formula,
                'calculated_value': node.get('calculated_value', ''),
                'display_value': node.get('value', ''),
                'node_type': node_type,
                'depth': depth,
                'address': raw_address,
                'display_formula': formula,
                'display_address': address
            }
            try:
                details_json = json.dumps(node_details)
                self.dependency_tree.item(item_id, tags=(details_json,))
            except Exception as e:
                basic_info = f"{node_details['workbook_path']}|{node_details['sheet_name']}|{node_details['cell_address']}"
                self.dependency_tree.item(item_id, tags=(basic_info,))
            for child in node.get('children', []):
                self.populate_tree(child, item_id)
            if depth < 3:
                self.dependency_tree.item(item_id, open=True)
        except Exception as e:
            print(f"Error populating tree node: {e}")

    def show_summary(self, summary):
        self.summary_text.delete(1.0, tk.END)
        summary_content = f"""Total Nodes: {summary['total_nodes']}\nMaximum Depth: {summary['max_depth']}\nCircular References: {summary['circular_references']}\n\nNode Type Distribution:\n"""
        for node_type, count in summary['type_distribution'].items():
            summary_content += f"  {node_type}: {count}\n"
        if summary['circular_ref_list']:
            summary_content += f"\nCircular References Found:\n"
            for ref in summary['circular_ref_list']:
                summary_content += f"  {ref}\n"
        self.summary_text.insert(1.0, summary_content)

    def refresh_tree_display(self):
        try:
            expanded_items = []
            def save_expanded_state(item=''):
                children = self.dependency_tree.get_children(item)
                for child in children:
                    if self.dependency_tree.item(child, 'open'):
                        expanded_items.append(child)
                    save_expanded_state(child)
            save_expanded_state()
            if self.tree_data:
                for item in self.dependency_tree.get_children():
                    self.dependency_tree.delete(item)
                self.populate_tree(self.tree_data)
                for item_id in expanded_items:
                    try:
                        self.dependency_tree.item(item_id, open=True)
                    except: pass
        except Exception as e:
            print(f"Error refreshing tree display: {e}")

    def on_tree_double_click(self, event):
        try:
            if not self.dependency_tree.selection(): return
            item = self.dependency_tree.selection()[0]
            tags = self.dependency_tree.item(item, "tags")
            node_details = None
            if tags:
                try:
                    node_details = json.loads(tags[0])
                except (json.JSONDecodeError, ValueError):
                    parts = tags[0].split('|')
                    if len(parts) >= 3:
                        node_details = {'workbook_path': parts[0], 'sheet_name': parts[1], 'cell_address': parts[2]}
            if node_details and 'workbook_path' in node_details:
                target_workbook_path = node_details['workbook_path']
                sheet_name = node_details.get('sheet_name', '')
                cell_address = node_details.get('cell_address', '')
                address_part = self.dependency_tree.item(item, "text").split(" ", 1)[1]
                if not os.path.exists(target_workbook_path):
                    filename = os.path.basename(target_workbook_path)
                    base_dir = os.path.dirname(self.workbook_path)
                    alt_path = os.path.join(base_dir, filename)
                    if os.path.exists(alt_path):
                        target_workbook_path = alt_path
                go_to_reference_new_tab(self.controller, target_workbook_path, sheet_name, cell_address, address_part)
        except Exception as e:
            messagebox.showerror("Navigation Error", f"Could not navigate:\n{str(e)}", parent=self)

    def show_context_menu(self, event):
        item = self.dependency_tree.identify_row(event.y)
        if item:
            self.dependency_tree.selection_set(item)
            context_menu = tk.Menu(self, tearoff=0)
            context_menu.add_command(label="Go to Reference", command=lambda: self.on_tree_double_click(None))
            context_menu.add_command(label="Copy Address", command=lambda: self.copy_address(item))
            context_menu.add_separator()
            context_menu.add_command(label="Expand All", command=lambda: self.expand_all(item))
            context_menu.add_command(label="Collapse All", command=lambda: self.collapse_all(item))
            context_menu.post(event.x_root, event.y_root)

    def copy_address(self, item):
        item_text = self.dependency_tree.item(item, "text")
        address_part = item_text.split(" ", 1)[1] if " " in item_text else item_text
        self.clipboard_clear()
        self.clipboard_append(address_part)

    def expand_all(self, item):
        self.dependency_tree.item(item, open=True)
        for child in self.dependency_tree.get_children(item):
            self.expand_all(child)

    def collapse_all(self, item):
        self.dependency_tree.item(item, open=False)
        for child in self.dependency_tree.get_children(item):
            self.collapse_all(child)