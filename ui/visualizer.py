import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import messagebox
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import os
from utils.range_optimizer import parse_cell_address, format_range # Assuming these are in the optimizer

class ChartVisualizer:
    def __init__(self, parent, pane, formulas_to_summarize, selected_link):
        self.parent = parent
        self.pane = pane
        self.formulas_to_summarize = formulas_to_summarize
        self.selected_link = selected_link
        self.affected_addresses = self._get_affected_addresses()

        if not self.affected_addresses:
            messagebox.showinfo(
                "No Affected Cells", 
                f"No cells found that are affected by the selected external link:\n{self.selected_link}",
                parent=self.parent
            )
            return

        self.create_chart_window()

    def _get_affected_addresses(self):
        affected = []
        try:
            address_idx = self.pane.view.tree_columns.index("address")
            formula_idx = self.pane.view.tree_columns.index("formula")
        except ValueError:
            address_idx, formula_idx = 1, 2 # Fallback

        for formula_data in self.formulas_to_summarize:
            if len(formula_data) > formula_idx and self.selected_link in str(formula_data[formula_idx]):
                if len(formula_data) > address_idx:
                    affected.append(formula_data[address_idx])
        return affected

    def create_chart_window(self):
        self.chart_window = tk.Toplevel(self.parent)
        self.chart_window.title("Visual Chart - Affected Ranges")
        self.chart_window.geometry("1000x700")
        self.chart_window.transient(self.parent)

        control_frame = ttk.Frame(self.chart_window)
        control_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(control_frame, text="View Mode: Full Used Range").pack(side=tk.LEFT, padx=(0, 5))

        summary_frame = ttk.LabelFrame(self.chart_window, text="Summary Information", padding=10)
        summary_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        self.summary_labels = {
            'worksheet': ttk.Label(summary_frame, text="Current File: "),
            'external_link': ttk.Label(summary_frame, text="Selected External Link: ", wraplength=800)
        }
        self.summary_labels['worksheet'].grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        self.summary_labels['external_link'].grid(row=1, column=0, columnspan=2, sticky=tk.W)

        self.chart_frame = ttk.Frame(self.chart_window)
        self.chart_frame.pack(fill=tk.BOTH, expand=True)

        self.create_chart()

        # Add bottom button frame
        bottom_button_frame = ttk.Frame(self.chart_window)
        bottom_button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        export_button = ttk.Button(bottom_button_frame, text="Export Chart", command=self.export_chart)
        export_button.pack(side=tk.RIGHT, padx=(0, 5))
        
        close_button = ttk.Button(bottom_button_frame, text="Close", command=self.chart_window.destroy)
        close_button.pack(side=tk.RIGHT)

    def create_chart(self):
        for widget in self.chart_frame.winfo_children():
            widget.destroy()

        parsed_coords = [coord for coord in (parse_cell_address(addr) for addr in self.affected_addresses) if coord]
        if not parsed_coords:
            messagebox.showerror("Error", "Could not parse cell addresses for visualization.", parent=self.chart_window)
            return

        try:
            used_range = self.pane.worksheet.UsedRange
            display_min_row, display_max_row = used_range.Row, used_range.Row + used_range.Rows.Count - 1
            display_min_col, display_max_col = used_range.Column, used_range.Column + used_range.Columns.Count - 1
            view_title = "Worksheet Overview - Full Used Range"
        except Exception:
            display_min_col, display_max_col = 1, 10
            display_min_row, display_max_row = 1, 20
            view_title = "Worksheet Overview - Default Range"

        fig, ax = plt.subplots(figsize=(12, 8))
        fig.patch.set_facecolor('white')
        plt.subplots_adjust(top=0.80, bottom=0.15, left=0.08, right=0.95)

        col_range = display_max_col - display_min_col + 1
        row_range = display_max_row - display_min_row + 1

        # Grid background
        col_step = max(1, col_range // 20) if col_range > 50 else 1
        row_step = max(1, row_range // 20) if row_range > 50 else 1
        ax.set_xticks([i - 0.5 for i in range(display_min_col, display_max_col + 2, col_step)], minor=True)
        ax.set_yticks([i - 0.5 for i in range(display_min_row, display_max_row + 2, row_step)], minor=True)
        ax.grid(which='minor', color='lightgray', linestyle='-', linewidth=0.5)

        # Used range background
        used_rect = patches.Rectangle((display_min_col - 0.5, display_min_row - 0.5), col_range, row_range, linewidth=2, edgecolor='blue', facecolor='lightblue', alpha=0.1)
        ax.add_patch(used_rect)

        # Highlight affected cells
        for col, row in parsed_coords:
            rect = patches.Rectangle((col - 0.5, row - 0.5), 1, 1, linewidth=1, edgecolor='red', facecolor='lightcoral', alpha=0.8)
            ax.add_patch(rect)

        ax.set_xlim(display_min_col - 0.5, display_max_col + 0.5)
        ax.set_ylim(display_max_row + 0.5, display_min_row - 0.5) # Inverted Y-axis

        # Labels
        def col_num_to_letter(n):
            string = ""
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(65 + remainder) + string
            return string

        col_ticks = list(range(display_min_col, display_max_col + 1, col_step))
        ax.set_xticks(col_ticks)
        ax.set_xticklabels([col_num_to_letter(c) for c in col_ticks])
        ax.xaxis.set_label_position('top')
        ax.xaxis.tick_top()

        row_ticks = list(range(display_min_row, display_max_row + 1, row_step))
        ax.set_yticks(row_ticks)
        ax.set_yticklabels(row_ticks)

        ax.set_title(f'{view_title}\n{len(self.affected_addresses)} cells affected by selected external link', fontsize=12, fontweight='bold', pad=30)
        ax.set_xlabel('Columns', fontsize=10)
        ax.set_ylabel('Rows', fontsize=10)

        legend_elements = [
            patches.Patch(facecolor='lightcoral', edgecolor='red', label='Affected Cells'),
            patches.Patch(facecolor='lightblue', edgecolor='blue', alpha=0.3, label='Used Range')
        ]
        ax.legend(handles=legend_elements, loc='lower center', bbox_to_anchor=(0.5, -0.18), ncol=2)

        self.update_summary_labels()

        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        toolbar_frame = ttk.Frame(self.chart_frame)
        toolbar_frame.pack(fill=tk.X)
        self.toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
        self.toolbar.update()
        self.fig = fig

    def update_summary_labels(self):
        try:
            workbook_path = self.pane.workbook.FullName
            workbook_name = os.path.basename(workbook_path)
            worksheet_name = self.pane.worksheet.Name
            full_path_info = f"{os.path.dirname(workbook_path)}\\n[{workbook_name}]{worksheet_name}"
        except Exception:
            full_path_info = "Unknown Workbook/Worksheet"
        
        self.summary_labels['worksheet'].config(text=f"Current File: {full_path_info}")
        self.summary_labels['external_link'].config(text=f"Selected External Link: {self.selected_link} ({len(self.affected_addresses)} cells affected)")

    def export_chart(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Chart",
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), ("SVG files", "*.svg")],
            parent=self.chart_window
        )
        if file_path and hasattr(self, 'fig'):
            try:
                self.fig.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Export Complete", f"Chart saved to:\n{file_path}", parent=self.chart_window)
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to save chart:\n{str(e)}", parent=self.chart_window)

def show_visual_chart(summary_window, summary_tree, pane, formulas_to_summarize):
    """
    Handles the logic for showing the visual chart.
    This function is called when the 'Show Visual Chart' button is clicked.
    """
    selected_items = summary_tree.selection()
    if not selected_items:
        messagebox.showwarning("No Selection", "Please select an external link from the list first.", parent=summary_window)
        return
    
    # It's possible the treeview has multiple columns in the future, but for now, the link is in the first value.
    selected_link = summary_tree.item(selected_items[0], "values")[0]
    
    # Launch the visualizer
    ChartVisualizer(summary_window, pane, formulas_to_summarize, selected_link)
