import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from ui.summary_window import SummaryWindow
from core.data_processor import _get_summary_data



def summarize_external_links(controller):
    if not controller.view.result_tree.get_children():
        messagebox.showinfo("No Data", "There are no formulas in the list to summarize.\nPlease scan a worksheet first, or adjust filters.")
        return

    formulas_to_summarize, is_filtered = _get_summary_data(controller)

    SummaryWindow(controller.root, controller, formulas_to_summarize, is_filtered)

