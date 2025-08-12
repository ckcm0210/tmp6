import os
import time
import win32com.client
from tkinter import messagebox
import psutil
import win32gui
import win32process
import win32con
from core.formula_classifier import classify_formula_type
from core.worksheet_tree import apply_filter
import traceback

def _get_formulas_from_excel(worksheet_com_obj, scan_range_com_obj, scan_mode, progress_update_callback):
    all_formulas_local = []
    formula_cells_found = 0
    
    try:
        xlCellTypeFormulas = -4123
        formula_range = scan_range_com_obj.SpecialCells(xlCellTypeFormulas)
        
        areas_to_process = []
        if formula_range.Areas.Count > 1:
            for area in formula_range.Areas:
                areas_to_process.append(area)
        else:
            areas_to_process.append(formula_range)
            
        total_cells_to_process = sum(area.Cells.Count for area in areas_to_process)
        current_cell_count = 0
        
        for area in areas_to_process:
            for cell in area.Cells:
                current_cell_count += 1
                formula_cells_found += 1
                
                formula = ""
                formula_type = "unknown"
                cell_value = None
                display_val = "Error"
                cell_text = "Error"
                cell_address = ""
                
                try:
                    formula = cell.Formula
                    formula_type = classify_formula_type(formula)
                    cell_value = cell.Value
                    display_val = str(cell_value)[:50] if cell_value is not None else "No Value"
                    if scan_mode == 'quick':
                        cell_text = "N/A (Quick Scan)"
                    else:
                        cell_text = str(cell.Text).strip()
                    cell_address = cell.Address.replace('$', '')
                    all_formulas_local.append((formula_type, cell_address, formula, display_val, cell_text))
                except Exception as cell_processing_e:
                    all_formulas_local.append((formula_type, cell_address if cell_address else "ERROR_ADDR", str(formula), str(display_val), f"ERROR: {cell_processing_e}"))
                
                if current_cell_count % 100 == 0 or current_cell_count == total_cells_to_process:
                    progress_update_callback(current_cell_count, total_cells_to_process, formula_cells_found)
                    
    except Exception as e:
        no_formula_error = (
            "(-2146827284, 'OLE error.', None, None)" in str(e)
            or "0x800A03EC" in str(e)
            or '找不到所要找的儲存格' in str(e)
            or 'Unable to get the' in str(e)
        )
        if no_formula_error:
            return [], 0, 0 # Return empty list if no formulas found
        else:
            raise # Re-raise other exceptions

    return all_formulas_local, formula_cells_found, total_cells_to_process

def refresh_data(controller, btn, scan_mode='full'):
    if not controller.view.ui_initialized:
        return

    controller.clear_filter_inputs()
    
    if btn is not None:
        btn.config(state='disabled')
    controller.view.progress_bar['value'] = 0
    controller.view.progress_label.config(text="Connecting to active Excel...")
    controller.root.update_idletasks()

    try:
        try:
            controller.xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception as e:
            try:
                controller.xl = win32com.client.Dispatch("Excel.Application")
                controller.xl.Visible = True
                messagebox.showinfo("Info", "No existing Excel instance detected. A new Excel instance has been started automatically. Please open your file manually and press Scan again.")
                if btn is not None:
                    btn.config(state='normal')
                controller.view.progress_bar['value'] = 0
                controller.view.progress_label.config(text="Please open your Excel file and try again.")
                return
            except Exception as e2:
                messagebox.showerror("Connection Error", f"Could not find an existing Excel instance or start a new one.\nPlease check for any leftover EXCEL.EXE processes and verify your permission.\n\nError: {e2}")
                if btn is not None:
                    btn.config(state='normal')
                controller.view.progress_bar['value'] = 0
                controller.view.progress_label.config(text="Connection Failed.")
                return
        try:
            controller.workbook = controller.xl.ActiveWorkbook
            controller.worksheet = controller.xl.ActiveSheet
        except Exception as e:
            messagebox.showerror("Connection Error", "Excel is open, but there is no active workbook or worksheet.\nPlease open a file and worksheet in Excel and try again.")
            if btn is not None:
                btn.config(state='normal')
            controller.view.progress_bar['value'] = 0
            controller.view.progress_label.config(text="No active workbook.")
            return

        controller.last_workbook_path = controller.workbook.FullName
        controller.last_worksheet_name = controller.worksheet.Name
        controller.view.progress_bar['value'] = 10
        controller.view.progress_label.config(text="Reading workbook information...")
        controller.root.update_idletasks()
        file_path = controller.workbook.FullName
        display_path = os.path.dirname(file_path)
        max_path_display_length = 60
        if len(display_path) > max_path_display_length:
            truncated_path = "..." + display_path[-(max_path_display_length-3):]
        else:
            truncated_path = display_path
        controller.view.file_label.config(text=os.path.basename(file_path), foreground="black")
        controller.view.path_label.config(text=truncated_path, foreground="black")
        controller.view.sheet_label.config(text=controller.worksheet.Name, foreground="black")
        current_scan_range = controller.worksheet.UsedRange
        current_scan_range_str = current_scan_range.Address.replace('$', '')
        controller.view.range_label.config(text=f"Scanning: UsedRange ({current_scan_range_str})", foreground="black")
        controller.all_formulas.clear()
        
        controller.view.progress_bar['value'] = 30
        is_selected_range_scan = getattr(controller, 'scanning_selected_range', False)
        selected_address = getattr(controller, 'selected_scan_address', None)
        
        if is_selected_range_scan and selected_address:
            try:
                scan_range = controller.worksheet.Range(selected_address)
                scan_info = f"Selected Range: {selected_address}"
            except Exception as e:
                scan_range = current_scan_range
                scan_info = "Full Worksheet"
                is_selected_range_scan = False
        else:
            scan_range = current_scan_range
            scan_info = "Full Worksheet"
        
        scan_range_str = scan_range.Address.replace('$', '')
        controller.view.range_label.config(text=f"Scanning: {scan_info} ({scan_range_str})", foreground="black")
        
        controller.view.progress_label.config(text=f"Searching for formulas in Excel in {scan_info.lower()} {scan_range_str} (this may take a moment)...")
        controller.root.update_idletasks()
        
        start_time = time.time()
        try:
            def progress_callback(current_cell_count, total_cells_to_process, formula_cells_found):
                progress = 30 + (current_cell_count / total_cells_to_process) * 60
                controller.view.progress_bar['value'] = min(int(progress), 90)
                controller.view.progress_label.config(text=f"Found {formula_cells_found} formulas. Processing {current_cell_count}/{total_cells_to_process} cells...")
                controller.root.update_idletasks()

            controller.all_formulas, formula_cells_found, total_cells_to_process = _get_formulas_from_excel(
                controller.worksheet, scan_range, scan_mode, progress_callback
            )
            
        except Exception as e:
            import traceback
            err_detail = traceback.format_exc()
            no_formula_error = (
                "(-2146827284, 'OLE error.', None, None)" in str(e)
                or "0x800A03EC" in str(e)
                or '找不到所要找的儲存格' in str(e)
                or 'Unable to get the' in str(e)
            )
            if no_formula_error:
                controller.view.progress_label.config(text=f"No formulas found in this worksheet's {scan_info.lower()} ({scan_range_str}).")
                if hasattr(controller, 'scanning_selected_range'):
                    controller.scanning_selected_range = False
                if hasattr(controller, 'selected_scan_address'):
                    controller.selected_scan_address = None
                if hasattr(controller, 'selected_scan_count'):
                    controller.selected_scan_count = None
                    
                controller.view.progress_bar['value'] = 100
                controller.root.update_idletasks()
                if controller.view.formula_list_label:
                    controller.view.formula_list_label.config(text="Formula List (No Formula Found)")
                apply_filter(controller)
                if btn is not None:
                    btn.config(state='normal')
                return
            else:
                messagebox.showerror("Scan Error", f"An error occurred while scanning formulas: {e}\n\nTraceback:\n{err_detail}")
            if btn is not None:
                btn.config(state='normal')
            return
        
        end_time = time.time()
        time_taken = end_time - start_time
        controller.view.progress_bar['value'] = 90
        controller.view.progress_label.config(text=f"Found {len(controller.all_formulas)} formulas. Loading... (Scan took {time_taken:.2f} seconds)")
        controller.root.update_idletasks()

        if hasattr(controller, 'original_user_selection') and controller.original_user_selection:
            controller._filter_results_to_original_selection()
        
        if hasattr(controller, 'scanning_selected_range'):
            controller.scanning_selected_range = False
        if hasattr(controller, 'selected_scan_address'):
            controller.selected_scan_address = None
        if hasattr(controller, 'selected_scan_count'):
            controller.selected_scan_count = None
        if hasattr(controller, 'original_user_selection'):
            controller.original_user_selection = None
        if hasattr(controller, 'original_user_count'):
            controller.original_user_count = None
            
        apply_filter(controller)
        controller.view.progress_bar['value'] = 100
        controller.view.progress_label.config(text=f"Completed: Found {len(controller.all_formulas)} formulas. (Total scan time: {time_taken:.2f} seconds)")
        if btn is not None:
            btn.config(state='normal')
        if controller.view.formula_list_label:
            total_count = len(controller.all_formulas)
            if total_count == 0:
                controller.view.formula_list_label.config(text="Formula List (No Formula Found)")
            else:
                controller.view.formula_list_label.config(text=f"Formula List ({total_count} records):")
    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        messagebox.showerror(
            "Connection Error",
            "Could not connect to Excel. Please ensure:\n"
            "1. Excel is open\n"
            "2. There are no leftover EXCEL.EXE processes (check with Task Manager)\n"
            "3. Permissions are consistent (do not run one as admin and the other not)\n\n"
            f"Error: {e}\n\nTraceback:\n{err_detail}"
        )
        if btn is not None:
            btn.config(state='normal')
        controller.view.progress_bar['value'] = 0
        controller.view.progress_label.config(text="Connection Failed.")
        return