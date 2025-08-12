import os
import win32com.client
from tkinter import messagebox
import win32gui
import win32con

def _perform_excel_reconnection(controller, last_workbook_path, last_worksheet_name):
    xl = None
    workbook = None
    worksheet = None

    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        xl.Visible = True
    except Exception:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True

    found_workbook = None
    for wb in xl.Workbooks:
        if wb.FullName == last_workbook_path:
            found_workbook = wb
            break
    
    if found_workbook:
        workbook = found_workbook
    else:
        if os.path.exists(last_workbook_path):
            workbook = xl.Workbooks.Open(last_workbook_path)
        else:
            raise FileNotFoundError(f"Saved path file does not exist: {last_workbook_path}")
    
    worksheet = workbook.Worksheets(last_worksheet_name)
    workbook.Activate()
    worksheet.Activate()
    
    return xl, workbook, worksheet

def reconnect_to_excel(controller):
    if not controller.last_workbook_path or not controller.last_worksheet_name:
        messagebox.showerror("Cannot Reconnect", "No saved workbook path or worksheet name.\nPlease scan a worksheet first.")
        return

    try:
        controller.xl, controller.workbook, controller.worksheet = _perform_excel_reconnection(controller, controller.last_workbook_path, controller.last_worksheet_name)
        activate_excel_window(controller)

        file_path = controller.workbook.FullName
        display_path = os.path.dirname(file_path)
        max_path_display_length = 60
        if len(display_path) > max_path_display_length:
            truncated_path = "..." + display_path[-(max_path_display_length - 3):]
        else:
            truncated_path = display_path
        controller.view.file_label.config(text=os.path.basename(file_path), foreground="black")
        controller.view.path_label.config(text=truncated_path, foreground="black")
        controller.view.sheet_label.config(text=controller.worksheet.Name, foreground="black")
        current_scan_range_str = controller.worksheet.UsedRange.Address.replace('$', '')
        controller.view.range_label.config(text=f"UsedRange ({current_scan_range_str})", foreground="black")
        messagebox.showinfo("Connection Successful", f"Successfully reconnected to:\n{controller.workbook.Name} - {controller.worksheet.Name}")

    except FileNotFoundError as e:
        messagebox.showerror("File Not Found", str(e))
        controller.view.file_label.config(text="Not Connected", foreground="red")
        controller.view.path_label.config(text="Not Connected", foreground="red")
        controller.view.sheet_label.config(text="Not Connected", foreground="red")
        controller.view.range_label.config(text="Not Connected", foreground="red")
    except Exception as e:
        messagebox.showerror("Connection Error", f"Could not connect to Excel or activate worksheet.\nError: {e}")
        controller.view.file_label.config(text="Not Connected", foreground="red")
        controller.view.path_label.config(text="Not Connected", foreground="red")
        controller.view.sheet_label.config(text="Not Connected", foreground="red")
        controller.view.range_label.config(text="Not Connected", foreground="red")

def activate_excel_window(controller):
    if not controller.xl:
        return
    try:
        controller.xl.Visible = True
        excel_hwnd = controller.xl.Hwnd
        if win32gui.IsIconic(excel_hwnd):
            win32gui.ShowWindow(excel_hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(excel_hwnd)
        controller.xl.ActiveWindow.Activate()
    except Exception as e:
        messagebox.showwarning("Activate Excel Window", f"Could not activate Excel window. Please switch to Excel manually. Error: {e}")

def find_external_workbook_path(controller, file_name):
    """Find the path of an external workbook by checking open workbooks first"""
    try:
        if not controller.xl:
            controller.xl = win32com.client.GetActiveObject("Excel.Application")
        
        for wb in controller.xl.Workbooks:
            if wb.Name.lower() == file_name.lower():
                return wb.FullName
        
        if controller.workbook and hasattr(controller.workbook, 'Path'):
            current_dir = controller.workbook.Path
            potential_path = os.path.join(current_dir, file_name)
            if os.path.exists(potential_path):
                return potential_path
        
        return file_name
        
    except Exception:
        return file_name
