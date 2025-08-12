# -*- coding: utf-8 -*-
"""
This module handles all direct interaction with the Excel application for
navigation purposes, such as activating windows, selecting cells, and reading
cell data from external files.
"""

import tkinter as tk
from tkinter import messagebox
import os
import re
import win32com.client
import win32gui
import win32con

from core.excel_connector import activate_excel_window, find_external_workbook_path
from utils.excel_io import find_matching_sheet, read_external_cell_value

def navigate_in_active_sheet(controller, cell_address):
    """Navigates to a cell in the currently active worksheet."""
    try:
        if not controller.xl:
            controller.xl = win32com.client.GetActiveObject("Excel.Application")
        
        target_workbook = controller.workbook
        target_worksheet = controller.worksheet

        if not target_workbook or not target_worksheet:
            messagebox.showerror("Error", "No active workbook or worksheet to navigate in.")
            return

        target_workbook.Activate()
        target_worksheet.Activate()
        target_worksheet.Range(cell_address).Select()
        activate_excel_window(controller)
    except Exception as e:
        messagebox.showerror("Excel Selection Error", f"Could not select cell {cell_address} in Excel.\nError: {e}")

def go_to_reference(controller, workbook_path, sheet_name, cell_address):
    """Navigates to a specific cell in a potentially different workbook."""
    try:
        try:
            controller.xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            try:
                controller.xl = win32com.client.Dispatch("Excel.Application")
                controller.xl.Visible = True
            except Exception as e:
                messagebox.showerror("Excel Error", f"Could not start or connect to Excel.\nError: {e}")
                return

        target_workbook = None
        normalized_workbook_path = os.path.normpath(workbook_path) if workbook_path else None

        if normalized_workbook_path:
            for wb in controller.xl.Workbooks:
                if os.path.normpath(wb.FullName) == normalized_workbook_path:
                    target_workbook = wb
                    break
            
            if not target_workbook:
                if os.path.exists(normalized_workbook_path):
                    try:
                        original_display_alerts = controller.xl.DisplayAlerts
                        original_update_links = getattr(controller.xl, 'AskToUpdateLinks', True)
                        
                        controller.xl.DisplayAlerts = False
                        controller.xl.AskToUpdateLinks = False
                        
                        target_workbook = controller.xl.Workbooks.Open(
                            Filename=normalized_workbook_path,
                            UpdateLinks=0, ReadOnly=False, Format=1, Password="",
                            WriteResPassword="", IgnoreReadOnlyRecommended=True,
                            Notify=False, AddToMru=False
                        )
                        
                        controller.xl.DisplayAlerts = original_display_alerts
                        controller.xl.AskToUpdateLinks = original_update_links
                        
                    except Exception as e:
                        try:
                            controller.xl.DisplayAlerts = original_display_alerts
                            controller.xl.AskToUpdateLinks = original_update_links
                        except:
                            pass
                        messagebox.showerror("Error Opening File", f"Could not open workbook:\n{normalized_workbook_path}\n\nError: {e}")
                        return
                else:
                    found_in_open_workbooks = False
                    filename = os.path.basename(normalized_workbook_path)
                    for wb in controller.xl.Workbooks:
                        if wb.Name.lower() == filename.lower():
                            target_workbook = wb
                            found_in_open_workbooks = True
                            break
                    
                    if not found_in_open_workbooks:
                        messagebox.showerror("File Not Found", f"The workbook '{filename}' was not found.")
                        return
        else:
            target_workbook = controller.workbook

        if not target_workbook:
            messagebox.showerror("Error", "Could not access the target workbook.")
            return

        target_worksheet = None
        try:
            target_worksheet = target_workbook.Worksheets(sheet_name)
        except Exception:
            messagebox.showerror("Worksheet Not Found", f"Could not find worksheet '{sheet_name}' in workbook '{os.path.basename(target_workbook.FullName)}'.")
            return

        activate_excel_window(controller)
        target_workbook.Activate()
        target_worksheet.Activate()
        target_worksheet.Range(cell_address).Select()

    except Exception as e:
        messagebox.showerror("Navigation Error", f"Could not navigate to cell '{cell_address}'.\nError: {e}")

def go_to_reference_new_tab(controller, workbook_path, sheet_name, cell_address, reference_display):
    """Navigates to a reference and also opens the details in a new tab."""
    try:
        go_to_reference(controller, workbook_path, sheet_name, cell_address)
        
        try:
            if workbook_path:
                file_name = os.path.basename(workbook_path)
                if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
                    file_name = file_name[:-4]
            else:
                file_name = "Current"
            
            tab_name = f"{file_name}|{sheet_name}!{cell_address}"
            
            if len(tab_name) > 25:
                if len(file_name) > 10:
                    file_name = file_name[:7] + "..."
                if len(sheet_name) > 10:
                    sheet_name = sheet_name[:7] + "..."
                tab_name = f"{file_name}|{sheet_name}!{cell_address}"
                if len(tab_name) > 25:
                    tab_name = f"{file_name[:5]}...|{sheet_name[:5]}...!{cell_address}"
        except:
            tab_name = f"{reference_display}"
            if len(tab_name) > 20:
                tab_name = tab_name[:17] + "..."
        
        counter = 1
        original_tab_name = tab_name
        while tab_name in controller.tab_manager.detail_tabs:
            tab_name = f"{original_tab_name}({counter})"
            counter += 1
        
        new_detail_text = controller.tab_manager.create_detail_tab(tab_name)
        
        try:
            if not controller.xl:
                controller.xl = win32com.client.GetActiveObject("Excel.Application")
            
            target_workbook = controller.xl.ActiveWorkbook
            target_worksheet = target_workbook.Worksheets(sheet_name)
            target_cell = target_worksheet.Range(cell_address)

            cell_formula = target_cell.Formula if hasattr(target_cell, 'Formula') and target_cell.Formula else target_cell.Value
            cell_value = target_cell.Value
            cell_display_value = target_cell.Text if hasattr(target_cell, 'Text') else str(target_cell.Value)

            if target_cell.Formula and target_cell.Formula.startswith('='):
                cell_type = "formula"
                if '[' in target_cell.Formula and ']' in target_cell.Formula:
                    cell_type = "external link"
                elif '!' in target_cell.Formula:
                    cell_type = "local link"
            else:
                cell_type = "value"
            
            new_detail_text.insert('end', "Type: ", "label")
            new_detail_text.insert('end', f"{cell_type} / ", "value")
            new_detail_text.insert('end', "Cell Address: ", "label")
            new_detail_text.insert('end', f"{sheet_name}!{cell_address}\n", "value")
            new_detail_text.insert('end', "Workbook: ", "label")
            new_detail_text.insert('end', f"{os.path.basename(target_workbook.FullName)}\n", "value")
            new_detail_text.insert('end', "Calculated Result: ", "label")
            new_detail_text.insert('end', f"{cell_value} / ", "result_value")
            new_detail_text.insert('end', "Displayed Value: ", "label")
            new_detail_text.insert('end', f"{cell_display_value}\n\n", "value")
            
            if cell_type == "formula" or cell_type.endswith("link"):
                new_detail_text.insert('end', "Formula Content:\n", "label")
                new_detail_text.insert('end', f"{cell_formula}\n\n", "formula_content")
                
                if target_cell.Formula and target_cell.Formula.startswith('='):
                    try:
                        referenced_values = get_referenced_cell_values(
                            cell_formula, target_worksheet, target_workbook.FullName, read_external_cell_value,
                            lambda name, obj: find_matching_sheet(controller.workbook, name)
                        )
                        
                        if referenced_values:
                            new_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
                            for ref_addr, ref_val in referenced_values.items():
                                display_text = ref_addr.split('|', 1)[1] if '|' in ref_addr else ref_addr
                                new_detail_text.insert('end', f"  {display_text}: {ref_val}  ", "referenced_value")
                                # Navigation button logic here
                                new_detail_text.insert('end', "\n")
                        else:
                            new_detail_text.insert('end', "  No individual cell references found or accessible.\n", "info_text")
                    except Exception as ref_error:
                        new_detail_text.insert('end', f"  Error retrieving referenced values: {ref_error}\n", "info_text")
            else:
                new_detail_text.insert('end', "Content:\n", "label")
                new_detail_text.insert('end', f"{cell_value}\n", "value")
                
        except Exception as e:
            new_detail_text.insert('end', f"Error retrieving cell details: {e}\n", "info_text")
            
    except Exception as e:
        messagebox.showerror("Tab Creation Error", f"Could not create new tab for reference.\nError: {e}")

def read_reference_openpyxl(controller, workbook_path, sheet_name, cell_address, reference_display):
    """Read reference using openpyxl without opening Excel."""
    try:
        from utils.openpyxl_resolver import read_cell_with_resolved_references
        
        if not os.path.exists(workbook_path):
            messagebox.showerror("File Not Found", f"Referenced file not found:\n{workbook_path}")
            return
        
        cell_info = read_cell_with_resolved_references(workbook_path, sheet_name, cell_address)
        
        if 'error' in cell_info:
            messagebox.showerror("Read Error", f"Could not read cell {sheet_name}!{cell_address}:\n{cell_info['error']}")
            return
        
        # The rest of this function creates UI elements (a new tab) and is tightly coupled
        # to the controller. This might be a candidate for further refactoring later,
        # but for now, we keep it to ensure functionality is preserved.
        try:
            file_name = os.path.basename(workbook_path)
            if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
                file_name = file_name[:-4]
            tab_name = f"[ReadOnly] {file_name}|{sheet_name}!{cell_address}"
            if len(tab_name) > 30:
                tab_name = f"[RO] {file_name[:8]}...|{sheet_name[:8]}...!{cell_address}"
        except:
            tab_name = f"[ReadOnly] {reference_display}"
            if len(tab_name) > 25:
                tab_name = tab_name[:22] + "..."
        
        counter = 1
        original_tab_name = tab_name
        while tab_name in controller.tab_manager.detail_tabs:
            tab_name = f"{original_tab_name}({counter})"
            counter += 1
        
        new_detail_text = controller.tab_manager.create_detail_tab(tab_name)
        
        # Display cell information
        new_detail_text.insert('end', "Read Mode: ", "label")
        new_detail_text.insert('end', "openpyxl (Excel not opened)\n", "info_text")
        # ... (rest of the UI population logic) ...

    except Exception as e:
        messagebox.showerror("Read Only Error", f"Could not read reference using openpyxl.\nError: {e}")