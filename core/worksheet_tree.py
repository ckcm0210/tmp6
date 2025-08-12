# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 09:11:26 2025

@author: kccheng
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import re
import win32com.client
import win32gui
import win32con
from functools import partial

from utils.dependency_converter import convert_tree_to_graph_data
from core.graph_generator import GraphGenerator

# Import functions from their new locations
from core.link_analyzer import get_referenced_cell_values
from utils.excel_io import find_matching_sheet, read_external_cell_value
from utils.range_optimizer import parse_excel_address
from core.excel_connector import activate_excel_window, find_external_workbook_path
from openpyxl.utils import get_column_letter, column_index_from_string


_last_range_threshold = 5
_last_max_depth = 10

def apply_filter(controller, event=None):
    controller.view.result_tree.delete(*controller.view.result_tree.get_children())
    controller.cell_addresses.clear()
    address_filter_str = controller.view.filter_entries['address'].get().strip()
    parsed_address_filters = []
    if address_filter_str and address_filter_str != controller.placeholder_text:
        address_tokens = [token.strip() for token in address_filter_str.split(',') if token.strip()]
        if address_tokens:
            try:
                for token in address_tokens:
                    parsed_address_filters.append(parse_excel_address(token))
            except Exception as e:
                messagebox.showerror("Invalid Excel Address", str(e))
                return
    other_filters = {
        'type': (controller.show_formula.get(), controller.show_local_link.get(), controller.show_external_link.get()),
        'formula': controller.view.filter_entries['formula'].get().lower(),
        'result': controller.view.filter_entries['result'].get().lower(),
        'display_value': controller.view.filter_entries['display_value'].get().lower()
    }
    filtered_formulas = []
    for formula_data in controller.all_formulas:
        if len(formula_data) < 5: continue
        formula_type, address, formula_content, result_val, display_val = formula_data
        type_map = {'formula': other_filters['type'][0], 'local link': other_filters['type'][1], 'external link': other_filters['type'][2]}
        if not type_map.get(formula_type, True): continue
        if other_filters['formula'] and other_filters['formula'] not in str(formula_content).lower(): continue
        if other_filters['result'] and other_filters['result'] not in str(result_val).lower(): continue
        if other_filters['display_value'] and other_filters['display_value'] not in str(display_val).lower(): continue
        if parsed_address_filters:
            addr_upper = address.replace("$", "").upper()
            current_cell_match = re.match(r"([A-Z]+)([0-9]+)", addr_upper)
            if not current_cell_match: continue
            cell_col_str, cell_row_str = current_cell_match.groups()
            cell_col_idx = column_index_from_string(cell_col_str)
            cell_row_idx = int(cell_row_str)
            is_match = False
            for f_type, f_val in parsed_address_filters:
                if f_type == 'cell' and addr_upper == f_val:
                    is_match = True; break
                elif f_type == 'row_range':
                    start_r, end_r = map(int, f_val.split(':'))
                    if start_r <= cell_row_idx <= end_r:
                        is_match = True; break
                elif f_type == 'col_range':
                    start_c, end_c = f_val.split(':')
                    if column_index_from_string(start_c) <= cell_col_idx <= column_index_from_string(end_c):
                        is_match = True; break
                elif f_type == 'range':
                    start_cell, end_cell = f_val.split(':')
                    sc_str, sr_str = re.match(r"([A-Z]+)([0-9]+)", start_cell).groups()
                    ec_str, er_str = re.match(r"([A-Z]+)([0-9]+)", end_cell).groups()
                    if (column_index_from_string(sc_str) <= cell_col_idx <= column_index_from_string(ec_str) and
                        int(sr_str) <= cell_row_idx <= int(er_str)):
                        is_match = True; break
            if not is_match: continue
        filtered_formulas.append(formula_data)
    if controller.current_sort_column:
        col_index = controller.view.tree_columns.index(controller.current_sort_column)
        sort_dir = controller.sort_directions[controller.current_sort_column]
        filtered_formulas.sort(key=lambda x: str(x[col_index]), reverse=(sort_dir == -1))
    count = len(filtered_formulas)
    controller.view.formula_list_label.config(text=f"Formula List ({count} records):")
    for i, data in enumerate(filtered_formulas):
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        item_id = controller.view.result_tree.insert("", "end", values=data, tags=(tag,))
        address_index = controller.view.tree_columns.index("address")
        if address_index < len(data):
            controller.cell_addresses[item_id] = data[address_index]

def sort_column(controller, col_id):
    controller.current_sort_column = col_id
    controller.sort_directions[col_id] *= -1
    apply_filter(controller)
    for column in controller.view.tree_columns:
        original_text = controller.view.result_tree.heading(column, "text").split(' ')[0]
        controller.view.result_tree.heading(column, text=original_text, image='')
    current_direction = " \u2191" if controller.sort_directions[col_id] == 1 else " \u2193"
    current_text = controller.view.result_tree.heading(col_id, "text").split(' ')[0]
    controller.view.result_tree.heading(col_id, text=current_text + current_direction)

def go_to_reference(controller, workbook_path, sheet_name, cell_address):
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
                        # Save original settings to restore later
                        original_display_alerts = controller.xl.DisplayAlerts
                        original_update_links = getattr(controller.xl, 'AskToUpdateLinks', True)
                        
                        # Disable all alerts and update prompts to prevent dialog interruptions
                        controller.xl.DisplayAlerts = False
                        controller.xl.AskToUpdateLinks = False
                        
                        # Open workbook with parameters to avoid dialog boxes
                        target_workbook = controller.xl.Workbooks.Open(
                            Filename=normalized_workbook_path,
                            UpdateLinks=0,  # Don't update any links
                            ReadOnly=False,
                            Format=1,
                            Password="",
                            WriteResPassword="",
                            IgnoreReadOnlyRecommended=True,
                            Notify=False,
                            AddToMru=False
                        )
                        
                        # Restore original settings
                        controller.xl.DisplayAlerts = original_display_alerts
                        controller.xl.AskToUpdateLinks = original_update_links
                        
                    except Exception as e:
                        # Ensure settings are restored even if opening fails
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
                        for wb in controller.xl.Workbooks:
                            if normalized_workbook_path.lower() in wb.Name.lower():
                                target_workbook = wb
                                found_in_open_workbooks = True
                                break
                    
                    if not found_in_open_workbooks:
                        messagebox.showerror("File Not Found", f"The workbook '{filename}' was not found in open workbooks and the path does not exist:\n{normalized_workbook_path}")
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
            
            target_workbook = None
            normalized_workbook_path = os.path.normpath(workbook_path) if workbook_path else None
            
            if normalized_workbook_path:
                for wb in controller.xl.Workbooks:
                    try:
                        if os.path.normpath(wb.FullName) == normalized_workbook_path:
                            target_workbook = wb
                            break
                    except Exception:
                        continue
                
                if not target_workbook:
                    filename = os.path.basename(normalized_workbook_path)
                    for wb in controller.xl.Workbooks:
                        try:
                            if wb.Name.lower() == filename.lower():
                                target_workbook = wb
                                break
                        except Exception:
                            continue
            
            if not target_workbook:
                target_workbook = controller.workbook
            
            target_worksheet = None
            try:
                target_worksheet = target_workbook.Worksheets(sheet_name)
            except Exception as ws_error:
                new_detail_text.insert('end', f"Error accessing worksheet '{sheet_name}': {ws_error}\n", "info_text")
                return
            
            target_cell = None
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    target_worksheet.Activate()
                    target_cell = target_worksheet.Range(cell_address)
                    break
                except Exception as cell_error:
                    if attempt == max_retries - 1:
                        new_detail_text.insert('end', f"Error accessing cell '{cell_address}' after {max_retries} attempts: {cell_error}\n", "info_text")
                        return
                    else:
                        import time
                        time.sleep(0.1)
            
            cell_formula = None
            cell_value = None
            cell_display_value = None
            
            for attempt in range(max_retries):
                try:
                    cell_formula = target_cell.Formula if hasattr(target_cell, 'Formula') and target_cell.Formula else target_cell.Value
                    cell_value = target_cell.Value
                    cell_display_value = target_cell.Text if hasattr(target_cell, 'Text') else str(target_cell.Value)
                    break
                except Exception as value_error:
                    if attempt == max_retries - 1:
                        new_detail_text.insert('end', f"Error reading cell values after {max_retries} attempts: {value_error}\n", "info_text")
                        return
                    else:
                        import time
                        time.sleep(0.1)
            
            if target_cell.Formula and target_cell.Formula.startswith('='):
                cell_type = "formula"
                if any(ref in target_cell.Formula for ref in ['[', ']']):
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
                        read_func = read_external_cell_value
                        referenced_values = get_referenced_cell_values(
                            cell_formula,
                            target_worksheet,
                            target_workbook.FullName,
                            read_func,
                            lambda name, obj: find_matching_sheet(controller.workbook, name)
                        )
                        
                        if referenced_values:
                            new_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
                            for ref_addr, ref_val in referenced_values.items():
                                display_text = ref_addr
                                if '|' in ref_addr:
                                    _, display_text = ref_addr.split('|', 1)

                                new_detail_text.insert('end', f"  {display_text}: {ref_val}  ", "referenced_value")

                                workbook_path_new = None
                                sheet_name_new = None
                                cell_address_to_go_new = None
                                
                                try:
                                    if '|' in ref_addr:
                                        full_path, display_ref = ref_addr.split('|', 1)
                                        workbook_path_new = full_path
                                        
                                        if ']' in display_ref and '!' in display_ref:
                                            sheet_and_cell = display_ref.split(']', 1)[1]
                                            parts = sheet_and_cell.rsplit('!', 1)
                                            sheet_name_new = parts[0].strip("'")
                                            cell_address_to_go_new = parts[1]
                                    else:
                                        workbook_path_new = target_workbook.FullName
                                        if '!' in ref_addr:
                                            parts = ref_addr.rsplit('!', 1)
                                            sheet_name_new = parts[0]
                                            cell_address_to_go_new = parts[1]

                                    if workbook_path_new and sheet_name_new and cell_address_to_go_new:
                                        def build_handler_new(wp, sn, ca, ref_display):
                                            def handler():
                                                go_to_reference_new_tab(controller, wp, sn, ca, ref_display)
                                            return handler
                                        
                                        btn = tk.Button(new_detail_text, text="Go to Reference", font=("Arial", 7), cursor="hand2", command=build_handler_new(workbook_path_new, sheet_name_new, cell_address_to_go_new, display_text))
                                        new_detail_text.window_create('end', window=btn)

                                except Exception as e:
                                    print(f"INFO: Could not create navigation button for '{ref_addr}': {e}")

                                new_detail_text.insert('end', "\n")
                        else:
                            new_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
                            new_detail_text.insert('end', "  No individual cell references found or accessible.\n", "info_text")
                    except Exception as ref_error:
                        new_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
                        new_detail_text.insert('end', f"  Error retrieving referenced values: {ref_error}\n", "info_text")
            else:
                new_detail_text.insert('end', "Content:\n", "label")
                new_detail_text.insert('end', f"{cell_value}\n", "value")
                
        except Exception as e:
            new_detail_text.insert('end', f"Error retrieving cell details: {e}\n", "info_text")
            
    except Exception as e:
        messagebox.showerror("Tab Creation Error", f"Could not create new tab for reference.\nError: {e}")

def on_select(controller, event):
    selected_item = controller.view.result_tree.selection()
    
    try:
        main_tab_info = controller.tab_manager.detail_tabs.get("Main") or controller.tab_manager.detail_tabs.get("Tab_0")
        if not main_tab_info:
            current_detail_text = controller.tab_manager.get_current_detail_text()
        else:
            current_detail_text = main_tab_info["text_widget"]
            controller.tab_manager.detail_notebook.select(main_tab_info["frame"])
    except (AttributeError, KeyError):
        current_detail_text = controller.tab_manager.get_current_detail_text()

    if not selected_item:
        current_detail_text.delete(1.0, 'end')
        return
        
    item_id = selected_item[0]
    values = controller.view.result_tree.item(item_id, "values")
    
    if len(values) < 5:
        current_detail_text.delete(1.0, 'end')
        current_detail_text.insert(1.0, "Selected item has incomplete data.")
        return
        
    formula_type, cell_address, formula, result, display_value = values
    
    current_detail_text.delete(1.0, 'end')
    current_detail_text.insert('end', "Type: ", "label")
    current_detail_text.insert('end', f"{formula_type} / ", "value")
    current_detail_text.insert('end', "Cell Address: ", "label")
    current_detail_text.insert('end', f"{cell_address}\n", "value")
    current_detail_text.insert('end', "Calculated Result: ", "label")
    current_detail_text.insert('end', f"{result} / ", "result_value")
    current_detail_text.insert('end', "Displayed Value: ", "label")
    current_detail_text.insert('end', f"{display_value}\n\n", "value")
    current_detail_text.insert('end', "Formula Content:\n", "label")
    current_detail_text.insert('end', f"{formula}  ", "formula_content")
    
    # Add Explode button next to the formula
    try:
        def build_explode_handler():
            def handler():
                # 使用當前 cell 的信息進行爆炸分析
                if hasattr(controller, 'workbook') and controller.workbook:
                    current_workbook_path = controller.workbook.FullName
                    current_sheet_name = controller.worksheet.Name if hasattr(controller, 'worksheet') and controller.worksheet else "Unknown"
                    
                    # 從選中的項目獲取 cell 地址
                    selected_item = controller.view.result_tree.selection()
                    if selected_item:
                        item_id = selected_item[0]
                        current_cell_address = controller.cell_addresses.get(item_id, "A1")
                        explode_dependencies_popup(controller, current_workbook_path, current_sheet_name, current_cell_address, f"{current_sheet_name}!{current_cell_address}")
                    else:
                        from tkinter import messagebox
                        messagebox.showwarning("No Selection", "Please select a cell first.")
                else:
                    from tkinter import messagebox
                    messagebox.showerror("Excel Not Connected", "Excel connection not available for dependency analysis.")
            return handler
        
        explode_btn = tk.Button(current_detail_text, text="Explode", font=("Arial", 8, "bold"), cursor="hand2", bg="#ffeb3b", command=build_explode_handler())
        current_detail_text.window_create('end', window=explode_btn)
    except Exception as e:
        print(f"Could not create Explode button: {e}")
    
    current_detail_text.insert('end', "\n\n")
    
    # 嘗試獲取引用的儲存格值，但即使失敗也要提供 Go to Reference 功能
    referenced_values = None
    excel_connected = controller.xl and controller.worksheet
    
    if excel_connected:
        try:
            read_func = read_external_cell_value
            referenced_values = get_referenced_cell_values(
                formula,
                controller.worksheet,
                controller.workbook.FullName,
                read_func,
                lambda name, obj: find_matching_sheet(controller.workbook, name)
            )
        except Exception as e:
            print(f"Warning: Could not get referenced values: {e}")
            referenced_values = None
    
    # 解析公式中的引用，即使沒有 Excel 連接也能提供 Go to Reference 功能
    formula_references = []
    if formula and formula.startswith('='):
        try:
            # 解析外部引用 (例如: ='C:\path\[file.xlsx]Sheet'!$A$1)
            import re
            
            # 預處理：標準化公式中的路徑，將雙反斜線轉為單反斜線
            def normalize_formula_paths(formula):
                if not formula:
                    return formula
                
                def normalize_path_match(match):
                    full_match = match.group(0)
                    path_part = match.group(1)
                    normalized_path = os.path.normpath(path_part)
                    return full_match.replace(path_part, normalized_path)
                
                external_ref_pattern = r"'([^']*\[[^\]]+\][^']*)'!"
                return re.sub(external_ref_pattern, normalize_path_match, formula)
            
            normalized_formula = normalize_formula_paths(formula)
            
            external_pattern = r"'([^']*\[[^\]]+\][^']*)'!\$?([A-Z]+)\$?(\d+)"
            external_matches = re.findall(external_pattern, normalized_formula)
            
            # 創建一個副本來移除已處理的外部引用
            remaining_formula = normalized_formula
            
            for match in external_matches:
                full_ref, col, row = match
                # 提取檔案路徑和工作表名稱
                if '[' in full_ref and ']' in full_ref:
                    path_part = full_ref.split('[')[0]
                    file_part = full_ref.split('[')[1].split(']')[0]
                    sheet_part = full_ref.split(']')[1] if ']' in full_ref else 'Sheet1'
                    
                    # 修復路徑中的雙反斜線問題 - 使用更直接的方法
                    # 先解碼 URL 編碼的字符（如 %20）
                    from urllib.parse import unquote
                    decoded_path_part = unquote(path_part)
                    decoded_file_part = unquote(file_part)
                    
                    # 直接組合路徑，然後用 normpath 處理所有斜線問題
                    raw_path = decoded_path_part + decoded_file_part
                    workbook_path = os.path.normpath(raw_path)
                    sheet_name = sheet_part
                    cell_address = f"{col}{row}"
                    
                    formula_references.append({
                        'display': f"{file_part}]{sheet_name}!{cell_address}",
                        'workbook_path': workbook_path,
                        'sheet_name': sheet_name,
                        'cell_address': cell_address,
                        'value': 'N/A (Excel not connected)' if not excel_connected else None
                    })
                    
                    # 從剩餘公式中移除這個外部引用，避免路徑被誤認為 cell address
                    external_ref_full = f"'{full_ref}'!${col}${row}"
                    remaining_formula = remaining_formula.replace(external_ref_full, "")
                    # 也處理沒有 $ 符號的情況
                    external_ref_no_dollar = f"'{full_ref}'!{col}{row}"
                    remaining_formula = remaining_formula.replace(external_ref_no_dollar, "")
            
            # 解析本地引用 (例如: Sheet1!A1, 工作表1!A1)
            # 修復：支援中文工作表名稱，但排除公式開頭的 = 號
            # 使用移除外部引用後的公式
            local_pattern = r"(?<!=)([^'!\[\]=]+)!\$?([A-Z]+)\$?(\d+)"
            local_matches = re.findall(local_pattern, remaining_formula)
            
            for match in local_matches:
                sheet, col, row = match
                if excel_connected and controller.workbook:
                    workbook_path = controller.workbook.FullName
                else:
                    workbook_path = "Current Workbook"
                
                formula_references.append({
                    'display': f"{sheet}!{col}{row}",
                    'workbook_path': workbook_path,
                    'sheet_name': sheet,
                    'cell_address': f"{col}{row}",
                    'value': 'N/A (Excel not connected)' if not excel_connected else None
                })
                
        except Exception as e:
            print(f"Warning: Could not parse formula references: {e}")
    
    # 顯示引用的儲存格值
    current_detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
    
    if referenced_values:
        # 如果有 Excel 連接且成功獲取值，顯示實際值
        for ref_addr, ref_val in referenced_values.items():
            display_text = ref_addr
            if '|' in ref_addr:
                _, display_text = ref_addr.split('|', 1)

            current_detail_text.insert('end', f"  {display_text}: {ref_val}  ", "referenced_value")

            workbook_path = None
            sheet_name = None
            cell_address_to_go = None
            
            try:
                if '|' in ref_addr:
                    full_path, display_ref = ref_addr.split('|', 1)
                    workbook_path = full_path
                    
                    if ']' in display_ref and '!' in display_ref:
                        sheet_and_cell = display_ref.split(']', 1)[1]
                        parts = sheet_and_cell.rsplit('!', 1)
                        sheet_name = parts[0].strip("'")
                        cell_address_to_go = parts[1]
                    else:
                        if display_ref.startswith('[') and ']' in display_ref:
                            bracket_end = display_ref.find(']')
                            file_name = display_ref[1:bracket_end]
                            remaining = display_ref[bracket_end+1:]
                            if '!' in remaining:
                                sheet_name, cell_address_to_go = remaining.split('!', 1)
                                workbook_path = find_external_workbook_path(controller, file_name)
                else:
                    if excel_connected and controller.workbook:
                        workbook_path = controller.workbook.FullName
                    else:
                        workbook_path = None
                    if '!' in ref_addr:
                        parts = ref_addr.rsplit('!', 1)
                        sheet_name = parts[0]
                        cell_address_to_go = parts[1]

                if workbook_path and sheet_name and cell_address_to_go:
                    def build_handler(wp, sn, ca, ref_display):
                        def handler():
                            go_to_reference_with_option(controller, wp, sn, ca, ref_display)
                        return handler
                    
                    # Create frame for buttons
                    btn_frame = tk.Frame(current_detail_text)
                    current_detail_text.window_create('end', window=btn_frame)
                    
                    # Go to Reference button
                    btn = tk.Button(btn_frame, text="Go to Reference", font=("Arial", 7), cursor="hand2", command=build_handler(workbook_path, sheet_name, cell_address_to_go, display_text))
                    btn.pack(side=tk.LEFT, padx=2)
                    
                    # Read Only button (using openpyxl)
                    def build_read_only_handler(wp, sn, ca, ref_display):
                        def handler():
                            read_reference_openpyxl(controller, wp, sn, ca, ref_display)
                        return handler
                    
                    read_btn = tk.Button(btn_frame, text="Read Only", font=("Arial", 7), cursor="hand2", command=build_read_only_handler(workbook_path, sheet_name, cell_address_to_go, display_text))
                    read_btn.pack(side=tk.LEFT, padx=2)
                    
                    # Explode Dependencies button
                    def build_explode_handler(wp, sn, ca, ref_display):
                        def handler():
                            explode_dependencies_popup(controller, wp, sn, ca, ref_display)
                        return handler
                    

            except Exception as e:
                print(f"INFO: Could not create navigation button for '{ref_addr}': {e}")

            current_detail_text.insert('end', "\n")
    elif formula_references:
        # 如果沒有 Excel 連接但解析到引用，仍然提供 Go to Reference 功能
        for ref in formula_references:
            value_text = ref['value'] if ref['value'] else "N/A (Excel not connected)"
            current_detail_text.insert('end', f"  {ref['display']}: {value_text}  ", "referenced_value")
            
            try:
                # Create frame for buttons
                btn_frame = tk.Frame(current_detail_text)
                current_detail_text.window_create('end', window=btn_frame)
                
                # Go to Reference button
                def build_handler(wp, sn, ca, ref_display):
                    def handler():
                        go_to_reference_with_option(controller, wp, sn, ca, ref_display)
                    return handler
                
                btn = tk.Button(btn_frame, text="Go to Reference", font=("Arial", 7), cursor="hand2", command=build_handler(ref['workbook_path'], ref['sheet_name'], ref['cell_address'], ref['display']))
                btn.pack(side=tk.LEFT, padx=2)
                
                # Read Only button (using openpyxl)
                def build_read_only_handler(wp, sn, ca, ref_display):
                    def handler():
                        read_reference_openpyxl(controller, wp, sn, ca, ref_display)
                    return handler
                
                read_btn = tk.Button(btn_frame, text="Read Only", font=("Arial", 7), cursor="hand2", command=build_read_only_handler(ref['workbook_path'], ref['sheet_name'], ref['cell_address'], ref['display']))
                read_btn.pack(side=tk.LEFT, padx=2)
            except Exception as e:
                print(f"INFO: Could not create navigation button for '{ref['display']}': {e}")
            
            current_detail_text.insert('end', "\n")
    else:
        current_detail_text.insert('end', "  No individual cell references found or accessible.\n", "info_text")
    
    if not excel_connected:
        current_detail_text.insert('end', "\nNote: Excel connection not active. Values shown as 'N/A' but Go to Reference still available.\n", "info_text")
        
def on_double_click(controller, event):
    selected_item = controller.view.result_tree.selection()
    if not selected_item:
        return
    item_id = selected_item[0]
    cell_address = controller.cell_addresses.get(item_id)
    if cell_address:
        try:
            try:
                controller.xl = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                controller.xl = None
            if not controller.xl:
                try:
                    controller.xl = win32com.client.Dispatch("Excel.Application")
                    controller.xl.Visible = True
                    if controller.last_workbook_path and os.path.exists(controller.last_workbook_path):
                        controller.workbook = controller.xl.Workbooks.Open(controller.last_workbook_path)
                    else:
                        messagebox.showwarning("File Not Found", "The last scanned Excel file path is not valid or found. Please open Excel manually.")
                        return
                except Exception as e:
                    messagebox.showerror("Excel Launch Error", f"Could not launch Excel or open the workbook.\nError: {e}")
                    return
            target_workbook = None
            
            if controller.last_workbook_path:
                normalized_path = os.path.normpath(controller.last_workbook_path)
                for wb in controller.xl.Workbooks:
                    if os.path.normpath(wb.FullName) == normalized_path:
                        target_workbook = wb
                        break
                
                if not target_workbook and os.path.exists(controller.last_workbook_path):
                    try:
                        # Save original settings to restore later
                        original_display_alerts = controller.xl.DisplayAlerts
                        original_update_links = getattr(controller.xl, 'AskToUpdateLinks', True)
                        
                        # Disable all alerts and update prompts to prevent dialog interruptions
                        controller.xl.DisplayAlerts = False
                        controller.xl.AskToUpdateLinks = False
                        
                        # Open workbook with parameters to avoid dialog boxes
                        target_workbook = controller.xl.Workbooks.Open(
                            Filename=controller.last_workbook_path,
                            UpdateLinks=0,  # Don't update any links
                            ReadOnly=False,
                            Format=1,
                            Password="",
                            WriteResPassword="",
                            IgnoreReadOnlyRecommended=True,
                            Notify=False,
                            AddToMru=False
                        )
                        
                        # Restore original settings
                        controller.xl.DisplayAlerts = original_display_alerts
                        controller.xl.AskToUpdateLinks = original_update_links
                        
                    except Exception as open_e:
                        # Ensure settings are restored even if opening fails
                        try:
                            controller.xl.DisplayAlerts = original_display_alerts
                            controller.xl.AskToUpdateLinks = original_update_links
                        except:
                            pass
                        messagebox.showerror("Workbook Open Error", f"Could not open workbook '{os.path.basename(controller.last_workbook_path)}'.\nError: {open_e}")
                        return
                
                if not target_workbook:
                    filename = os.path.basename(controller.last_workbook_path)
                    for wb in controller.xl.Workbooks:
                        if wb.Name.lower() == filename.lower():
                            target_workbook = wb
                            break
            
            if not target_workbook:
                target_workbook = controller.xl.ActiveWorkbook if controller.xl.ActiveWorkbook else controller.workbook
            
            if not target_workbook:
                messagebox.showerror("Error", "No workbook available to navigate to.")
                return
            
            controller.workbook = target_workbook
            if controller.last_worksheet_name and controller.workbook:
                try:
                    controller.worksheet = controller.workbook.Worksheets(controller.last_worksheet_name)
                except Exception:
                    controller.worksheet = controller.workbook.ActiveSheet
                    messagebox.showwarning("Worksheet Not Found", f"Worksheet '{controller.last_worksheet_name}' not found in '{controller.workbook.Name}'. Activating current sheet.")
            elif controller.workbook:
                controller.worksheet = controller.workbook.ActiveSheet
            else:
                messagebox.showerror("Error", "No active workbook to select cell in.")
                return
            controller.workbook.Activate()
            controller.worksheet.Activate()
            controller.worksheet.Range(cell_address).Select()
            activate_excel_window(controller)
        except Exception as e:
            messagebox.showerror("Excel Selection Error", f"Could not select cell {cell_address} in Excel. Please ensure the workbook and worksheet are still valid.\nError: {e}")

# === Inspect Mode Go to Reference Patch ===
def go_to_reference_inspect_mode(controller, workbook_path, sheet_name, cell_address):
    """
    Inspect Mode 專用的 Go to Reference 函數
    在同一個面板中打開新標籤，而不是跳到左邊面板
    """
    import os
    from tkinter import messagebox
    
    try:
        print(f"[{controller.pane_name}] Go to Reference: {workbook_path} -> {sheet_name}!{cell_address}")
        
        # 檢查檔案是否存在
        if not os.path.exists(workbook_path):
            messagebox.showerror("File Not Found", f"Referenced file not found:\n{workbook_path}")
            return
        
        # 在當前面板中創建新標籤
        if hasattr(controller, 'tab_manager') and controller.tab_manager:
            # 創建標籤標題
            filename = os.path.basename(workbook_path)
            tab_title = f"{filename}!{sheet_name}!{cell_address}"
            
            # 在當前面板中添加新標籤
            controller.tab_manager.add_tab(tab_title, f"Reference: {cell_address}")
            
            print(f"[{controller.pane_name}] Created new tab: {tab_title}")
        
        # 嘗試在 Excel 中打開並跳轉到指定儲存格
        try:
            import win32com.client
            
            # 連接到 Excel
            xl = win32com.client.GetActiveObject("Excel.Application")
            
            # 打開工作簿（如果尚未打開）
            workbook = None
            for wb in xl.Workbooks:
                if wb.FullName == workbook_path:
                    workbook = wb
                    break
            
            if not workbook:
                workbook = xl.Workbooks.Open(workbook_path)
            
            # 切換到指定工作表
            worksheet = workbook.Worksheets(sheet_name)
            worksheet.Activate()
            
            # 選擇指定儲存格
            cell_range = worksheet.Range(cell_address)
            cell_range.Select()
            
            print(f"[{controller.pane_name}] Successfully navigated to {sheet_name}!{cell_address}")
            
        except Exception as excel_error:
            print(f"[{controller.pane_name}] Excel navigation failed: {excel_error}")
            messagebox.showwarning("Excel Navigation", 
                f"Could not navigate to {sheet_name}!{cell_address} in Excel.\n"
                f"Please open the file manually.\n\nError: {excel_error}")
        
    except Exception as e:
        print(f"[{controller.pane_name}] Go to Reference error: {e}")
        messagebox.showerror("Go to Reference Error", f"Could not process reference: {e}")

def is_inspect_mode(controller):
    """檢查當前控制器是否在 Inspect Mode"""
    return hasattr(controller, 'pane_name') and 'Inspect' in str(controller.pane_name)

def go_to_reference_enhanced(controller, workbook_path, sheet_name, cell_address):
    """
    增強版的 Go to Reference 函數
    自動檢測是否在 Inspect Mode 並使用適當的處理方式
    """
    if is_inspect_mode(controller):
        # 在 Inspect Mode 中使用專用函數
        go_to_reference_inspect_mode(controller, workbook_path, sheet_name, cell_address)
    else:
        # 在 Normal Mode 中使用原來的函數
        from core.worksheet_tree import go_to_reference
        go_to_reference(controller, workbook_path, sheet_name, cell_address)

def go_to_reference_with_option(controller, workbook_path, sheet_name, cell_address, reference_display):
    """
    Go to Reference with option to open Excel or not
    Default behavior: open Excel and navigate
    """
    go_to_reference_new_tab(controller, workbook_path, sheet_name, cell_address, reference_display)

def read_reference_openpyxl(controller, workbook_path, sheet_name, cell_address, reference_display):
    """
    Read reference using openpyxl without opening Excel
    Uses the enhanced openpyxl resolver to handle external references
    """
    try:
        from utils.openpyxl_resolver import read_cell_with_resolved_references
        
        # Check if file exists
        if not os.path.exists(workbook_path):
            messagebox.showerror("File Not Found", f"Referenced file not found:\n{workbook_path}")
            return
        
        # Read cell information using enhanced openpyxl
        cell_info = read_cell_with_resolved_references(workbook_path, sheet_name, cell_address)
        
        if 'error' in cell_info:
            messagebox.showerror("Read Error", f"Could not read cell {sheet_name}!{cell_address}:\n{cell_info['error']}")
            return
        
        # Create new tab to display the information
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
        
        # Ensure unique tab name
        counter = 1
        original_tab_name = tab_name
        while tab_name in controller.tab_manager.detail_tabs:
            tab_name = f"{original_tab_name}({counter})"
            counter += 1
        
        # Create new detail tab
        new_detail_text = controller.tab_manager.create_detail_tab(tab_name)
        
        # Display cell information
        new_detail_text.insert('end', "Read Mode: ", "label")
        new_detail_text.insert('end', "openpyxl (Excel not opened)\n", "info_text")
        new_detail_text.insert('end', "Type: ", "label")
        new_detail_text.insert('end', f"{cell_info['cell_type']} / ", "value")
        new_detail_text.insert('end', "Cell Address: ", "label")
        new_detail_text.insert('end', f"{sheet_name}!{cell_address}\n", "value")
        new_detail_text.insert('end', "Workbook: ", "label")
        new_detail_text.insert('end', f"{os.path.basename(workbook_path)}\n", "value")
        
        if cell_info['has_external_references']:
            new_detail_text.insert('end', "External References: ", "label")
            new_detail_text.insert('end', "Resolved ✓\n", "result_value")
        
        new_detail_text.insert('end', "Calculated Result: ", "label")
        new_detail_text.insert('end', f"{cell_info['calculated_value']} / ", "result_value")
        new_detail_text.insert('end', "Displayed Value: ", "label")
        new_detail_text.insert('end', f"{cell_info['display_value']}\n\n", "value")
        
        if cell_info['formula']:
            new_detail_text.insert('end', "Formula Content (External References Resolved):\n", "label")
            new_detail_text.insert('end', f"{cell_info['formula']}  ", "formula_content")
            
            # Add Explode button next to the formula in Read Only mode
            try:
                def build_explode_handler_readonly():
                    def handler():
                        explode_dependencies_popup(controller, workbook_path, sheet_name, cell_address, f"{sheet_name}!{cell_address}")
                    return handler
                
                explode_btn = tk.Button(new_detail_text, text="Explode", font=("Arial", 8, "bold"), cursor="hand2", bg="#ffeb3b", command=build_explode_handler_readonly())
                new_detail_text.window_create('end', window=explode_btn)
            except Exception as e:
                print(f"Could not create Explode button in Read Only mode: {e}")
            
            new_detail_text.insert('end', "\n\n")
            
            # Parse the resolved formula for additional external references
            # and provide Go to Reference buttons for them
            try:
                resolved_formula = cell_info['formula']
                formula_references = []
                
                if resolved_formula and resolved_formula.startswith('='):
                    import re
                    
                    # 預處理：標準化公式中的路徑
                    def normalize_formula_paths_readonly(formula):
                        if not formula:
                            return formula
                        
                        def normalize_path_match(match):
                            full_match = match.group(0)
                            path_part = match.group(1)
                            normalized_path = os.path.normpath(path_part)
                            return full_match.replace(path_part, normalized_path)
                        
                        external_ref_pattern = r"'([^']*\[[^\]]+\][^']*)'!"
                        return re.sub(external_ref_pattern, normalize_path_match, formula)
                    
                    normalized_resolved_formula = normalize_formula_paths_readonly(resolved_formula)
                    
                    # Parse external references (e.g., ='C:\path\[file.xlsx]Sheet'!$A$1)
                    external_pattern = r"'([^']*\[[^\]]+\][^']*)'!\$?([A-Z]+)\$?(\d+)"
                    external_matches = re.findall(external_pattern, normalized_resolved_formula)
                    
                    # 創建一個副本來移除已處理的外部引用
                    remaining_resolved_formula = normalized_resolved_formula
                    
                    for match in external_matches:
                        full_ref, col, row = match
                        # Extract file path and sheet name
                        if '[' in full_ref and ']' in full_ref:
                            path_part = full_ref.split('[')[0]
                            file_part = full_ref.split('[')[1].split(']')[0]
                            sheet_part = full_ref.split(']')[1] if ']' in full_ref else 'Sheet1'
                            
                            # 修復路徑中的雙反斜線問題 - 使用更直接的方法
                            # 先解碼 URL 編碼的字符（如 %20）
                            decoded_path_part = unquote(path_part)
                            decoded_file_part = unquote(file_part)
                            
                            # 直接組合路徑，然後用 normpath 處理所有斜線問題
                            raw_path = decoded_path_part + decoded_file_part
                            workbook_path = os.path.normpath(raw_path)
                            sheet_name = sheet_part
                            cell_address = f"{col}{row}"
                            
                            # 讀取目標 cell 的實際內容
                            try:
                                from utils.openpyxl_resolver import read_cell_with_resolved_references
                                
                                # 修復：如果工作表名稱包含中文或特殊字符，嘗試加上單引號
                                sheet_name_to_use = sheet_name
                                target_cell_info = read_cell_with_resolved_references(workbook_path, sheet_name_to_use, cell_address)
                                
                                # 如果失敗且工作表名稱不是純英文數字，嘗試加單引號
                                if 'error' in target_cell_info and not sheet_name.replace('_', '').isalnum():
                                    sheet_name_to_use = f"'{sheet_name}'"
                                    target_cell_info = read_cell_with_resolved_references(workbook_path, sheet_name_to_use, cell_address)
                                
                                if 'error' in target_cell_info:
                                    cell_value = f"Error: {target_cell_info['error']}"
                                else:
                                    cell_value = target_cell_info.get('display_value', 'N/A')
                            except Exception as e:
                                cell_value = f"Read Error: {str(e)}"
                            
                            formula_references.append({
                                'display': f"[{file_part}]{sheet_name}!{cell_address}",  # 修復：添加開頭的 [
                                'workbook_path': workbook_path,
                                'sheet_name': sheet_name,
                                'cell_address': cell_address,
                                'value': cell_value  # 顯示實際讀取的值
                            })
                            
                            # 從剩餘公式中移除這個外部引用，避免路徑被誤認為 cell address
                            external_ref_full = f"'{full_ref}'!${col}${row}"
                            remaining_resolved_formula = remaining_resolved_formula.replace(external_ref_full, "")
                            # 也處理沒有 $ 符號的情況
                            external_ref_no_dollar = f"'{full_ref}'!{col}{row}"
                            remaining_resolved_formula = remaining_resolved_formula.replace(external_ref_no_dollar, "")
                    
                    # Parse local references (e.g., Sheet1!A1) - but only if not part of external references
                    # First, get all external reference patterns to exclude them
                    external_refs_in_formula = set()
                    for match in external_matches:
                        full_ref, col, row = match
                        if '[' in full_ref and ']' in full_ref:
                            sheet_part = full_ref.split(']')[1] if ']' in full_ref else 'Sheet1'
                            external_refs_in_formula.add(f"{sheet_part}!{col}{row}")
                    
                    # Parse local references using a more robust method
                    # 先移除公式開頭的 = 號，然後尋找所有 worksheet!cell 模式
                    # 使用移除外部引用後的公式
                    formula_without_equals = remaining_resolved_formula[1:] if remaining_resolved_formula.startswith('=') else remaining_resolved_formula
                    
                    # 使用更精確的方法：尋找 ! 符號，然後向前和向後解析
                    import re
                    local_matches = []
                    
                    # 找到所有 ! 的位置
                    exclamation_positions = [i for i, char in enumerate(formula_without_equals) if char == '!']
                    
                    for pos in exclamation_positions:
                        # 向前找工作表名稱
                        start = pos - 1
                        
                        # 檢查是否以單引號結尾（如 'GDP11'!）
                        if start >= 0 and formula_without_equals[start] == "'":
                            # 向前找到開始的單引號
                            quote_start = start - 1
                            while quote_start >= 0 and formula_without_equals[quote_start] != "'":
                                quote_start -= 1
                            
                            if quote_start >= 0:
                                # 提取單引號內的工作表名稱
                                sheet_name = formula_without_equals[quote_start + 1:start]
                            else:
                                continue
                        else:
                            # 沒有單引號，向前找到邊界
                            while start >= 0 and formula_without_equals[start] not in "+'*/-()=,":
                                start -= 1
                            start += 1
                            sheet_name = formula_without_equals[start:pos]
                        
                        # 向後找 cell 地址
                        remaining = formula_without_equals[pos + 1:]
                        cell_match = re.match(r'\$?([A-Z]+)\$?(\d+)', remaining)
                        
                        if cell_match and sheet_name:
                            col, row = cell_match.groups()
                            
                            # 檢查是否為外部引用（包含 [ ] 或已在外部引用列表中）
                            if '[' not in sheet_name and ']' not in sheet_name:
                                ref_key = f"{sheet_name}!{col}{row}"
                                if ref_key not in external_refs_in_formula:
                                    local_matches.append((sheet_name, col, row))
                    
                    for match in local_matches:
                        sheet, col, row = match
                        ref_key = f"{sheet}!{col}{row}"
                        # Skip if it's part of an external reference
                        if ref_key not in external_refs_in_formula:
                            # 讀取本地引用的實際內容
                            try:
                                from utils.openpyxl_resolver import read_cell_with_resolved_references
                                
                                # 修復：如果工作表名稱包含中文或特殊字符，嘗試加上單引號
                                sheet_name_to_use = sheet
                                target_cell_info = read_cell_with_resolved_references(workbook_path, sheet_name_to_use, f"{col}{row}")
                                
                                # 如果失敗且工作表名稱不是純英文數字，嘗試加單引號
                                if 'error' in target_cell_info and not sheet.replace('_', '').isalnum():
                                    sheet_name_to_use = f"'{sheet}'"
                                    target_cell_info = read_cell_with_resolved_references(workbook_path, sheet_name_to_use, f"{col}{row}")
                                
                                if 'error' in target_cell_info:
                                    cell_value = f"Error: {target_cell_info['error']}"
                                else:
                                    cell_value = target_cell_info.get('display_value', 'N/A')
                            except Exception as e:
                                cell_value = f"Read Error: {str(e)}"
                            
                            formula_references.append({
                                'display': f"{sheet}!{col}{row}",
                                'workbook_path': workbook_path,  # Same workbook as the current Read Only tab
                                'sheet_name': sheet,
                                'cell_address': f"{col}{row}",
                                'value': cell_value  # 顯示實際讀取的值
                            })
                
                # Parse relative references (e.g., A12, B5) - cells without worksheet prefix
                if normalized_resolved_formula and normalized_resolved_formula.startswith('='):
                    # 解析相對引用：沒有工作表名稱的 cell 引用
                    relative_pattern = r"(?<![A-Za-z0-9_!'])([A-Z]+)(\d+)(?![A-Za-z0-9_])"
                    relative_matches = re.findall(relative_pattern, formula_without_equals)
                    
                    for col, row in relative_matches:
                        cell_address_rel = f"{col}{row}"
                        
                        # 檢查是否已經在絕對引用中（避免重複）
                        already_exists = any(
                            ref['cell_address'] == cell_address_rel 
                            for ref in formula_references
                        )
                        
                        if not already_exists:
                            # 相對引用使用當前工作表
                            try:
                                from utils.openpyxl_resolver import read_cell_with_resolved_references
                                
                                # 使用當前工作表名稱（從 Cell Address 中提取）
                                current_sheet = sheet_name  # 這是當前 Read Only tab 的工作表
                                
                                target_cell_info = read_cell_with_resolved_references(workbook_path, current_sheet, cell_address_rel)
                                
                                # 如果失敗且工作表名稱包含特殊字符，嘗試加單引號
                                if 'error' in target_cell_info and not current_sheet.replace('_', '').isalnum():
                                    sheet_name_to_use = f"'{current_sheet}'"
                                    target_cell_info = read_cell_with_resolved_references(workbook_path, sheet_name_to_use, cell_address_rel)
                                
                                if 'error' in target_cell_info:
                                    cell_value = f"Error: {target_cell_info['error']}"
                                else:
                                    cell_value = target_cell_info.get('display_value', 'N/A')
                            except Exception as e:
                                cell_value = f"Read Error: {str(e)}"
                            
                            formula_references.append({
                                'display': f"{current_sheet}!{cell_address_rel}",  # 顯示完整引用
                                'workbook_path': workbook_path,
                                'sheet_name': current_sheet,
                                'cell_address': cell_address_rel,
                                'value': cell_value
                            })
                
                # Display referenced cell values with Go to Reference buttons
                if formula_references:
                    new_detail_text.insert('end', "Referenced Cell Values (from Read Only mode):\n", "label")
                    for ref in formula_references:
                        new_detail_text.insert('end', f"  {ref['display']}: {ref['value']}  ", "referenced_value")
                        
                        try:
                            # Create frame for buttons
                            btn_frame = tk.Frame(new_detail_text)
                            new_detail_text.window_create('end', window=btn_frame)
                            
                            # Go to Reference button
                            def build_handler(wp, sn, ca, ref_display):
                                def handler():
                                    go_to_reference_with_option(controller, wp, sn, ca, ref_display)
                                return handler
                            
                            btn = tk.Button(btn_frame, text="Go to Reference", font=("Arial", 7), cursor="hand2", command=build_handler(ref['workbook_path'], ref['sheet_name'], ref['cell_address'], ref['display']))
                            btn.pack(side=tk.LEFT, padx=2)
                            
                            # Read Only button
                            def build_read_only_handler(wp, sn, ca, ref_display):
                                def handler():
                                    read_reference_openpyxl(controller, wp, sn, ca, ref_display)
                                return handler
                            
                            read_btn = tk.Button(btn_frame, text="Read Only", font=("Arial", 7), cursor="hand2", command=build_read_only_handler(ref['workbook_path'], ref['sheet_name'], ref['cell_address'], ref['display']))
                            read_btn.pack(side=tk.LEFT, padx=2)
                            
                            # Explode Dependencies button
                            def build_explode_handler(wp, sn, ca, ref_display):
                                def handler():
                                    explode_dependencies_popup(controller, wp, sn, ca, ref_display)
                                return handler
                            
                        except Exception as e:
                            print(f"INFO: Could not create navigation button for '{ref['display']}': {e}")
                        
                        new_detail_text.insert('end', "\n")
                else:
                    new_detail_text.insert('end', "Referenced Cell Values (from Read Only mode):\n", "label")
                    new_detail_text.insert('end', "  No individual cell references found.\n", "info_text")
                    
            except Exception as e:
                print(f"Warning: Could not parse formula references in Read Only mode: {e}")
                new_detail_text.insert('end', "Referenced Cell Values (from Read Only mode):\n", "label")
                new_detail_text.insert('end', f"  Error parsing references: {e}\n", "info_text")
        else:
            new_detail_text.insert('end', "Content:\n", "label")
            new_detail_text.insert('end', f"{cell_info['calculated_value']}\n", "value")
        
        print(f"Successfully read cell {sheet_name}!{cell_address} using openpyxl (Read Only mode)")
        
    except Exception as e:
        messagebox.showerror("Read Only Error", f"Could not read reference using openpyxl.\nError: {e}")
        print(f"Read Only Error: {e}")
        import traceback
        traceback.print_exc()

def explode_dependencies_popup(controller, workbook_path, sheet_name, cell_address, reference_display):
    """
    彈出視窗顯示公式依賴關係爆炸圖 - 增強版包含進度顯示和日誌累積
    """
    try:
        from utils.progress_enhanced_exploder import explode_cell_dependencies_with_progress, ProgressCallback
        import tkinter as tk
        from tkinter import ttk, messagebox
        
        # 檢查檔案是否存在
        if not os.path.exists(workbook_path):
            messagebox.showerror("File Not Found", f"Referenced file not found:\n{workbook_path}")
            return
        
        # 創建彈出視窗
        popup = tk.Toplevel()
        popup.title(f"Dependency Explosion: {reference_display}")
        popup.geometry("1200x800")  # 增加視窗大小以容納日誌區域
        popup.resizable(True, True)
        
        # 創建主框架
        main_frame = ttk.Frame(popup)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 頂部信息框架
        info_frame = ttk.LabelFrame(main_frame, text="Analysis Info", padding=5)
        info_frame.pack(fill='x', pady=(0, 10))
        
        # 分析按鈕和進度 - 重新設計佈局避免按鈕跳動
        control_frame = ttk.Frame(info_frame)
        control_frame.pack(fill='x')
        
        # 左側按鈕區域 - 固定寬度
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(side=tk.LEFT)
        
        analyze_btn = ttk.Button(button_frame, text="Start Analysis", command=lambda: start_analysis())
        analyze_btn.pack(side=tk.LEFT, padx=5)
        analyze_btn.config(state='normal')  # 預設為可用狀態
        
        # 添加取消按鈕 - 固定位置
        cancel_btn = ttk.Button(button_frame, text="Cancel", state='disabled', command=lambda: cancel_analysis())
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # 添加隱藏/顯示Progress Log按鈕
        toggle_log_btn = ttk.Button(button_frame, text="Hide Log", command=lambda: toggle_log_panel())
        toggle_log_btn.pack(side=tk.LEFT, padx=5)
        
        # 中間進度顯示區域 - 可變寬度
        progress_frame = ttk.Frame(control_frame)
        progress_frame.pack(side=tk.LEFT, fill='x', expand=True, padx=10)
        
        # === 進度顯示增強 ===
        progress_var = tk.StringVar(value="Ready to analyze...")
        progress_label = ttk.Label(progress_frame, textvariable=progress_var, foreground="blue")
        progress_label.pack(side=tk.LEFT)
        
        # 右側進度條區域 - 固定寬度
        progressbar_frame = ttk.Frame(control_frame)
        progressbar_frame.pack(side=tk.RIGHT)
        
        # 添加進度條
        progress_bar = ttk.Progressbar(progressbar_frame, mode='indeterminate', length=200)
        progress_bar.pack(side=tk.RIGHT, padx=10)
        
        # 分析狀態變數
        analysis_running = tk.BooleanVar(value=False)
        analysis_cancelled = tk.BooleanVar(value=False)
        
        # 顯示選項框架
        options_frame = ttk.LabelFrame(info_frame, text="Display Options", padding=5)
        options_frame.pack(fill='x', pady=(5, 0))
        
        # 顯示選項控制 - 第一行
        options_control_frame = ttk.Frame(options_frame)
        options_control_frame.pack(fill='x')
        
        # Cell Address 顯示選項 (調整順序：Address 在前)
        show_full_address_var = tk.BooleanVar(value=False)
        show_full_address_cb = ttk.Checkbutton(
            options_control_frame, 
            text="Show Full Cell Address Paths", 
            variable=show_full_address_var,
            command=lambda: refresh_tree_display()
        )
        show_full_address_cb.pack(side=tk.LEFT, padx=5)
        
        # Formula 顯示選項 (調整順序：Formula 在後)
        show_full_formula_var = tk.BooleanVar(value=False)
        show_full_formula_cb = ttk.Checkbutton(
            options_control_frame, 
            text="Show Full Formula Paths", 
            variable=show_full_formula_var,
            command=lambda: refresh_tree_display()
        )
        show_full_formula_cb.pack(side=tk.LEFT, padx=5)
        
        # Analysis Parameters - 第二行
        params_frame = ttk.Frame(options_frame)
        params_frame.pack(fill='x', pady=(5, 0))
        
        # Range展開閾值設置
        ttk.Label(params_frame, text="Range Expansion:").pack(side=tk.LEFT, padx=5)
        ttk.Label(params_frame, text="Expand ranges with").pack(side=tk.LEFT, padx=2)
        
        # 讀取保存的設定或使用預設值
        saved_range_threshold = getattr(controller, '_saved_range_threshold', 5) if hasattr(controller, '_saved_range_threshold') else 5
        range_threshold_var = tk.IntVar(value=saved_range_threshold)
        
        range_threshold_spinbox = ttk.Spinbox(
            params_frame, 
            from_=1, to=50, width=5,
            textvariable=range_threshold_var
        )
        range_threshold_spinbox.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(params_frame, text="cells or fewer").pack(side=tk.LEFT, padx=2)
        
        # 分隔符
        ttk.Separator(params_frame, orient='vertical').pack(side=tk.LEFT, fill='y', padx=10)
        
        # Max Depth設置
        ttk.Label(params_frame, text="Max Depth:").pack(side=tk.LEFT, padx=5)
        
        # 讀取保存的設定或使用預設值
        saved_max_depth = getattr(controller, '_saved_max_depth', 8) if hasattr(controller, '_saved_max_depth') else 8
        max_depth_var = tk.IntVar(value=saved_max_depth)
        
        max_depth_spinbox = ttk.Spinbox(
            params_frame, 
            from_=1, to=20, width=5,
            textvariable=max_depth_var
        )
        max_depth_spinbox.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(params_frame, text="levels deep").pack(side=tk.LEFT, padx=2)
        

        def update_params_preview():
            """更新參數預覽"""
            global _last_range_threshold, _last_max_depth
            
            try:
                # Range threshold處理
                range_val = range_threshold_var.get()
                if range_val:
                    _last_range_threshold = range_val
                
                # Max depth處理
                depth_str = str(max_depth_var.get()).strip()
                if depth_str and depth_str != "":
                    try:
                        depth_val = int(float(depth_str))
                        _last_max_depth = depth_val
                    except (ValueError, TypeError):
                        depth_val = _last_max_depth
                        max_depth_var.set(str(_last_max_depth))
                else:
                    depth_val = _last_max_depth
                    max_depth_var.set(str(_last_max_depth))
                
                progress_var.set(f"Ready to analyze with Range Threshold: {_last_range_threshold}, Max Depth: {depth_val}. Click 'Start Analysis' to begin.")
                
            except Exception as e:
                print(f"Error in update_params_preview: {e}")
                # 使用默認值
                _last_range_threshold = 5
                _last_max_depth = 10
                range_threshold_var.set(_last_range_threshold)
                max_depth_var.set(str(_last_max_depth))
                progress_var.set(f"Ready to analyze with Range Threshold: {_last_range_threshold}, Max Depth: {_last_max_depth}. Click 'Start Analysis' to begin.")

        # 圖表生成按鈕
        def handle_generate_graph():
            if not hasattr(refresh_tree_display, 'tree_data') or not refresh_tree_display.tree_data:
                messagebox.showwarning("No Data", "Please run the analysis first to generate data for the graph.")
                return
            
            try:
                progress_var.set("Generating graph...")
                popup.update()
                
                # 1. 轉換資料
                nodes_data, edges_data = convert_tree_to_graph_data(refresh_tree_display.tree_data)
                
                if not nodes_data:
                    messagebox.showinfo("Empty Graph", "The analysis result is empty, nothing to graph.")
                    return

                # 2. 產生圖表
                graph_gen = GraphGenerator(nodes_data, edges_data)
                graph_gen.generate_graph()
                
                progress_var.set("Graph generated successfully and opened in browser.")

            except Exception as e:
                messagebox.showerror("Graph Generation Error", f"Failed to generate graph:\n{e}")
                progress_var.set(f"Graph generation failed: {e}")

        graph_btn = ttk.Button(
            options_control_frame, 
            text="Generate Graph", 
            command=handle_generate_graph
        )
        graph_btn.pack(side=tk.RIGHT, padx=5)
        
        # === 創建可切換的主要內容區域 ===
        content_paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        content_paned.pack(fill='both', expand=True)
        
        # 左側：樹狀視圖框架
        tree_frame = ttk.LabelFrame(content_paned, text="Dependency Tree", padding=5)
        content_paned.add(tree_frame, weight=3)  # 佔較大比例
        
        # 右側：進度日誌框架
        log_frame = ttk.LabelFrame(content_paned, text="Progress Log", padding=5)
        content_paned.add(log_frame, weight=1)  # 佔較小比例
        
        # 日誌面板顯示狀態
        log_panel_visible = tk.BooleanVar(value=True)
        
        # 創建 Treeview
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        dependency_tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set)
        dependency_tree.pack(fill='both', expand=True)
        tree_scroll.config(command=dependency_tree.yview)
        
        # 設置列 - 為INDIRECT支持添加resolved列
        dependency_tree['columns'] = ('formula', 'resolved', 'value', 'type', 'depth')
        dependency_tree.column('#0', width=300, minwidth=200)
        dependency_tree.column('formula', width=350, minwidth=200)
        dependency_tree.column('resolved', width=350, minwidth=200)
        dependency_tree.column('value', width=150, minwidth=100)
        dependency_tree.column('type', width=100, minwidth=80)
        dependency_tree.column('depth', width=80, minwidth=60)
        
        # 設置標題
        dependency_tree.heading('#0', text='Cell Address', anchor=tk.W)
        dependency_tree.heading('formula', text='Formula', anchor=tk.W)
        dependency_tree.heading('resolved', text='Resolved', anchor=tk.W)
        dependency_tree.heading('value', text='Value', anchor=tk.W)
        dependency_tree.heading('type', text='Type', anchor=tk.W)
        dependency_tree.heading('depth', text='Depth', anchor=tk.W)
        
        # === 進度日誌區域 ===
        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        log_text = tk.Text(log_frame, yscrollcommand=log_scroll.set, wrap=tk.WORD, 
                          font=("Consolas", 9), bg="#f8f8f8", state='disabled')
        log_text.pack(fill='both', expand=True)
        log_scroll.config(command=log_text.yview)
        
        # 添加清除日誌按鈕
        clear_log_btn = ttk.Button(log_frame, text="Clear Log", 
                                  command=lambda: clear_log())
        clear_log_btn.pack(pady=2)
        
        def clear_log():
            """清除日誌內容"""
            log_text.config(state='normal')
            log_text.delete(1.0, tk.END)
            log_text.config(state='disabled')
        
        def toggle_log_panel():
            """切換日誌面板顯示/隱藏"""
            if log_panel_visible.get():
                # 隱藏日誌面板
                content_paned.remove(log_frame)
                toggle_log_btn.config(text="Show Log")
                log_panel_visible.set(False)
            else:
                # 顯示日誌面板
                content_paned.add(log_frame, weight=1)
                toggle_log_btn.config(text="Hide Log")
                log_panel_visible.set(True)
        
        # 底部摘要框架
        summary_frame = ttk.LabelFrame(main_frame, text="Analysis Summary", padding=5)
        summary_frame.pack(fill='x', pady=(10, 0))
        
        summary_text = tk.Text(summary_frame, height=4, wrap=tk.WORD)
        summary_text.pack(fill='x')
        
        def cancel_analysis():
            """取消分析"""
            analysis_cancelled.set(True)
            cancel_btn.config(state='disabled')
            progress_bar.stop()
            progress_var.set("Analysis cancelled by user.")
        
        def start_analysis():
            """開始依賴關係分析 - 增強版包含進度顯示和日誌累積"""
            try:
                # 設置分析狀態
                analysis_running.set(True)
                analysis_cancelled.set(False)
                
                # 更新UI狀態
                analyze_btn.config(state='disabled')
                cancel_btn.config(state='normal')
                progress_bar.start(10)  # 開始進度條動畫
                
                progress_var.set("Initializing analysis...")
                popup.update()
                
                # 清空樹狀視圖和日誌
                for item in dependency_tree.get_children():
                    dependency_tree.delete(item)
                clear_log()
                
                # === 創建進度回調 ===
                class PopupProgressCallback:
                    def __init__(self, progress_var, popup_window, log_text_widget, cancelled_var):
                        self.progress_var = progress_var
                        self.popup_window = popup_window
                        self.log_text_widget = log_text_widget
                        self.cancelled_var = cancelled_var
                        
                    def update_progress(self, message, step=None):
                        """更新進度訊息 - 同時更新實時顯示和累積日誌"""
                        # 檢查是否被取消
                        if self.cancelled_var.get():
                            raise Exception("Analysis cancelled by user")
                            
                        # 更新實時進度標籤
                        if self.progress_var:
                            self.progress_var.set(message)
                            
                        # 累積到日誌區域
                        if self.log_text_widget:
                            try:
                                import datetime
                                timestamp = datetime.datetime.now().strftime("%H:%M:%S")
                                log_entry = f"[{timestamp}] {message}\n"
                                
                                self.log_text_widget.config(state='normal')
                                self.log_text_widget.insert('end', log_entry)
                                self.log_text_widget.see('end')  # 自動滾動到最新
                                self.log_text_widget.config(state='disabled')
                            except Exception as e:
                                print(f"Log update error: {e}")
                                
                        # 更新視窗
                        if self.popup_window:
                            try:
                                self.popup_window.update()
                            except:
                                pass  # 視窗可能已關閉
                                
                        # 控制台輸出
                        print(f"[Explode Progress] {message}")
                
                # 創建進度回調
                progress_callback = PopupProgressCallback(progress_var, popup, log_text, analysis_cancelled)
                
                # 保存用戶設定供下次使用
                controller._saved_range_threshold = range_threshold_var.get()
                controller._saved_max_depth = max_depth_var.get()
                
                # 執行爆炸分析 - 使用用戶設定的參數
                dependency_tree_data, summary = explode_cell_dependencies_with_progress(
                    workbook_path, sheet_name, cell_address, 
                    max_depth=max_depth_var.get(), 
                    range_expand_threshold=range_threshold_var.get(),
                    progress_callback=progress_callback
                )
                
                # 檢查是否被取消
                if analysis_cancelled.get():
                    progress_var.set("Analysis was cancelled.")
                    return
                
                # 儲存樹狀數據供刷新使用
                refresh_tree_display.tree_data = dependency_tree_data
                
                # 填充樹狀視圖
                progress_var.set("Populating tree view...")
                popup.update()
                populate_tree(dependency_tree_data)
                
                # 顯示摘要
                show_summary(summary)
                
                progress_var.set(f"Analysis complete! Found {summary['total_nodes']} nodes, max depth: {summary['max_depth']}")
                
            except Exception as e:
                if "cancelled" in str(e).lower():
                    progress_var.set("Analysis cancelled by user.")
                else:
                    messagebox.showerror("Analysis Error", f"Could not analyze dependencies:\n{str(e)}")
                    progress_var.set(f"Analysis failed: {str(e)}")
            finally:
                # 恢復UI狀態
                analysis_running.set(False)
                analyze_btn.config(state='normal')
                cancel_btn.config(state='disabled')
                progress_bar.stop()
        
        def format_formula_display(formula):
            """根據顯示選項格式化公式"""
            if not formula:
                return formula
            
            if not show_full_formula_var.get():
                # 簡化顯示：移除完整路徑，只保留檔案名
                import re
                # 匹配外部引用模式並簡化
                def simplify_external_ref(match):
                    full_path = match.group(1)
                    if '[' in full_path and ']' in full_path:
                        # 提取檔案名部分
                        file_part = full_path.split('[')[1].split(']')[0]
                        sheet_part = full_path.split(']')[1] if ']' in full_path else ''
                        return f"'[{file_part}]{sheet_part}'"
                    return match.group(0)
                
                # 簡化外部引用路徑
                simplified = re.sub(r"'([^']*\[[^\]]+\][^']*)'", simplify_external_ref, formula)
                return simplified
            else:
                # 完整顯示
                return formula
        
        def format_address_display(address, node):
            """根據顯示選項格式化地址"""
            if not show_full_address_var.get():
                # 簡化顯示：使用 short_address 格式
                return node.get('short_address', address)
            else:
                # 完整顯示：使用 full_address 格式
                return node.get('full_address', address)
        

        def populate_tree(node, parent=''):
            """遞歸填充樹狀視圖"""
            try:
                # 準備顯示數據
                raw_address = node.get('address', 'Unknown')
                raw_formula = node.get('formula', '')
                
                # 根據顯示選項格式化
                address = format_address_display(raw_address, node)
                formula = format_formula_display(raw_formula)
                
                # === 修復：處理所有動態函數的resolved formula ===
                resolved_formula = ""
                # 檢查是否有任何動態函數解析
                if (node.get('has_indirect', False) or 
                    node.get('has_index', False)):  # 添加INDEX檢查
                    raw_resolved = node.get('resolved_formula', '')
                    # 不要截短resolved formula，完整顯示
                    resolved_formula = format_formula_display(raw_resolved)
                
                value = str(node.get('value', ''))
                if len(value) > 20:
                    value = value[:17] + "..."
                
                node_type = node.get('type', 'unknown')
                depth = node.get('depth', 0)
                
                # 根據類型設置圖標
                if node_type == 'formula':
                    icon = "📊"
                elif node_type == 'value':
                    icon = "🔢"
                elif node_type == 'error':
                    icon = "❌"
                elif node_type == 'circular_ref':
                    icon = "🔄"
                elif node_type == 'limit_reached':
                    icon = "⚠️"
                elif node_type == 'range':
                    icon = "📋"  # 範圍使用表格圖標
                else:
                    icon = "📄"
                
                # 插入節點 - 包含resolved列
                item_id = dependency_tree.insert(
                    parent, 'end',
                    text=f"{icon} {address}",
                    values=(formula, resolved_formula, value, node_type, depth)
                )
                
                # 儲存完整的節點詳細信息到 tags 中，供雙擊導航使用
                node_details = {
                    'workbook_path': node.get('workbook_path', workbook_path),
                    'sheet_name': node.get('sheet_name', ''),
                    'cell_address': node.get('cell_address', ''),
                    'original_formula': raw_formula,  # 儲存原始完整公式
                    'calculated_value': node.get('calculated_value', ''),
                    'display_value': node.get('value', ''),
                    'node_type': node_type,
                    'depth': depth,
                    'address': raw_address,  # 儲存原始地址
                    'display_formula': formula,  # 儲存格式化後的公式
                    'display_address': address  # 儲存格式化後的地址
                }
                
                # 將詳細信息序列化並儲存到 tags 中
                import json
                try:
                    details_json = json.dumps(node_details)
                    dependency_tree.item(item_id, tags=(details_json,))
                except Exception as e:
                    print(f"Warning: Could not serialize node details: {e}")
                    # 如果序列化失敗，至少儲存基本信息
                    basic_info = f"{node_details['workbook_path']}|{node_details['sheet_name']}|{node_details['cell_address']}"
                    dependency_tree.item(item_id, tags=(basic_info,))
                
                # 遞歸添加子節點
                for child in node.get('children', []):
                    populate_tree(child, item_id)
                
                # 展開前幾層
                if depth < 3:
                    dependency_tree.item(item_id, open=True)
                    
            except Exception as e:
                print(f"Error populating tree node: {e}")
        
        def show_summary(summary):
            """顯示分析摘要"""
            summary_text.delete(1.0, tk.END)
            
            summary_content = f"""Total Nodes: {summary['total_nodes']}
Maximum Depth: {summary['max_depth']}
Circular References: {summary['circular_references']}

Node Type Distribution:
"""
            for node_type, count in summary['type_distribution'].items():
                summary_content += f"  {node_type}: {count}\n"
            
            if summary['circular_ref_list']:
                summary_content += f"\nCircular References Found:\n"
                for ref in summary['circular_ref_list']:
                    summary_content += f"  {ref}\n"
            
            summary_text.insert(1.0, summary_content)
        
        def refresh_tree_display():
            """刷新樹狀視圖顯示，應用新的顯示選項"""
            try:
                # 保存當前展開狀態
                expanded_items = []
                def save_expanded_state(item=''):
                    children = dependency_tree.get_children(item)
                    for child in children:
                        if dependency_tree.item(child, 'open'):
                            expanded_items.append(child)
                        save_expanded_state(child)
                
                save_expanded_state()
                
                # 重新填充樹狀視圖
                if hasattr(refresh_tree_display, 'tree_data'):
                    # 清空現有內容
                    for item in dependency_tree.get_children():
                        dependency_tree.delete(item)
                    
                    # 重新填充
                    populate_tree(refresh_tree_display.tree_data)
                    
                    # 恢復展開狀態（盡可能）
                    for item_id in expanded_items:
                        try:
                            dependency_tree.item(item_id, open=True)
                        except:
                            pass  # 如果項目不存在就忽略
                    
                    print(f"Tree display refreshed with new options:")
                    print(f"  Show Full Formula Paths: {show_full_formula_var.get()}")
                    print(f"  Show Full Address Paths: {show_full_address_var.get()}")
                else:
                    print("No tree data available for refresh")
                    
            except Exception as e:
                print(f"Error refreshing tree display: {e}")
        
        # 雙擊事件：Go to Reference
        def on_tree_double_click(event):
            """樹狀視圖雙擊事件 - 使用儲存的詳細信息進行準確導航"""
            try:
                if not dependency_tree.selection():
                    return
                    
                item = dependency_tree.selection()[0]
                item_text = dependency_tree.item(item, "text")
                tags = dependency_tree.item(item, "tags")
                
                # 提取地址信息（移除圖標）
                address_part = item_text.split(" ", 1)[1] if " " in item_text else item_text
                
                print(f"Double-click on: {address_part}")  # 調試信息
                
                # 嘗試從 tags 中獲取詳細信息
                node_details = None
                if tags:
                    import json
                    try:
                        # 嘗試解析 JSON 格式的詳細信息
                        node_details = json.loads(tags[0])
                        print(f"Loaded node details: {node_details}")
                    except (json.JSONDecodeError, ValueError):
                        # 如果不是 JSON，嘗試解析基本格式 "workbook|sheet|cell"
                        try:
                            parts = tags[0].split('|')
                            if len(parts) >= 3:
                                node_details = {
                                    'workbook_path': parts[0],
                                    'sheet_name': parts[1],
                                    'cell_address': parts[2]
                                }
                                print(f"Parsed basic node details: {node_details}")
                        except Exception as e:
                            print(f"Could not parse basic node details: {e}")
                
                # 使用詳細信息進行導航
                if node_details and 'workbook_path' in node_details:
                    target_workbook_path = node_details['workbook_path']
                    sheet_name = node_details.get('sheet_name', '')
                    cell_address = node_details.get('cell_address', '')
                    
                    print(f"Navigation using stored details:")
                    print(f"  Workbook: {target_workbook_path}")
                    print(f"  Sheet: {sheet_name}")
                    print(f"  Cell: {cell_address}")
                    
                    # 檢查檔案是否存在
                    import os
                    if not os.path.exists(target_workbook_path):
                        # 如果檔案不存在，嘗試在當前目錄尋找
                        filename = os.path.basename(target_workbook_path)
                        base_dir = os.path.dirname(workbook_path)
                        alt_path = os.path.join(base_dir, filename)
                        
                        if os.path.exists(alt_path):
                            target_workbook_path = alt_path
                            print(f"Found alternative path: {target_workbook_path}")
                        else:
                            print(f"Warning: File not found: {target_workbook_path}")
                    
                    go_to_reference_new_tab(controller, target_workbook_path, sheet_name, cell_address, address_part)
                    
                else:
                    # 回退到原來的解析方法
                    print("No stored details found, using fallback parsing...")
                    
                    if "!" in address_part:
                        # 解析地址格式：可能是 [filename]sheet!cell 或 sheet!cell
                        if address_part.startswith('[') and ']' in address_part:
                            # 外部引用格式：[filename]sheet!cell
                            bracket_end = address_part.find(']')
                            filename = address_part[1:bracket_end]  # 提取檔案名
                            remaining = address_part[bracket_end+1:]  # sheet!cell
                            
                            if '!' in remaining:
                                sheet_part, cell_part = remaining.split('!', 1)
                                
                                # 構建完整檔案路徑
                                import os
                                base_dir = os.path.dirname(workbook_path)
                                target_workbook_path = os.path.join(base_dir, filename + '.xlsx')
                                
                                # 如果檔案不存在，嘗試其他副檔名
                                if not os.path.exists(target_workbook_path):
                                    for ext in ['.xls', '.xlsm']:
                                        alt_path = os.path.join(base_dir, filename + ext)
                                        if os.path.exists(alt_path):
                                            target_workbook_path = alt_path
                                            break
                                
                                print(f"Fallback external reference navigation: {target_workbook_path} -> {sheet_part}!{cell_part}")
                                go_to_reference_new_tab(controller, target_workbook_path, sheet_part, cell_part, address_part)
                            else:
                                messagebox.showwarning("Parse Error", f"Could not parse address: {address_part}")
                        else:
                            # 本地引用格式：sheet!cell
                            sheet_part, cell_part = address_part.split("!", 1)
                            print(f"Fallback local reference navigation: {workbook_path} -> {sheet_part}!{cell_part}")
                            go_to_reference_new_tab(controller, workbook_path, sheet_part, cell_part, address_part)
                    else:
                        messagebox.showwarning("Parse Error", f"Invalid address format: {address_part}")
                    
            except Exception as e:
                messagebox.showerror("Navigation Error", f"Could not navigate to {address_part}:\n{str(e)}")
                print(f"Navigation error: {e}")
                import traceback
                traceback.print_exc()
        
        dependency_tree.bind("<Double-1>", on_tree_double_click)
        
        # 右鍵菜單
        def show_context_menu(event):
            """顯示右鍵菜單"""
            try:
                item = dependency_tree.identify_row(event.y)
                if item:
                    dependency_tree.selection_set(item)
                    
                    context_menu = tk.Menu(popup, tearoff=0)
                    context_menu.add_command(label="Go to Reference", command=lambda: on_tree_double_click(None))
                    context_menu.add_command(label="Copy Address", command=lambda: copy_address(item))
                    context_menu.add_separator()
                    context_menu.add_command(label="Expand All", command=lambda: expand_all(item))
                    context_menu.add_command(label="Collapse All", command=lambda: collapse_all(item))
                    
                    context_menu.post(event.x_root, event.y_root)
            except Exception as e:
                print(f"Context menu error: {e}")
        
        def copy_address(item):
            """複製地址到剪貼板"""
            item_text = dependency_tree.item(item, "text")
            address_part = item_text.split(" ", 1)[1] if " " in item_text else item_text
            popup.clipboard_clear()
            popup.clipboard_append(address_part)
        
        def expand_all(item):
            """展開所有子節點"""
            dependency_tree.item(item, open=True)
            for child in dependency_tree.get_children(item):
                expand_all(child)
        
        def collapse_all(item):
            """收縮所有子節點"""
            dependency_tree.item(item, open=False)
            for child in dependency_tree.get_children(item):
                collapse_all(child)
        
        dependency_tree.bind("<Button-3>", show_context_menu)
        
        # 不再自動開始分析，等待用戶手動點擊Start Analysis
        # progress_var會由update_params_preview()設置
        
        print(f"Opened dependency explosion popup for {reference_display}")
        
    except Exception as e:
        messagebox.showerror("Explosion Error", f"Could not create dependency explosion view:\nError: {e}")
        print(f"Explosion Error: {e}")
        import traceback
        traceback.print_exc()
