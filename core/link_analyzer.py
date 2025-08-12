# -*- coding: utf-8 -*-
"""
Link Analyzer Module

This module contains functions for analyzing and processing Excel formula links,
including external references and local references.
"""

import re
import os


def is_external_link_regex_match(formula_str):
    """
    Check if a formula string contains external link references.
    
    Args:
        formula_str (str): The formula string to check
        
    Returns:
        bool: True if external link pattern is found, False otherwise
    """
    external_link_pattern = re.compile(r"\[([^\]]+?\.(?:xlsx|xls|xlsm|xlsb))\]", re.IGNORECASE)
    return bool(external_link_pattern.search(formula_str))


def get_referenced_cell_values(
    formula_str, 
    current_sheet_com_obj, 
    current_workbook_path,
    read_external_cell_value_func,
    find_matching_sheet_func
):
    """
    Extract and retrieve values from all cell references in a formula.
    
    Args:
        formula_str (str): The formula string to analyze
        current_sheet_com_obj: COM object of the current worksheet
        current_workbook_path (str): Path to the current workbook
        read_external_cell_value_func: Function to read external cell values
        find_matching_sheet_func: Function to find matching worksheets
        
    Returns:
        dict: Dictionary mapping reference addresses to their values
    """
    referenced_data = {}
    processed_spans = []

    def is_span_processed(start, end):
        for p_start, p_end in processed_spans:
            if start < p_end and end > p_start:
                return True
        return False

    def add_processed_span(start, end):
        processed_spans.append((start, end))

    patterns = [
        (
            'external',
            re.compile(
                r"'?((?:[a-zA-Z]:\\)?[^']*)\[([^\]]+\.(?:xlsx|xls|xlsm|xlsb))\]([^']*)'?\s*!\s*(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)",
                re.IGNORECASE
            )
        ),
        (
            'local_quoted',
            re.compile(
                r"'([^']+)'!(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)",
                re.IGNORECASE
            )
        ),
        (
            'local_unquoted',
            re.compile(
                r"([a-zA-Z0-9_\u4e00-\u9fa5][a-zA-Z0-9_\s\.\u4e00-\u9fa5]{0,30})!(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)",
                re.IGNORECASE
            )
        ),
        (
            'current_range',
            re.compile(
                r"(?<![!'\[\]a-zA-Z0-9_\u4e00-\u9fa5])(\$?[A-Z]{1,3}\$?\d{1,7}:\s*\$?[A-Z]{1,3}\$?\d{1,7})(?![a-zA-Z0-9_])",
                re.IGNORECASE
            )
        ),
        (
            'current_single',
            re.compile(
                r"(?<![!'\[\]a-zA-Z0-9_\u4e00-\u9fa5])(\$?[A-Z]{1,3}\$?\d{1,7})(?![a-zA-Z0-9_:\(])",
                re.IGNORECASE
            )
        )
    ]

    all_matches = []
    # Normalize backslashes to handle cases with single or double backslashes
    normalized_formula_str = formula_str.replace('\\\\', '\\')
    for p_type, pattern in patterns:
        for match in pattern.finditer(normalized_formula_str):
            all_matches.append({'type': p_type, 'match': match, 'span': match.span()})

    all_matches.sort(key=lambda x: (x['span'][0], x['span'][1] - x['span'][0]))

    for item in all_matches:
        match = item['match']
        m_type = item['type']
        start, end = item['span']

        if is_span_processed(start, end):
            continue

        try:
            if m_type == 'external':
                dir_path, file_name, sheet_name, cell_ref = match.groups()
                sheet_name = sheet_name.strip("'")
                
                full_file_path = os.path.join(dir_path, file_name)
                if not dir_path and file_name.lower() == os.path.basename(current_workbook_path).lower():
                    full_file_path = current_workbook_path
                
                display_ref = f"[{os.path.basename(full_file_path)}]{sheet_name}!{cell_ref.replace('$', '')}"
                display_ref_with_path = f"{full_file_path}|{display_ref}"

                if ':' in cell_ref:
                    value = "(Range Reference)"
                else:
                    value = read_external_cell_value_func(
                        current_workbook_path, full_file_path, sheet_name, cell_ref.replace('$', '')
                    )
                if display_ref_with_path not in referenced_data:
                    referenced_data[display_ref_with_path] = value

            elif m_type in ('local_quoted', 'local_unquoted'):
                sheet_name, cell_ref = match.groups()
                sheet_name = sheet_name.strip("'")
                
                if sheet_name.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
                    continue

                display_ref = f"{sheet_name}!{cell_ref.replace('$', '')}"
                
                if ':' in cell_ref:
                    value = "(Range Reference)"
                else:
                    target_sheet = find_matching_sheet_func(sheet_name, current_sheet_com_obj)
                    if target_sheet:
                        cell_val = target_sheet.Range(cell_ref).Value
                        value = f"Local: {cell_val if cell_val is not None else 'Empty'}"
                    else:
                        value = f"Local (Sheet '{sheet_name}' Not Found)"
                
                if display_ref not in referenced_data:
                    referenced_data[display_ref] = value

            elif m_type in ('current_range', 'current_single'):
                cell_ref = match.group(1)
                display_ref = f"{current_sheet_com_obj.Name}!{cell_ref.replace('$', '')}"

                if ':' in cell_ref:
                    value = "(Range Reference)"
                else:
                    cell_val = current_sheet_com_obj.Range(cell_ref).Value
                    value = f"Current: {cell_val if cell_val is not None else 'Empty'}"
                
                if display_ref not in referenced_data:
                    referenced_data[display_ref] = value

            add_processed_span(start, end)
        except Exception as e:
            print(f"ERROR: Could not process reference from match '{match.group(0)}': {e}")

    return referenced_data


def parse_external_path_and_sheet(path_and_sheet):
    """
    Parse external path and sheet information from a reference string.
    
    Args:
        path_and_sheet (str): String containing path and sheet information
        
    Returns:
        tuple: (file_name, sheet_name)
    """
    if '[' in path_and_sheet and ']' in path_and_sheet:
        file_and_sheet = path_and_sheet.split('[')[1]
        if ']' in file_and_sheet:
            file_name = file_and_sheet.split(']')[0]
            sheet_name_part = file_and_sheet.split(']')[1]
            if sheet_name_part.startswith("'"):
                sheet_name = sheet_name_part.lstrip("'")
            else:
                sheet_name = sheet_name_part.strip('!')
        else:
            file_name = file_and_sheet
            sheet_name = ''
    else:
        file_name = ''
        sheet_name = path_and_sheet.strip('!')
    return file_name, sheet_name