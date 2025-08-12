# -*- coding: utf-8 -*-
"""
INDEX Solver - 從 progress_enhanced_exploder.py 中提取的INDEX解析邏輯
純粹的程式碼搬移，不修改任何邏輯
"""

import re

class IndexSolver:
    """INDEX函數解析器 - 從原始程式碼中完全搬移"""
    
    def __init__(self, excel_manager, progress_callback, main_analyzer=None):
        self.excel_manager = excel_manager
        self.progress_callback = progress_callback
        self.main_analyzer = main_analyzer
    
    def _resolve_index_with_excel_corrected_simple(self, formula, workbook_path, sheet_name, cell_address):
        """正確的INDEX解析 - 簡化版本"""
        try:
            self.progress_callback.update_progress(f"[INDEX-SIMPLE] 開始解析: {formula}")
            
            # 1. 提取所有 INDEX 函數
            index_functions = self._extract_all_index_functions_debug(formula)
            if not index_functions:
                return {'success': False, 'error': 'No INDEX functions found'}
            
            resolved_formula = formula
            static_references = []
            calculation_details = []
            internal_references = []
            
            # 2. 逐個解析 INDEX 函數
            for i, index_func in enumerate(index_functions):
                self.progress_callback.update_progress(f"[INDEX-SIMPLE] 處理INDEX#{i+1}: {index_func['content']}")
                
                # 3. 解析參數
                params_result = self._extract_index_parameters_accurate_debug(index_func['content'])
                if not params_result['success']:
                    continue
                    
                array_param = params_result['array']
                row_param = params_result['row']
                col_param = params_result['column']
                
                self.progress_callback.update_progress(f"[INDEX-SIMPLE] 參數: array='{array_param}', row='{row_param}', col='{col_param}'")
                
                # 4. 分析array範圍的內部引用
                if self.main_analyzer:
                    temp_references = self.main_analyzer._parse_formula_references_accurate(f"={array_param}", workbook_path, sheet_name)
                    internal_references.extend(temp_references)
                
                # 5. 檢查row和col是否為簡單數字
                try:
                    if self._is_simple_number(row_param) and self._is_simple_number(col_param):
                        # 直接使用數字，不需要Excel計算
                        row_value = int(float(row_param))
                        col_value = int(float(col_param))
                        self.progress_callback.update_progress(f"[INDEX-SIMPLE] 使用直接數值: row={row_value}, col={col_value}")
                    else:
                        # 只有複雜公式才需要Excel計算
                        self.progress_callback.update_progress(f"[INDEX-SIMPLE] 複雜參數，需要Excel計算...")
                        row_calc = self.excel_manager.calculate_safely(row_param, workbook_path, sheet_name, cell_address)
                        col_calc = self.excel_manager.calculate_safely(col_param, workbook_path, sheet_name, cell_address)
                        
                        if not row_calc['success'] or not col_calc['success']:
                            continue
                            
                        row_value = int(float(row_calc['static_reference']))
                        col_value = int(float(col_calc['static_reference']))
                        
                except Exception as e:
                    self.progress_callback.update_progress(f"[INDEX-SIMPLE] 參數處理失敗: {e}")
                    continue
                
                # 6. 手動構建靜態引用
                static_ref_result = self._build_static_reference_from_index_simple(
                    array_param, row_value, col_value, workbook_path, sheet_name
                )
                
                if static_ref_result['success']:
                    final_static_ref = static_ref_result['static_reference']
                    
                    # 7. 替換原公式
                    resolved_formula = resolved_formula.replace(index_func['full_function'], final_static_ref)
                    self.progress_callback.update_progress(f"[INDEX-SIMPLE] 替換: {index_func['full_function']} -> {final_static_ref}")
                    
                    static_references.append(final_static_ref)
                    calculation_details.append({
                        'original_function': index_func['full_function'],
                        'content': index_func['content'],
                        'static_reference': final_static_ref,
                        'array_param': array_param,
                        'row_value': row_value,
                        'col_value': col_value,
                        'build_details': static_ref_result
                    })
            
            return {
                'success': len(static_references) > 0,
                'resolved_formula': resolved_formula,
                'static_references': static_references,
                'calculation_details': calculation_details,
                'original_formula': formula,
                'internal_references': internal_references
            }
            
        except Exception as e:
            self.progress_callback.update_progress(f"[INDEX-SIMPLE] 解析異常: {e}")
            return {'success': False, 'error': str(e), 'original_formula': formula, 'internal_references': []}

    def _is_simple_number(self, param):
        """檢查是否為簡單數字"""
        try:
            float(param.strip())
            return True
        except:
            return False

    def _build_static_reference_from_index_simple(self, array_param, row_offset, col_offset, workbook_path, sheet_name):
        """簡化版靜態引用構建"""
        try:
            # 1. 解析範圍起始點
            array_info = self._parse_array_reference_debug(array_param, workbook_path, sheet_name)
            if not array_info['success']:
                return array_info
            
            start_cell = array_info['start_cell']  # 例如: U12
            
            # 2. 解析起始位置
            start_col, start_row = self._parse_cell_address_debug(start_cell)
            
            # 3. 計算最終位置 (Excel是1-based)
            final_row = start_row + row_offset - 1
            final_col = start_col + col_offset - 1
            
            final_cell = f"{self._col_num_to_letters(final_col)}{final_row}"
            
            # 4. 根據引用類型構建完整引用
            if array_info['type'] == 'external':
                static_ref = f"{array_info['prefix']}{final_cell}"
            elif array_info['type'] == 'local':
                static_ref = f"{array_info['target_sheet']}!{final_cell}"
            else:  # current
                static_ref = final_cell
            
            return {
                'success': True,
                'static_reference': static_ref,
                'array_info': array_info,
                'final_cell': final_cell,
                'calculated_position': {'row': final_row, 'col': final_col, 'col_letter': self._col_num_to_letters(final_col)}
            }
            
        except Exception as e:
            error_msg = f'靜態引用構建失敗: {str(e)}'
            return {'success': False, 'error': error_msg}
    
    def _extract_all_index_functions_debug(self, formula):
        """提取公式中所有的 INDEX 函數"""
        index_functions = []
        search_start = 0
        
        while True:
            index_pos = formula.upper().find('INDEX(', search_start)
            if index_pos == -1:
                break
            
            start_pos = index_pos + len('INDEX(')
            bracket_count = 1
            current_pos = start_pos
            
            while current_pos < len(formula) and bracket_count > 0:
                char = formula[current_pos]
                if char == '(':
                    bracket_count += 1
                elif char == ')':
                    bracket_count -= 1
                current_pos += 1
            
            if bracket_count == 0:
                content = formula[start_pos:current_pos-1]
                full_function = formula[index_pos:current_pos]
                
                index_functions.append({
                    'full_function': full_function,
                    'content': content,
                    'start_pos': index_pos,
                    'end_pos': current_pos
                })
            
            search_start = current_pos
        
        return index_functions
    
    def _extract_index_parameters_accurate_debug(self, content):
        """精確提取INDEX函數的三個參數"""
        try:
            params = []
            current_param = ""
            bracket_count = 0
            quote_count = 0
            in_quotes = False
            
            for char in content:
                if char == '"':
                    quote_count += 1
                    in_quotes = not in_quotes
                elif char == '(' and not in_quotes:
                    bracket_count += 1
                elif char == ')' and not in_quotes:
                    bracket_count -= 1
                elif char == ',' and bracket_count == 0 and not in_quotes:
                    params.append(current_param.strip())
                    current_param = ""
                    continue
                
                current_param += char
            
            # 添加最後一個參數
            if current_param.strip():
                params.append(current_param.strip())
            
            # INDEX函數至少需要2個參數，最多3個
            if len(params) < 2:
                return {'success': False, 'error': f'INDEX函數參數不足，需要至少2個參數，得到{len(params)}個'}
            
            # 如果只有2個參數，column默認為1
            if len(params) == 2:
                params.append('1')
            
            return {
                'success': True,
                'array': params[0],
                'row': params[1],
                'column': params[2] if len(params) > 2 else '1'
            }
            
        except Exception as e:
            return {'success': False, 'error': f'參數提取失敗: {str(e)}'}
    
    def _parse_array_reference_debug(self, array_param, workbook_path, sheet_name):
        """解析array參數，確定引用類型和起始位置"""
        try:
            original_param = array_param
            array_param = array_param.strip().strip('"').strip("'")
            
            # 檢查是否為常數數組
            if array_param.startswith('{') and array_param.endswith('}'):
                return {'success': False, 'error': 'INDEX暫不支持常數數組，請使用儲存格範圍'}
            
            # 解析不同類型的引用
            if '[' in array_param and ']' in array_param:
                # 外部文件引用：'C:\\Users\\user\\Desktop\\pytest\\[File.xlsx]Sheet1'!A1:Z100
                match = re.match(r"'?([^']*\[[^\]]+\][^']*)'?!(.+)", array_param)
                if match:
                    file_sheet_part = match.group(1)
                    range_part = match.group(2)
                    
                    start_cell = range_part.split(':')[0].replace('$', '') if ':' in range_part else range_part.replace('$', '')
                    
                    return {
                        'success': True,
                        'type': 'external',
                        'prefix': f"'{file_sheet_part}'!",
                        'range': range_part,
                        'start_cell': start_cell,
                        'target_sheet': file_sheet_part
                    }
            
            elif '!' in array_param:
                # 其他工作表引用：工作表2!A1:B100
                sheet_part, range_part = array_param.split('!', 1)
                sheet_part = sheet_part.strip("'")
                
                start_cell = range_part.split(':')[0].replace('$', '') if ':' in range_part else range_part.replace('$', '')
                
                return {
                    'success': True,
                    'type': 'local',
                    'prefix': f"{sheet_part}!",
                    'range': range_part,
                    'start_cell': start_cell,
                    'target_sheet': sheet_part
                }
            
            else:
                # 當前工作表引用：A1:Z100
                start_cell = array_param.split(':')[0].replace('$', '') if ':' in array_param else array_param.replace('$', '')
                
                return {
                    'success': True,
                    'type': 'current',
                    'prefix': '',
                    'range': array_param,
                    'start_cell': start_cell,
                    'target_sheet': sheet_name
                }
                
        except Exception as e:
            return {'success': False, 'error': f'Array參數解析失敗: {str(e)}'}
    
    def _parse_cell_address_debug(self, cell_address):
        """解析儲存格地址為列號和行號"""
        match = re.match(r'([A-Z]+)(\d+)', cell_address.upper())
        if not match:
            raise ValueError(f"Invalid cell address: {cell_address}")
        
        col_letters = match.group(1)
        row_num = int(match.group(2))
        
        col_num = 0
        for char in col_letters:
            col_num = col_num * 26 + (ord(char) - ord('A') + 1)
        
        return col_num, row_num

    def _col_num_to_letters(self, col_num):
        """將列號轉換為字母"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(ord('A') + (col_num % 26)) + result
            col_num //= 26
        return result