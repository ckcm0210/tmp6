# -*- coding: utf-8 -*-
"""
VLOOKUP Solver - 解析並靜態化 VLOOKUP 函數（只支援精確匹配 FALSE）
設計與 IndexSolver/IndirectSolver 一致：
- 依賴 ExcelComManager 進行必要的計算（MATCH、複雜參數）
- 可選擇從 main_analyzer 取得內部引用以便圖譜顯示
"""

import re

class VLookupSolver:
    """VLOOKUP 函數解析器"""

    def __init__(self, excel_manager, progress_callback, main_analyzer=None):
        self.excel_manager = excel_manager
        self.progress_callback = progress_callback
        self.main_analyzer = main_analyzer

    def resolve_vlookup(self, formula, workbook_path, sheet_name, cell_address):
        """
        將公式中的 VLOOKUP 函數轉換為靜態引用（只支援第四參數為 FALSE 的情況）。

        Returns dict:
            {
                'success': bool,
                'resolved_formula': str,
                'static_references': list[str],
                'calculation_details': list[dict],
                'original_formula': str,
                'internal_references': list[dict],
                'errors': list[str]
            }
        """
        try:
            self.progress_callback.update_progress(f"[VLOOKUP] 開始解析: {formula}")

            vlookups = self._extract_all_vlookup_functions(formula)
            if not vlookups:
                return {'success': False, 'error': 'No VLOOKUP functions found'}

            resolved_formula = formula
            static_references = []
            calculation_details = []
            internal_references = []
            errors = []

            for i, vinfo in enumerate(vlookups):
                full_fn = vinfo['full_function']
                content = vinfo['content']
                self.progress_callback.update_progress(f"[VLOOKUP] 處理第 {i+1} 個: {content}")

                params_res = self._extract_vlookup_parameters(content)
                if not params_res['success']:
                    errors.append(params_res['error'])
                    continue

                lookup_param = params_res['lookup_value']
                table_param = params_res['table_array']
                col_index_param = params_res['col_index']
                range_lookup_param = params_res['range_lookup']

                # 僅支援 FALSE（精確匹配）
                try:
                    if not self._is_param_false(range_lookup_param, workbook_path, sheet_name, cell_address):
                        msg = f"第四參數只支援 FALSE，實際為: {range_lookup_param}"
                        self.progress_callback.update_progress(f"[VLOOKUP] {msg}")
                        errors.append(msg)
                        continue
                except Exception as e:
                    errors.append(f"檢查第四參數失敗: {e}")
                    continue

                # 解析 table_array 的類型與起始位置
                array_info = self._parse_array_reference_debug(table_param, workbook_path, sheet_name)
                if not array_info['success']:
                    errors.append(array_info.get('error', '表陣列參數解析失敗'))
                    continue

                # 內部引用（提供給圖譜）
                if self.main_analyzer:
                    try:
                        refs1 = self.main_analyzer._parse_formula_references_accurate(f"={lookup_param}", workbook_path, sheet_name)
                        refs2 = self.main_analyzer._parse_formula_references_accurate(f"={table_param}", workbook_path, sheet_name)
                        internal_references.extend(refs1)
                        internal_references.extend(refs2)
                    except Exception:
                        pass

                # 解析列索引
                try:
                    col_index = self._resolve_to_integer(col_index_param, workbook_path, sheet_name, cell_address)
                    if col_index < 1:
                        errors.append(f"列索引無效: {col_index_param}")
                        continue
                except Exception as e:
                    errors.append(f"列索引解析失敗: {e}")
                    continue

                # 構建第一列搜尋範圍（精確匹配）
                try:
                    first_col_letter = self._col_letters_of_cell(array_info['start_cell'])
                    min_row = self._row_of_cell(array_info['start_cell'])
                    max_row = self._max_row_from_range(array_info['range'])

                    search_range = f"{first_col_letter}{min_row}:{first_col_letter}{max_row}"
                    if array_info['type'] == 'external':
                        search_range = f"{array_info['prefix']}{search_range}"
                    elif array_info['type'] == 'local':
                        search_range = f"{array_info['prefix']}{search_range}"
                    # current: 無前綴
                except Exception as e:
                    errors.append(f"無法構建搜尋範圍: {e}")
                    continue

                # 使用 Excel 計算 MATCH 以獲得行偏移（1-based）
                try:
                    match_content = f"MATCH({lookup_param}, {search_range}, 0)"
                    mres = self.excel_manager.calculate_safely(match_content, workbook_path, sheet_name, cell_address)
                    if not mres['success']:
                        errors.append(f"MATCH 計算失敗: {mres.get('error')}")
                        continue
                    row_offset = int(float(str(mres['static_reference']).strip()))
                except Exception as e:
                    errors.append(f"MATCH 解析行偏移失敗: {e}")
                    continue

                # 計算最終行列
                try:
                    start_col_num = self._col_num_of_letters(first_col_letter)
                    final_col_num = start_col_num + col_index - 1
                    final_col_letters = self._col_num_to_letters(final_col_num)
                    final_row = min_row + row_offset - 1
                except Exception as e:
                    errors.append(f"計算最終位置失敗: {e}")
                    continue

                # 構建靜態引用（含前綴）
                try:
                    if array_info['type'] == 'external':
                        static_ref = f"{array_info['prefix']}{final_col_letters}{final_row}"
                    elif array_info['type'] == 'local':
                        static_ref = f"{array_info['prefix']}{final_col_letters}{final_row}"
                    else:
                        static_ref = f"{final_col_letters}{final_row}"
                except Exception as e:
                    errors.append(f"構建靜態引用失敗: {e}")
                    continue

                # 替換原公式片段
                try:
                    resolved_formula = resolved_formula.replace(full_fn, static_ref)
                except Exception:
                    pass

                static_references.append(static_ref)
                calculation_details.append({
                    'original_function': full_fn,
                    'content': content,
                    'lookup_param': lookup_param,
                    'table_param': table_param,
                    'col_index_param': col_index_param,
                    'range_lookup_param': range_lookup_param,
                    'search_range': search_range,
                    'row_offset': row_offset,
                    'final_ref': static_ref
                })

            success = len(static_references) > 0
            return {
                'success': success,
                'resolved_formula': resolved_formula,
                'static_references': static_references,
                'calculation_details': calculation_details,
                'original_formula': formula,
                'internal_references': internal_references,
                'errors': errors
            }

        except Exception as e:
            self.progress_callback.update_progress(f"[VLOOKUP] 解析異常: {e}")
            return {'success': False, 'error': str(e), 'original_formula': formula, 'internal_references': [], 'errors': [str(e)]}

    # ---- helpers ----

    def _extract_all_vlookup_functions(self, formula):
        """提取所有 VLOOKUP(...) 片段，返回 list[{full_function, content}]"""
        items = []
        search_start = 0
        up = formula.upper()
        while True:
            pos = up.find('VLOOKUP(', search_start)
            if pos == -1:
                break
            start_pos = pos + len('VLOOKUP(')
            bracket = 1
            i = start_pos
            in_quotes = False
            while i < len(formula) and bracket > 0:
                ch = formula[i]
                if ch == '"':
                    in_quotes = not in_quotes
                elif not in_quotes:
                    if ch == '(':
                        bracket += 1
                    elif ch == ')':
                        bracket -= 1
                i += 1
            if bracket == 0:
                content = formula[start_pos:i-1]
                full_function = formula[pos:i]
                items.append({'full_function': full_function, 'content': content})
            search_start = i
        return items

    def _extract_vlookup_parameters(self, content):
        """健壯地分割四個參數（處理括號與引號）"""
        try:
            params = []
            cur = ''
            bracket = 0
            in_quotes = False
            for ch in content:
                if ch == '"':
                    in_quotes = not in_quotes
                elif ch == '(' and not in_quotes:
                    bracket += 1
                elif ch == ')' and not in_quotes:
                    bracket -= 1
                elif ch == ',' and bracket == 0 and not in_quotes:
                    params.append(cur.strip())
                    cur = ''
                    continue
                cur += ch
            if cur.strip():
                params.append(cur.strip())
            if len(params) < 3:
                return {'success': False, 'error': f'VLOOKUP 參數不足，得到 {len(params)} 個'}
            if len(params) == 3:
                params.append('FALSE')  # 預設為精確匹配
            return {
                'success': True,
                'lookup_value': params[0],
                'table_array': params[1],
                'col_index': params[2],
                'range_lookup': params[3]
            }
        except Exception as e:
            return {'success': False, 'error': f'參數解析失敗: {e}'}

    def _is_param_false(self, param, workbook_path, sheet_name, cell_address):
        """確認第四參數為 FALSE（字面或計算後）"""
        if isinstance(param, str) and param.strip().upper() in ('FALSE', '0'):
            return True
        # 複雜情況交給 Excel 計算
        res = self.excel_manager.calculate_safely(param, workbook_path, sheet_name, cell_address)
        if not res['success']:
            return False
        val = str(res['static_reference']).strip().upper()
        return val in ('FALSE', '0')

    def _resolve_to_integer(self, param, workbook_path, sheet_name, cell_address):
        p = param.strip()
        # 直接數字
        try:
            return int(float(p))
        except:
            pass
        # 需要 Excel 計算
        cres = self.excel_manager.calculate_safely(p, workbook_path, sheet_name, cell_address)
        if not cres['success']:
            raise ValueError(f"參數計算失敗: {cres.get('error')}")
        return int(float(str(cres['static_reference']).strip()))

    def _parse_array_reference_debug(self, array_param, workbook_path, sheet_name):
        """與 IndexSolver 的解析邏輯對齊，解析 table_array 類型與起始位置"""
        try:
            original_param = array_param
            array_param = array_param.strip().strip('"').strip("'")

            # 外部文件引用
            if '[' in array_param and ']' in array_param:
                m = re.match(r"'?([^']*\[[^\]]+\][^']*)'?!(.+)", array_param)
                if m:
                    file_sheet_part = m.group(1)
                    range_part = m.group(2)
                    start_cell = range_part.split(':')[0].replace('$', '') if ':' in range_part else range_part.replace('$', '')
                    return {'success': True, 'type': 'external', 'prefix': f"'{file_sheet_part}'!", 'range': range_part, 'start_cell': start_cell, 'target_sheet': file_sheet_part}

            # 其他工作表引用
            if '!' in array_param:
                sheet_part, range_part = array_param.split('!', 1)
                sheet_part = sheet_part.strip("'")
                start_cell = range_part.split(':')[0].replace('$', '') if ':' in range_part else range_part.replace('$', '')
                return {'success': True, 'type': 'local', 'prefix': f"{sheet_part}!", 'range': range_part, 'start_cell': start_cell, 'target_sheet': sheet_part}

            # 當前工作表引用
            start_cell = array_param.split(':')[0].replace('$', '') if ':' in array_param else array_param.replace('$', '')
            return {'success': True, 'type': 'current', 'prefix': '', 'range': array_param, 'start_cell': start_cell, 'target_sheet': sheet_name}
        except Exception as e:
            return {'success': False, 'error': f'表陣列參數解析失敗: {e}'}

    def _parse_cell_address_debug(self, cell_address):
        m = re.match(r'([A-Z]+)(\d+)', cell_address.upper())
        if not m:
            raise ValueError(f"Invalid cell address: {cell_address}")
        col_letters = m.group(1)
        row_num = int(m.group(2))
        col_num = 0
        for ch in col_letters:
            col_num = col_num * 26 + (ord(ch) - ord('A') + 1)
        return col_num, row_num

    def _col_num_to_letters(self, col_num):
        res = ''
        while col_num > 0:
            col_num -= 1
            res = chr(ord('A') + (col_num % 26)) + res
            col_num //= 26
        return res

    def _col_num_of_letters(self, letters):
        num = 0
        for ch in letters.upper():
            num = num * 26 + (ord(ch) - ord('A') + 1)
        return num

    def _col_letters_of_cell(self, cell):
        m = re.match(r'([A-Z]+)(\d+)', cell.upper())
        if not m:
            raise ValueError(f"Invalid cell: {cell}")
        return m.group(1)

    def _row_of_cell(self, cell):
        m = re.match(r'([A-Z]+)(\d+)', cell.upper())
        if not m:
            raise ValueError(f"Invalid cell: {cell}")
        return int(m.group(2))

    def _max_row_from_range(self, range_part):
        # range_part 例如 A5:S19 或 A5
        clean = range_part.replace('$', '')
        if ':' not in clean:
            return self._row_of_cell(clean)
        _, end_cell = clean.split(':', 1)
        return self._row_of_cell(end_cell)
