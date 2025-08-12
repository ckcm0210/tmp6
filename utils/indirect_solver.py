# -*- coding: utf-8 -*-
"""
INDIRECT Solver - 從 progress_enhanced_exploder.py 中提取的INDIRECT解析邏輯
純粹的程式碼搬移，不修改任何邏輯
"""

class IndirectSolver:
    """INDIRECT函數解析器 - 從原始程式碼中完全搬移"""
    
    def __init__(self, excel_manager, progress_callback, main_analyzer=None):
        self.excel_manager = excel_manager
        self.progress_callback = progress_callback
        self.main_analyzer = main_analyzer
    
    def _resolve_indirect_with_excel(self, formula, workbook_path, sheet_name, cell_address):
        """使用安全的 Excel 管理解析 INDIRECT"""
        try:
            self.progress_callback.update_progress(f"[INDIRECT] 開始解析: {formula}")
            
            # 提取所有 INDIRECT 函數
            indirect_functions = self._extract_all_indirect_functions(formula)
            if not indirect_functions:
                return {'success': False, 'error': 'No INDIRECT functions found'}
            
            resolved_formula = formula
            static_references = []
            calculation_details = []
            internal_references = []
            
            # 分析 INDIRECT 內部引用
            for indirect_func in indirect_functions:
                if self.main_analyzer:
                    temp_references = self.main_analyzer._parse_formula_references_accurate(
                        f"={indirect_func['content']}", workbook_path, sheet_name
                    )
                    internal_references.extend(temp_references)
            
            # 逐個解析 INDIRECT 函數
            for i, indirect_func in enumerate(indirect_functions):
                self.progress_callback.update_progress(f"[INDIRECT] 處理第 {i+1} 個: {indirect_func['content']}")
                
                calc_result = self.excel_manager.calculate_safely(
                    indirect_func['content'], workbook_path, sheet_name, cell_address
                )
                
                if calc_result and calc_result['success']:
                    static_ref = calc_result['static_reference']
                    
                    if '!' in static_ref:
                        final_static_ref = static_ref
                    else:
                        final_static_ref = f"{sheet_name}!{static_ref}"
                    
                    old_formula = resolved_formula
                    resolved_formula = resolved_formula.replace(
                        indirect_func['full_function'], 
                        final_static_ref
                    )
                    
                    self.progress_callback.update_progress(f"[INDIRECT] 替換: {indirect_func['full_function']} -> {final_static_ref}")
                    
                    static_references.append(final_static_ref)
                    calculation_details.append({
                        'original_function': indirect_func['full_function'],
                        'content': indirect_func['content'],
                        'static_reference': final_static_ref,
                        'raw_excel_result': static_ref
                    })
            
            success = len(static_references) > 0
            
            return {
                'success': success,
                'resolved_formula': resolved_formula,
                'static_references': static_references,
                'calculation_details': calculation_details,
                'original_formula': formula,
                'internal_references': internal_references
            }
            
        except Exception as e:
            self.progress_callback.update_progress(f"[INDIRECT] 解析異常: {e}")
            return {
                'success': False,
                'error': str(e),
                'original_formula': formula,
                'internal_references': []
            }

    def _extract_all_indirect_functions(self, formula):
        """提取公式中所有的 INDIRECT 函數"""
        indirect_functions = []
        search_start = 0
        
        while True:
            indirect_pos = formula.upper().find('INDIRECT(', search_start)
            if indirect_pos == -1:
                break
            
            start_pos = indirect_pos + len('INDIRECT(')
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
                full_function = formula[indirect_pos:current_pos]
                
                indirect_functions.append({
                    'full_function': full_function,
                    'content': content,
                    'start_pos': indirect_pos,
                    'end_pos': current_pos
                })
            
            search_start = current_pos
        
        return indirect_functions