# -*- coding: utf-8 -*-
"""
Core INDIRECT Resolver - 直接使用你的unified_indirect_resolver核心邏輯
移植核心邏輯，去掉GUI，專注於INDIRECT解析
"""

import sys
import os
import traceback

def resolve_indirect_core(formula, workbook_path, sheet_name, current_cell=None):
    """
    核心INDIRECT解析函數 - 直接使用你的unified_indirect_resolver邏輯
    
    Args:
        formula: 包含INDIRECT的公式 (例如: =INDIRECT(D32&"!"&"A8"))
        workbook_path: Excel文件路徑
        sheet_name: 工作表名稱
        current_cell: 當前儲存格地址 (例如: B32)
        
    Returns:
        dict: {
            'success': bool,
            'original_formula': str,
            'resolved_formula': str,
            'error': str or None
        }
    """
    try:
        # 檢查是否包含INDIRECT
        if not formula or 'INDIRECT' not in formula.upper():
            return {
                'success': False,
                'original_formula': formula,
                'resolved_formula': formula,
                'error': 'No INDIRECT function found'
            }
        
        print(f"Core INDIRECT resolver starting...")
        print(f"Formula: {formula}")
        print(f"Workbook: {workbook_path}")
        print(f"Sheet: {sheet_name}")
        print(f"Current cell: {current_cell}")
        
        # === 直接使用你的unified_indirect_resolver ===
        # 添加indirect_tool路徑
        indirect_tool_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'indirect_tool')
        if indirect_tool_path not in sys.path:
            sys.path.append(indirect_tool_path)
        
        from unified_indirect_resolver import UnifiedIndirectResolver
        
        # 創建resolver實例
        resolver = UnifiedIndirectResolver()
        
        # 設置必要的屬性
        resolver.workbook_path = workbook_path
        resolver.sheet_name = sheet_name
        resolver.current_cell = current_cell
        resolver.formula = formula
        
        print("Calling resolve_indirect_unified()...")
        
        # 執行解析 - 使用你的核心邏輯
        result = resolver.resolve_indirect_unified()
        
        print(f"Resolver result: {result}")
        
        if result and isinstance(result, dict):
            # 成功解析
            resolved_formula = result.get('resolved_formula', formula)
            
            return {
                'success': True,
                'original_formula': formula,
                'resolved_formula': resolved_formula,
                'error': None
            }
        else:
            # 解析失敗
            return {
                'success': False,
                'original_formula': formula,
                'resolved_formula': formula,
                'error': 'Resolver returned no result'
            }
        
    except Exception as e:
        print(f"Error in core INDIRECT resolver: {e}")
        import traceback
        traceback.print_exc()
        
        return {
            'success': False,
            'original_formula': formula,
            'resolved_formula': formula,
            'error': str(e)
        }


def process_formula_with_indirect(formula, workbook_path, sheet_name, current_cell=None):
    """
    處理包含INDIRECT的公式 - 便捷函數
    
    Args:
        formula: 公式字串
        workbook_path: Excel文件路徑
        sheet_name: 工作表名稱
        current_cell: 當前儲存格地址
        
    Returns:
        dict: {
            'has_indirect': bool,
            'original_formula': str,
            'resolved_formula': str,
            'success': bool,
            'error': str or None
        }
    """
    try:
        # 檢查是否包含INDIRECT
        if not formula or 'INDIRECT' not in formula.upper():
            return {
                'has_indirect': False,
                'original_formula': formula,
                'resolved_formula': formula,
                'success': True,
                'error': None
            }
        
        # 使用核心解析器
        result = resolve_indirect_core(formula, workbook_path, sheet_name, current_cell)
        
        return {
            'has_indirect': True,
            'original_formula': result['original_formula'],
            'resolved_formula': result['resolved_formula'],
            'success': result['success'],
            'error': result['error']
        }
        
    except Exception as e:
        return {
            'has_indirect': True,
            'original_formula': formula,
            'resolved_formula': formula,
            'success': False,
            'error': str(e)
        }


# 測試函數
if __name__ == "__main__":
    # 測試用例
    test_formula = '=INDIRECT(D32&"!"&"A8")'
    test_workbook = r"C:\Users\user\Excel_tools_develop\Excel_tools_develop_v70\File5_v4.xlsx"
    test_sheet = "工作表1"
    test_cell = "B32"
    
    print("=== 測試核心INDIRECT解析器 ===")
    
    try:
        result = process_formula_with_indirect(test_formula, test_workbook, test_sheet, test_cell)
        
        print(f"測試結果:")
        print(f"  Has INDIRECT: {result['has_indirect']}")
        print(f"  Success: {result['success']}")
        print(f"  Original: {result['original_formula']}")
        print(f"  Resolved: {result['resolved_formula']}")
        print(f"  Error: {result['error']}")
        
        if result['success'] and result['resolved_formula'] != result['original_formula']:
            print("🎉 核心INDIRECT解析器工作正常！")
        else:
            print("❌ 核心INDIRECT解析器需要調整")
            
    except Exception as e:
        print(f"測試失敗: {e}")
    
    input("按Enter退出...")