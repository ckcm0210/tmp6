# -*- coding: utf-8 -*-
"""
Pure INDIRECT Logic - 只提取你程式碼中的核心INDIRECT處理邏輯
不包含GUI，只有純邏輯
"""

import re
import os
import openpyxl
from urllib.parse import unquote
import win32com.client as win32

def resolve_indirect_pure(indirect_content, workbook_path, sheet_name, current_cell=None):
    """
    純INDIRECT解析邏輯 - 提取自你的unified_indirect_resolver
    
    Args:
        indirect_content: INDIRECT函數內容 (例如: D32&"!"&"A8")
        workbook_path: Excel文件路徑
        sheet_name: 工作表名稱
        current_cell: 當前儲存格 (例如: B32)
        
    Returns:
        str: 解析後的引用 (例如: 工作表2!A8)
    """
    try:
        print(f"🔍 [INDIRECT-CALC-1] 開始INDIRECT計算")
        print(f"    INDIRECT內容: {indirect_content}")
        print(f"    目標Excel文件: {workbook_path}")
        print(f"    目標工作表: {sheet_name}")
        print(f"    當前儲存格: {current_cell}")
        
        # 載入工作簿
        print(f"🔍 [INDIRECT-CALC-2] 正在載入Excel工作簿...")
        workbook = openpyxl.load_workbook(workbook_path, data_only=False)
        worksheet = workbook[sheet_name]
        print(f"✅ [INDIRECT-CALC-2] 成功載入工作簿和工作表")
        
        # 獲取外部連結映射
        print(f"🔍 [INDIRECT-CALC-3] 正在獲取外部連結映射...")
        external_links_map = get_external_links_map(workbook, workbook_path)
        print(f"✅ [INDIRECT-CALC-3] 外部連結映射: {external_links_map}")
        
        # 修復外部引用
        print(f"🔍 [INDIRECT-CALC-4] 正在修復外部引用...")
        print(f"    修復前: {indirect_content}")
        fixed_content = fix_external_references(indirect_content, external_links_map)
        print(f"    修復後: {fixed_content}")
        
        # 解析字串連接
        if '&' in fixed_content:
            print(f"🔍 [INDIRECT-CALC-5] 檢測到字串連接，開始解析...")
            result = resolve_concatenation(fixed_content, worksheet, current_cell, workbook_path)
            print(f"🔍 [INDIRECT-CALC-6] 字串連接解析完成")
            print(f"    最終靜態引用: {result}")
            
            # 驗證結果是否為有效的Excel引用
            if result and ('!' in result or re.match(r'^[A-Z]+\d+$', result)):
                print(f"✅ [INDIRECT-CALC-7] 成功生成有效的靜態引用: {result}")
                return result
            else:
                print(f"❌ [INDIRECT-CALC-7] 生成的引用無效: {result}")
                return None
        else:
            print(f"🔍 [INDIRECT-CALC-5] 簡單引用，無需字串連接")
            # 簡單引用，移除引號
            if fixed_content.startswith('"') and fixed_content.endswith('"'):
                fixed_content = fixed_content[1:-1]
            print(f"✅ [INDIRECT-CALC-6] 簡單引用結果: {fixed_content}")
            return fixed_content
            
    except Exception as e:
        print(f"Error in pure INDIRECT logic: {e}")
        return None

def get_external_links_map(workbook, workbook_path):
    """獲取外部連結映射 - 提取自你的邏輯"""
    external_links_map = {}
    
    try:
        if hasattr(workbook, '_external_links'):
            external_links = workbook._external_links
            if external_links:
                for i, link in enumerate(external_links, 1):
                    if hasattr(link, 'file_link') and link.file_link:
                        file_path = link.file_link.Target
                        if file_path:
                            decoded_path = unquote(file_path)
                            if decoded_path.startswith('file:///'):
                                decoded_path = decoded_path[8:]
                            elif decoded_path.startswith('file://'):
                                decoded_path = decoded_path[7:]
                            
                            external_links_map[str(i)] = decoded_path
        
        # 如果沒有找到，推斷常見的外部連結
        if not external_links_map:
            base_dir = os.path.dirname(workbook_path)
            common_files = [
                "Link1.xlsx", "Link2.xlsx", "Link3.xlsx",
                "File1.xlsx", "File2.xlsx", "File3.xlsx",
                "Data.xlsx", "GDP.xlsx", "Test.xlsx"
            ]
            
            index = 1
            for filename in common_files:
                full_path = os.path.join(base_dir, filename)
                if os.path.exists(full_path):
                    external_links_map[str(index)] = full_path
                    index += 1
                    
    except Exception as e:
        print(f"Error getting external links: {e}")
    
    return external_links_map

def fix_external_references(content, external_links_map):
    """修復外部引用 - 提取自你的邏輯"""
    try:
        def replace_ref(match):
            ref_num = match.group(1)
            if ref_num in external_links_map:
                full_path = external_links_map[ref_num]
                decoded_path = unquote(full_path) if isinstance(full_path, str) else full_path
                if decoded_path.startswith('file:///'):
                    decoded_path = decoded_path[8:]
                
                filename = os.path.basename(decoded_path)
                directory = os.path.dirname(decoded_path)
                return f"'[{directory}\\{filename}]'"
            return f"[Unknown_{ref_num}]"
        
        pattern = r'\[(\d+)\]'
        return re.sub(pattern, replace_ref, content)
    except:
        return content

def resolve_concatenation(content, worksheet, current_cell=None, workbook_path=None):
    """解析字串連接 - 提取自你的邏輯"""
    try:
        print(f"    🔍 [CONCAT-1] 開始解析字串連接")
        print(f"        輸入內容: {content}")
        print(f"        當前儲存格: {current_cell}")
        
        # 按 & 分割（智能處理引號內的&）
        parts = smart_split_by_ampersand(content)
        print(f"    🔍 [CONCAT-2] 分割結果: {parts}")
        
        result_parts = []
        for i, part in enumerate(parts, 1):
            part = part.strip()
            print(f"    🔍 [CONCAT-3-{i}] 正在處理第{i}個部分: {part}")
            
            # 字串常數
            if (part.startswith('"') and part.endswith('"')) or \
               (part.startswith("'") and part.endswith("'")):
                value = part[1:-1]
                result_parts.append(value)
                print(f"        ✅ 字串常數: '{value}'")
            
            # 儲存格引用
            elif re.match(r'^\$?[A-Z]+\$?\d+$', part):
                try:
                    print(f"        🔍 正在讀取儲存格 {part} 的值...")
                    cell_value = worksheet[part].value
                    result_parts.append(str(cell_value) if cell_value is not None else "")
                    print(f"        ✅ 儲存格 {part} 的值: {cell_value}")
                except Exception as e:
                    result_parts.append("")
                    print(f"        ❌ 讀取儲存格 {part} 失敗: {e}")
            
            # ROW()函數
            elif 'ROW()' in part.upper() and current_cell:
                try:
                    print(f"        🔍 正在處理ROW()函數...")
                    row_num = int(re.search(r'\d+', current_cell).group())
                    if '+' in part:
                        match = re.search(r'ROW\(\)\s*\+\s*(\d+)', part, re.IGNORECASE)
                        if match:
                            add_num = int(match.group(1))
                            result_parts.append(str(row_num + add_num))
                            print(f"        ✅ ROW()+{add_num}: {row_num}+{add_num}={row_num + add_num}")
                        else:
                            result_parts.append(str(row_num))
                            print(f"        ✅ ROW(): {row_num}")
                    else:
                        result_parts.append(str(row_num))
                        print(f"        ✅ ROW(): {row_num}")
                except Exception as e:
                    result_parts.append("ROW()")
                    print(f"        ❌ ROW()函數處理失敗: {e}")
            
            # COLUMN()函數
            elif 'COLUMN()' in part.upper() and current_cell:
                try:
                    print(f"        🔍 正在處理COLUMN()函數...")
                    col_letters = re.search(r'[A-Z]+', current_cell).group()
                    col_num = 0
                    for char in col_letters:
                        col_num = col_num * 26 + (ord(char) - ord('A') + 1)
                    result_parts.append(str(col_num))
                    print(f"        ✅ COLUMN(): {col_letters} -> {col_num}")
                except Exception as e:
                    result_parts.append("COLUMN()")
                    print(f"        ❌ COLUMN()函數處理失敗: {e}")
            
            # OFFSET()或其他複雜函數
            elif 'OFFSET(' in part.upper():
                print(f"        🔍 檢測到OFFSET函數，嘗試使用Excel計算: {part}")
                try:
                    # 使用Excel COM計算OFFSET函數，需要傳遞工作簿路徑
                    workbook_path = getattr(worksheet.parent, 'path', None)
                    excel_result = calculate_excel_function(part, worksheet, current_cell, workbook_path)
                    if excel_result:
                        result_parts.append(str(excel_result))
                        print(f"        ✅ Excel計算OFFSET結果: {part} -> {excel_result}")
                    else:
                        result_parts.append(part)
                        print(f"        ❌ Excel計算失敗，保持原樣: {part}")
                except Exception as e:
                    result_parts.append(part)
                    print(f"        ❌ Excel計算異常，保持原樣: {part} - {e}")
            
            else:
                # 其他，保持原樣
                result_parts.append(part)
                print(f"        📝 其他內容，保持原樣: {part}")
        
        print(f"    🔍 [CONCAT-4] 所有部分處理完成，正在組合結果...")
        print(f"        組合部分: {result_parts}")
        
        final_result = ''.join(result_parts)
        print(f"    ✅ [CONCAT-5] 字串連接最終結果: '{final_result}'")
        return final_result
        
    except Exception as e:
        print(f"Error in concatenation: {e}")
        return content

def calculate_excel_function(function_str, worksheet, current_cell=None, workbook_path=None):
    """使用Excel COM計算複雜函數（如OFFSET）"""
    xl = None
    excel_workbook = None
    
    try:
        print(f"            🔍 [EXCEL-CALC-1] 開始Excel COM計算")
        print(f"                函數: {function_str}")
        print(f"                當前儲存格: {current_cell}")
        print(f"                工作簿路徑: {workbook_path}")
        
        # 嘗試導入Excel COM
        try:
            import win32com.client as win32
        except ImportError:
            print(f"            ❌ [EXCEL-CALC-2] Excel COM不可用")
            return None
        
        # 連接Excel
        try:
            xl = win32.Dispatch("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            print(f"            ✅ [EXCEL-CALC-2] 創建Excel實例")
        except Exception as e:
            print(f"            ❌ [EXCEL-CALC-2] 無法連接Excel: {e}")
            return None
        
        # 獲取當前工作簿和工作表
        try:
            if workbook_path and os.path.exists(workbook_path):
                excel_workbook = xl.Workbooks.Open(workbook_path, UpdateLinks=0, ReadOnly=True)
                excel_ws = excel_workbook.Worksheets(worksheet.title)
                print(f"            ✅ [EXCEL-CALC-3] 打開工作簿: {workbook_path}")
                print(f"            ✅ [EXCEL-CALC-3] 獲取工作表: {worksheet.title}")
            else:
                print(f"            ❌ [EXCEL-CALC-3] 工作簿路徑無效: {workbook_path}")
                return None
        except Exception as e:
            print(f"            ❌ [EXCEL-CALC-3] 獲取工作表失敗: {e}")
            return None
        
        # 找一個空白儲存格進行計算
        test_cell = excel_ws.Range("ZZ999")  # 使用ZZ999作為測試儲存格
        
        # 保存原始狀態
        original_formula = test_cell.Formula
        original_calculation = xl.Calculation
        original_events = xl.EnableEvents
        original_screen_updating = xl.ScreenUpdating
        
        try:
            print(f"            🔍 [EXCEL-CALC-4] 設置保護模式並計算函數")
            
            # 設置保護模式
            xl.Calculation = -4135  # xlCalculationManual
            xl.EnableEvents = False
            xl.ScreenUpdating = False
            
            # 在Excel中計算函數
            test_formula = f"={function_str}"
            test_cell.Formula = test_formula
            test_cell.Calculate()
            
            # 獲取計算結果
            result_value = test_cell.Value
            print(f"            ✅ [EXCEL-CALC-5] Excel計算結果: {result_value} (類型: {type(result_value)})")
            
            # 處理結果
            if result_value is not None:
                # 如果是字串且看起來像儲存格地址
                if isinstance(result_value, str) and ('!' in result_value or re.match(r'^[A-Z]+\d+$', result_value)):
                    print(f"            ✅ [EXCEL-CALC-6] 識別為儲存格地址: {result_value}")
                    return result_value
                # 如果是數字，可能是儲存格的值，需要轉換為地址
                elif isinstance(result_value, (int, float)):
                    # 對於OFFSET函數，結果通常是儲存格的值，不是地址
                    # 但我們需要的是地址，所以嘗試獲取公式的地址
                    try:
                        # 嘗試獲取公式引用的地址
                        address_formula = f"=ADDRESS(ROW({function_str}),COLUMN({function_str}))"
                        test_cell.Formula = address_formula
                        test_cell.Calculate()
                        address_result = test_cell.Value
                        if address_result and isinstance(address_result, str):
                            print(f"            ✅ [EXCEL-CALC-6] 轉換為地址: {address_result}")
                            return address_result
                    except:
                        pass
                    
                    # 如果無法獲取地址，返回值本身
                    print(f"            ⚠️ [EXCEL-CALC-6] 返回計算值: {result_value}")
                    return str(result_value)
                else:
                    print(f"            ⚠️ [EXCEL-CALC-6] 未知結果類型，轉為字串: {result_value}")
                    return str(result_value)
            
            print(f"            ❌ [EXCEL-CALC-6] 計算結果為None")
            return None
            
        finally:
            # 恢復所有狀態
            try:
                test_cell.Formula = original_formula
                xl.Calculation = original_calculation
                xl.EnableEvents = original_events
                xl.ScreenUpdating = original_screen_updating
                print(f"            ✅ [EXCEL-CALC-7] 恢復Excel狀態")
            except Exception as e:
                print(f"            ⚠️ [EXCEL-CALC-7] 恢復狀態時出錯: {e}")
            
            # 關閉工作簿
            try:
                if excel_workbook:
                    excel_workbook.Close(SaveChanges=False)
                    print(f"            ✅ [EXCEL-CALC-8] 關閉工作簿")
            except Exception as e:
                print(f"            ⚠️ [EXCEL-CALC-8] 關閉工作簿時出錯: {e}")
            
            # 退出Excel
            try:
                if xl:
                    xl.Quit()
                    print(f"            ✅ [EXCEL-CALC-9] 退出Excel")
            except Exception as e:
                print(f"            ⚠️ [EXCEL-CALC-9] 退出Excel時出錯: {e}")
            
    except Exception as e:
        print(f"            ❌ [EXCEL-CALC-ERROR] Excel計算錯誤: {e}")
        
        # 緊急清理
        try:
            if excel_workbook:
                excel_workbook.Close(SaveChanges=False)
            if xl:
                xl.Quit()
        except:
            pass
            
        return None

def smart_split_by_ampersand(content):
    """按 & 分割，但不會分割引號內的 & - 提取自你的邏輯"""
    try:
        parts = []
        current_part = ""
        in_quotes = False
        quote_char = None
        
        i = 0
        while i < len(content):
            char = content[i]
            
            # 處理引號
            if char in ['"', "'"] and not in_quotes:
                in_quotes = True
                quote_char = char
                current_part += char
            elif char == quote_char and in_quotes:
                in_quotes = False
                quote_char = None
                current_part += char
            elif char == '&' and not in_quotes:
                # 分割點
                if current_part.strip():
                    parts.append(current_part.strip())
                current_part = ""
            else:
                current_part += char
            
            i += 1
        
        # 加最後一部分
        if current_part.strip():
            parts.append(current_part.strip())
        
        return parts
    except Exception as e:
        print(f"Error in smart split: {e}")
        return [content]

def process_formula_with_pure_indirect(formula, workbook_path, sheet_name, current_cell=None):
    """
    使用純邏輯處理包含INDIRECT的公式
    
    Args:
        formula: 公式字串 (例如: =INDIRECT(D32&"!"&"A8"))
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
        
        print(f"Processing formula with pure INDIRECT logic: {formula}")
        
        # === 修復：正確提取INDIRECT函數內容，處理嵌套括號 ===
        def extract_indirect_content_fixed(formula):
            """正確提取INDIRECT內容，處理嵌套括號"""
            print(f"│  • [EXTRACT-1] - 開始提取")
            print(f"│    輸入公式: {formula}")
            
            formula_upper = formula.upper()
            indirect_pos = formula_upper.find('INDIRECT(')
            if indirect_pos == -1:
                print(f"│  • [EXTRACT-2] - INDIRECT位置檢查: 找不到INDIRECT(")
                return None
                
            start_pos = indirect_pos + len('INDIRECT(')
            bracket_count = 1
            current_pos = start_pos
            
            print(f"│  • [EXTRACT-2] - INDIRECT位置檢查: INDIRECT(位置 {indirect_pos}, 內容開始位置 {start_pos}")
            print(f"│  • [EXTRACT-3] - 括號匹配過程（逐字符顯示）")
            print(f"開始提取INDIRECT內容，起始位置: {start_pos}")
            
            while current_pos < len(formula) and bracket_count > 0:
                char = formula[current_pos]
                print(f"    位置 {current_pos}: '{char}', bracket_count: {bracket_count}")
                
                if char == '(':
                    bracket_count += 1
                elif char == ')':
                    bracket_count -= 1
                current_pos += 1
            
            print(f"│  • [EXTRACT-4] - 括號匹配結果")
            print(f"│    最終位置: {current_pos}, bracket_count: {bracket_count}")
            
            if bracket_count == 0:
                content = formula[start_pos:current_pos-1]
                print(f"│  • [EXTRACT-5] - 最終提取結果")
                print(f"成功提取INDIRECT內容: '{content}'")
                return content
            else:
                print(f"│  • [EXTRACT-5] - 最終提取結果")
                print(f"括號不匹配，bracket_count: {bracket_count}")
                return None
        
        # 使用修復後的提取函數
        indirect_content = extract_indirect_content_fixed(formula)
        if not indirect_content:
            return {
                'has_indirect': True,
                'original_formula': formula,
                'resolved_formula': formula,
                'success': False,
                'error': 'Could not extract INDIRECT content'
            }
        
        print(f"INDIRECT content: {indirect_content}")
        
        print(f"🔍 [FINAL-1] 開始使用純邏輯解析INDIRECT...")
        # 使用純邏輯解析
        resolved_ref = resolve_indirect_pure(indirect_content, workbook_path, sheet_name, current_cell)
        
        print(f"🔍 [FINAL-2] 純邏輯解析完成")
        print(f"    解析結果: {resolved_ref}")
        
        if resolved_ref:
            print(f"🔍 [FINAL-3] 開始替換原始公式中的INDIRECT函數...")
            # 替換INDIRECT函數為解析後的引用
            indirect_function = f"INDIRECT({indirect_content})"
            resolved_formula = formula.replace(indirect_function, resolved_ref)
            
            print(f"✅ [FINAL-4] INDIRECT函數替換成功!")
            print(f"    原始公式: {formula}")
            print(f"    INDIRECT函數: {indirect_function}")
            print(f"    靜態引用: {resolved_ref}")
            print(f"    最終公式: {resolved_formula}")
            
            return {
                'has_indirect': True,
                'original_formula': formula,
                'resolved_formula': resolved_formula,
                'success': True,
                'error': None
            }
        else:
            print(f"❌ [FINAL-3] INDIRECT解析失敗")
            print(f"    原因: resolve_indirect_pure返回None")
            return {
                'has_indirect': True,
                'original_formula': formula,
                'resolved_formula': formula,
                'success': False,
                'error': 'INDIRECT resolution failed - resolve_indirect_pure returned None'
            }
        
    except Exception as e:
        print(f"Error processing formula: {e}")
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
    test_workbook = r'C:\Users\user\Excel_tools_develop\Excel_tools_develop_v70\File5_v4.xlsx'
    test_sheet = "工作表1"
    test_cell = "B32"
    
    print("=== 測試純INDIRECT邏輯 ===")
    
    try:
        result = process_formula_with_pure_indirect(test_formula, test_workbook, test_sheet, test_cell)
        
        print(f"測試結果:")
        print(f"  Has INDIRECT: {result['has_indirect']}")
        print(f"  Success: {result['success']}")
        print(f"  Original: {result['original_formula']}")
        print(f"  Resolved: {result['resolved_formula']}")
        print(f"  Error: {result['error']}")
        
        if result['success'] and '!A8' in result['resolved_formula']:
            print("🎉 純INDIRECT邏輯工作正常！")
        else:
            print("❌ 純INDIRECT邏輯需要調整")
            
    except Exception as e:
        print(f"測試失敗: {e}")
    
    input("按Enter退出...")