# -*- coding: utf-8 -*-
"""
Enhanced Dependency Exploder - 公式依賴鏈遞歸分析器 (增強版)
支援 INDIRECT 函數動態解析
"""

import re
import os
import win32com.client
from urllib.parse import unquote
from utils.openpyxl_resolver import read_cell_with_resolved_references
import traceback
import hashlib

class DependencyExploder:
    """增強版公式依賴鏈爆炸分析器 - 支援 INDIRECT 函數解析"""
    
    def __init__(self, max_depth=10, range_expand_threshold=5, enable_indirect_resolution=True):
        self.max_depth = max_depth
        self.range_expand_threshold = range_expand_threshold
        self.enable_indirect_resolution = enable_indirect_resolution
        self.visited_cells = set()
        self.circular_refs = []
        self.excel_app = None  # 復用 Excel 實例
        self.indirect_resolution_log = []  # INDIRECT 解析日誌
    
    def __del__(self):
        """析構函數：確保 Excel 進程正確關閉"""
        self._cleanup_excel()
    
    def _cleanup_excel(self):
        """清理 Excel 進程"""
        try:
            if self.excel_app:
                self.excel_app.Quit()
                self.excel_app = None
        except:
            pass
    
    def explode_dependencies(self, workbook_path, sheet_name, cell_address, current_depth=0, root_workbook_path=None):
        """
        遞歸展開公式依賴鏈 (增強版 - 支援 INDIRECT)
        
        Args:
            workbook_path: Excel 檔案路徑
            sheet_name: 工作表名稱
            cell_address: 儲存格地址 (如 A1)
            current_depth: 當前遞歸深度
            root_workbook_path: 根工作簿路徑
            
        Returns:
            dict: 依賴樹結構
        """
        # 創建唯一標識符
        cell_id = f"{workbook_path}|{sheet_name}|{cell_address}"
        
        # 檢查遞歸深度限制
        if current_depth >= self.max_depth:
            return self._create_limit_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
        
        # 檢查循環引用
        if cell_id in self.visited_cells:
            self.circular_refs.append(cell_id)
            return self._create_circular_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
        
        # 標記為已訪問
        self.visited_cells.add(cell_id)
        
        try:
            # 讀取儲存格內容
            cell_info = read_cell_with_resolved_references(workbook_path, sheet_name, cell_address)
            
            if 'error' in cell_info:
                return self._create_error_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, cell_info['error'])
            
            # 基本節點信息
            original_formula = cell_info.get('formula')
            processed_formula = original_formula
            indirect_resolved = False
            indirect_details = None
            
            # 增強的公式清理
            if original_formula:
                processed_formula = self._clean_formula(original_formula)
            
            # *** 新增：INDIRECT 檢測和處理 ***
            if (self.enable_indirect_resolution and 
                cell_info.get('cell_type') == 'formula' and 
                processed_formula and 
                self._contains_indirect(processed_formula)):
                
                print(f"[INDIRECT] ===============================================")
                print(f"[INDIRECT] 🔍 檢測到 INDIRECT 函數！")
                print(f"[INDIRECT] 📍 位置: {sheet_name}!{cell_address}")
                print(f"[INDIRECT] 📂 文件: {workbook_path}")
                print(f"[INDIRECT] 📝 公式: {processed_formula}")
                print(f"[INDIRECT] ===============================================")
                
                try:
                    print(f"[INDIRECT] 🚀 開始解析 INDIRECT...")
                    resolved_result = self._resolve_indirect_formula(
                        processed_formula, workbook_path, sheet_name, cell_address
                    )
                    
                    print(f"[INDIRECT] 📊 解析結果摘要:")
                    print(f"[INDIRECT]   成功: {resolved_result.get('success', False)}")
                    if resolved_result.get('success'):
                        print(f"[INDIRECT]   靜態引用: {resolved_result.get('static_references', [])}")
                    else:
                        print(f"[INDIRECT]   錯誤: {resolved_result.get('error', 'Unknown')}")
                    
                    if resolved_result and resolved_result['success']:
                        old_formula = processed_formula
                        processed_formula = resolved_result['resolved_formula']
                        indirect_resolved = True
                        indirect_details = resolved_result
                        
                        print(f"[INDIRECT] ✅ INDIRECT 解析成功！")
                        print(f"[INDIRECT] 📝 原始公式: {old_formula}")
                        print(f"[INDIRECT] 🎯 解析後: {processed_formula}")
                        print(f"[INDIRECT] 📋 靜態引用: {resolved_result.get('static_references', [])}")
                        
                        # 記錄解析日誌
                        self.indirect_resolution_log.append({
                            'cell': f"{sheet_name}!{cell_address}",
                            'original': original_formula,
                            'resolved': processed_formula,
                            'details': resolved_result
                        })
                    else:
                        print(f"[INDIRECT] ❌ INDIRECT 解析失敗")
                        print(f"[INDIRECT] 🚫 錯誤: {resolved_result.get('error', 'Unknown error')}")
                        
                except Exception as e:
                    print(f"[INDIRECT] ❌❌❌ INDIRECT 處理發生異常 ❌❌❌")
                    print(f"[INDIRECT] 異常: {e}")
                    import traceback
                    for line in traceback.format_exc().split('\n'):
                        print(f"[INDIRECT]   {line}")
                    
                    # INDIRECT 解析失敗，記錄錯誤但繼續處理
                    self.indirect_resolution_log.append({
                        'cell': f"{sheet_name}!{cell_address}",
                        'original': original_formula,
                        'error': str(e),
                        'resolved': False
                    })
            
            # 構建節點
            node = self._create_base_node(
                workbook_path, sheet_name, cell_address, current_depth, 
                root_workbook_path, cell_info, original_formula, processed_formula,
                indirect_resolved, indirect_details
            )
            
            # 如果是公式，解析依賴關係（使用處理後的公式）
            if cell_info.get('cell_type') == 'formula' and processed_formula:
                references = self.parse_formula_references(processed_formula, workbook_path, sheet_name)
                
                # 遞歸展開每個引用
                for ref in references:
                    try:
                        child_node = self.explode_dependencies(
                            ref['workbook_path'],
                            ref['sheet_name'],
                            ref['cell_address'],
                            current_depth + 1,
                            root_workbook_path or workbook_path
                        )
                        node['children'].append(child_node)
                    except Exception as e:
                        # 添加錯誤節點
                        error_node = self._create_reference_error_node(
                            ref, current_depth + 1, root_workbook_path, str(e)
                        )
                        node['children'].append(error_node)
            
            # 移除已訪問標記（允許在不同分支中重複訪問）
            self.visited_cells.discard(cell_id)
            
            return node
            
        except Exception as e:
            # 移除已訪問標記
            self.visited_cells.discard(cell_id)
            return self._create_exception_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, str(e))
    
    def _contains_indirect(self, formula):
        """檢查公式是否包含 INDIRECT 函數"""
        return formula and 'INDIRECT' in formula.upper()
    
    def _resolve_indirect_formula(self, formula, workbook_path, sheet_name, cell_address):
        """
        解析 INDIRECT 公式為靜態引用
        
        Returns:
            dict: {
                'success': bool,
                'resolved_formula': str,
                'static_references': list,
                'calculation_details': dict
            }
        """
        try:
            print(f"[RESOLVE] 🔧 開始解析 INDIRECT 公式: {formula}")
            
            # 提取所有 INDIRECT 函數
            indirect_functions = self._extract_all_indirect_functions(formula)
            if not indirect_functions:
                print(f"[RESOLVE] ❌ 未找到 INDIRECT 函數")
                return {'success': False, 'error': 'No INDIRECT functions found'}
            
            print(f"[RESOLVE] 📋 找到 {len(indirect_functions)} 個 INDIRECT 函數")
            for i, func in enumerate(indirect_functions):
                print(f"[RESOLVE] INDIRECT {i+1}: {func['full_function']}")
                print(f"[RESOLVE] 內容: {func['content']}")
            
            resolved_formula = formula
            static_references = []
            calculation_details = []
            
            # 逐個解析 INDIRECT 函數
            for i, indirect_func in enumerate(indirect_functions):
                print(f"[RESOLVE] 🎯 處理第 {i+1} 個 INDIRECT 函數...")
                try:
                    # 計算 INDIRECT 內容
                    calc_result = self._calculate_indirect_with_excel(
                        indirect_func['content'], workbook_path, sheet_name, cell_address
                    )
                    
                    print(f"[RESOLVE] 📊 第 {i+1} 個 INDIRECT 計算結果: {calc_result.get('success', False)}")
                    
                    if calc_result and calc_result['success']:
                        static_ref = calc_result['static_reference']
                        print(f"[RESOLVE] ✅ 獲得靜態引用: {static_ref}")
                        
                        # *** 關鍵修正：檢查靜態引用格式 ***
                        # 如果靜態引用包含 !，說明這是一個完整的引用，需要解析
                        if '!' in static_ref:
                            print(f"[RESOLVE] 📋 靜態引用包含工作表引用，直接使用: {static_ref}")
                            final_static_ref = static_ref
                        else:
                            # 如果沒有 !，說明是同一工作表的儲存格，需要加上工作表名
                            final_static_ref = f"{sheet_name}!{static_ref}"
                            print(f"[RESOLVE] 📝 添加工作表名: {final_static_ref}")
                        
                        # 替換公式中的 INDIRECT 函數
                        old_formula = resolved_formula
                        resolved_formula = resolved_formula.replace(
                            indirect_func['full_function'], 
                            final_static_ref
                        )
                        print(f"[RESOLVE] 🔄 公式替換:")
                        print(f"[RESOLVE]   替換前: {old_formula}")
                        print(f"[RESOLVE]   替換後: {resolved_formula}")
                        
                        static_references.append(final_static_ref)
                        calculation_details.append({
                            'original_function': indirect_func['full_function'],
                            'content': indirect_func['content'],
                            'static_reference': final_static_ref,
                            'raw_excel_result': static_ref,
                            'calculation_details': calc_result.get('details', {})
                        })
                    else:
                        # 部分解析失敗，記錄但繼續
                        print(f"[RESOLVE] ❌ 第 {i+1} 個 INDIRECT 計算失敗: {calc_result.get('error', 'Unknown error')}")
                        calculation_details.append({
                            'original_function': indirect_func['full_function'],
                            'content': indirect_func['content'],
                            'error': calc_result.get('error', 'Unknown error')
                        })
                        
                except Exception as e:
                    print(f"[RESOLVE] ❌ 處理第 {i+1} 個 INDIRECT 函數異常: {e}")
                    calculation_details.append({
                        'original_function': indirect_func['full_function'],
                        'content': indirect_func['content'],
                        'error': str(e)
                    })
            
            success = len(static_references) > 0
            print(f"[RESOLVE] 📈 解析結果總結:")
            print(f"[RESOLVE]   成功: {success}")
            print(f"[RESOLVE]   最終公式: {resolved_formula}")
            print(f"[RESOLVE]   靜態引用: {static_references}")
            
            return {
                'success': success,
                'resolved_formula': resolved_formula,
                'static_references': static_references,
                'calculation_details': calculation_details,
                'original_formula': formula
            }
            
        except Exception as e:
            print(f"[RESOLVE] ❌❌❌ 解析 INDIRECT 公式發生異常 ❌❌❌")
            print(f"[RESOLVE] 異常: {e}")
            import traceback
            for line in traceback.format_exc().split('\n'):
                print(f"[RESOLVE]   {line}")
            return {
                'success': False,
                'error': str(e),
                'original_formula': formula
            }
    
    def _extract_all_indirect_functions(self, formula):
        """提取公式中所有的 INDIRECT 函數"""
        print(f"[EXTRACT] 🔍 開始提取 INDIRECT 函數: {formula}")
        indirect_functions = []
        
        # 查找所有 INDIRECT 位置
        search_start = 0
        while True:
            indirect_pos = formula.upper().find('INDIRECT(', search_start)
            if indirect_pos == -1:
                break
            
            print(f"[EXTRACT] 📍 找到 INDIRECT 位置: {indirect_pos}")
            
            # 提取完整的 INDIRECT 函數
            start_pos = indirect_pos + len('INDIRECT(')
            bracket_count = 1
            current_pos = start_pos
            
            print(f"[EXTRACT] 🔗 開始括號匹配，起始位置: {start_pos}")
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
                
                print(f"[EXTRACT] ✅ 成功提取 INDIRECT:")
                print(f"[EXTRACT]   完整函數: {full_function}")
                print(f"[EXTRACT]   內容: {content}")
                
                indirect_functions.append({
                    'full_function': full_function,
                    'content': content,
                    'start_pos': indirect_pos,
                    'end_pos': current_pos
                })
            else:
                print(f"[EXTRACT] ❌ 括號不匹配，跳過")
            
            search_start = current_pos
        
        print(f"[EXTRACT] 📋 總共提取到 {len(indirect_functions)} 個 INDIRECT 函數")
        return indirect_functions
    
    def _calculate_indirect_with_excel(self, indirect_content, workbook_path, sheet_name, cell_address):
        """使用 Excel 引擎計算 INDIRECT 內容"""
        wb = None
        
        try:
            print(f"[DEBUG] ===========================================")
            print(f"[DEBUG] 🚀 開始 Excel COM 計算 INDIRECT")
            print(f"[DEBUG] 📝 INDIRECT 內容: {indirect_content}")
            print(f"[DEBUG] 📂 目標文件: {workbook_path}")
            print(f"[DEBUG] 📊 目標工作表: {sheet_name}")
            print(f"[DEBUG] 📍 目標儲存格: {cell_address}")
            print(f"[DEBUG] ===========================================")
            
            # 驗證文件路徑
            if not os.path.exists(workbook_path):
                print(f"[DEBUG] ❌ 文件不存在: {workbook_path}")
                return {
                    'success': False,
                    'error': f'文件不存在: {workbook_path}',
                    'indirect_content': indirect_content
                }
            
            print(f"[DEBUG] ✅ 文件存在，準備開啟")
            
            # 確保 Excel 應用程序已啟動
            if not self.excel_app:
                print(f"[DEBUG] 📈 啟動 Excel 應用程序...")
                self.excel_app = win32com.client.Dispatch("Excel.Application")
                self.excel_app.Visible = False
                self.excel_app.DisplayAlerts = False
                self.excel_app.EnableEvents = False
                self.excel_app.ScreenUpdating = False
                print(f"[DEBUG] ✅ Excel 啟動完成")
            else:
                print(f"[DEBUG] ♻️ 復用現有 Excel 實例")
            
            # 以只讀模式打開文件
            print(f"[DEBUG] 📂 正在開啟工作簿...")
            print(f"[DEBUG]    文件路徑: {workbook_path}")
            try:
                wb = self.excel_app.Workbooks.Open(
                    workbook_path,
                    UpdateLinks=0,        # 不更新連結
                    ReadOnly=True,        # 只讀模式
                    IgnoreReadOnlyRecommended=True,
                    Notify=False
                )
                print(f"[DEBUG] ✅ 工作簿開啟成功")
                print(f"[DEBUG]    工作簿名稱: {wb.Name}")
                print(f"[DEBUG]    工作表數量: {wb.Worksheets.Count}")
            except Exception as e:
                print(f"[DEBUG] ❌ 開啟工作簿失敗: {e}")
                return {
                    'success': False,
                    'error': f'無法開啟工作簿: {e}',
                    'indirect_content': indirect_content
                }
            
            # 定位到原始儲存格位置（保持位置上下文）
            print(f"[DEBUG] 📍 定位到工作表和儲存格...")
            try:
                ws = wb.Worksheets(sheet_name)
                print(f"[DEBUG] ✅ 工作表定位成功: {ws.Name}")
                
                target_cell = ws.Range(cell_address)
                print(f"[DEBUG] ✅ 儲存格定位成功: {target_cell.Address}")
            except Exception as e:
                print(f"[DEBUG] ❌ 定位失敗: {e}")
                wb.Close(SaveChanges=False)
                return {
                    'success': False,
                    'error': f'無法定位工作表或儲存格: {e}',
                    'indirect_content': indirect_content
                }
            
            # 備份原始內容
            print(f"[DEBUG] 💾 備份原始儲存格內容...")
            original_value = None
            original_formula = None
            try:
                original_value = target_cell.Value
                original_formula = target_cell.Formula
                print(f"[DEBUG] ✅ 原始值: {original_value}")
                print(f"[DEBUG] ✅ 原始公式: {original_formula}")
            except Exception as e:
                print(f"[DEBUG] ⚠️ 備份原始內容失敗: {e}")
            
            # 設置測試公式並計算
            test_formula = f"={indirect_content}"
            print(f"[DEBUG] 🧮 設置測試公式: {test_formula}")
            print(f"[DEBUG] 📍 在位置 {sheet_name}!{cell_address} 計算")
            
            try:
                target_cell.Formula = test_formula
                print(f"[DEBUG] ✅ 公式設置成功")
            except Exception as e:
                print(f"[DEBUG] ❌ 設置公式失敗: {e}")
                wb.Close(SaveChanges=False)
                return {
                    'success': False,
                    'error': f'設置公式失敗: {e}',
                    'indirect_content': indirect_content
                }
            
            # 強制計算
            print(f"[DEBUG] ⚡ 開始強制計算...")
            try:
                target_cell.Calculate()
                print(f"[DEBUG] ✅ 儲存格計算完成")
                
                ws.Calculate()
                print(f"[DEBUG] ✅ 工作表計算完成")
                
                # 也可以嘗試整個工作簿計算
                wb.Application.Calculate()
                print(f"[DEBUG] ✅ 應用程序級別計算完成")
                
            except Exception as e:
                print(f"[DEBUG] ⚠️ 計算過程出現警告: {e}")
            
            # 獲取計算結果
            print(f"[DEBUG] 📊 獲取計算結果...")
            try:
                result = target_cell.Value
                print(f"[DEBUG] ✅ 計算結果: '{result}' (類型: {type(result)})")
                
                # 也嘗試獲取其他可能的結果格式
                try:
                    result_text = target_cell.Text
                    print(f"[DEBUG] 📝 結果文本格式: '{result_text}'")
                except:
                    pass
                    
                try:
                    result_formula = target_cell.Formula
                    print(f"[DEBUG] 📐 結果公式: '{result_formula}'")
                except:
                    pass
                    
            except Exception as e:
                print(f"[DEBUG] ❌ 獲取結果失敗: {e}")
                result = None
            
            # 檢查是否為錯誤
            if result is None:
                print(f"[DEBUG] ❌ 結果為 None - 計算可能失敗")
            elif self._is_excel_error(result):
                print(f"[DEBUG] ❌ 結果是 Excel 錯誤值: {result}")
                # 解碼錯誤值
                error_meanings = {
                    -2146826281: "#DIV/0!",
                    -2146826246: "#N/A",
                    -2146826259: "#NAME?",
                    -2146826288: "#NULL!",
                    -2146826252: "#NUM!",
                    -2146826265: "#REF!",
                    -2146826273: "#VALUE!"
                }
                if isinstance(result, int) and result in error_meanings:
                    print(f"[DEBUG] 錯誤解碼: {error_meanings[result]}")
            else:
                print(f"[DEBUG] ✅ 結果看起來正常")
            
            # 立即還原原始內容
            print(f"[DEBUG] 🔄 還原原始儲存格內容...")
            try:
                if original_formula:
                    target_cell.Formula = original_formula
                    print(f"[DEBUG] ✅ 還原原始公式: {original_formula}")
                elif original_value is not None:
                    target_cell.Value = original_value
                    print(f"[DEBUG] ✅ 還原原始值: {original_value}")
                else:
                    target_cell.Clear()
                    print(f"[DEBUG] ✅ 清除儲存格")
            except Exception as e:
                print(f"[DEBUG] ⚠️ 還原內容失敗: {e}")
                try:
                    target_cell.Clear()
                    print(f"[DEBUG] ✅ 強制清除儲存格")
                except:
                    pass
            
            # 關閉工作簿（不保存）
            print(f"[DEBUG] 📚 關閉工作簿...")
            try:
                wb.Close(SaveChanges=False)
                wb = None
                print(f"[DEBUG] ✅ 工作簿關閉成功")
            except Exception as e:
                print(f"[DEBUG] ⚠️ 關閉工作簿失敗: {e}")
            
            # 處理計算結果
            print(f"[DEBUG] 🔍 分析計算結果...")
            if result is None or self._is_excel_error(result):
                print(f"[DEBUG] ❌ 計算失敗，無法獲得有效結果")
                return {
                    'success': False,
                    'error': f'Excel 計算失敗. 結果: {result}',
                    'indirect_content': indirect_content
                }
            
            # 轉換結果為靜態引用
            static_reference = str(result).strip()
            print(f"[DEBUG] 🎯 轉換為靜態引用: '{static_reference}'")
            
            if not static_reference:
                print(f"[DEBUG] ❌ 靜態引用為空字串")
                return {
                    'success': False,
                    'error': '計算結果為空',
                    'indirect_content': indirect_content
                }
            
            print(f"[DEBUG] ===========================================")
            print(f"[DEBUG] ✅ Excel COM 計算成功完成！")
            print(f"[DEBUG] 📥 輸入: {indirect_content}")
            print(f"[DEBUG] 📤 輸出: {static_reference}")
            print(f"[DEBUG] ===========================================")
            
            return {
                'success': True,
                'static_reference': static_reference,
                'indirect_content': indirect_content,
                'details': {
                    'calculation_location': f"{sheet_name}!{cell_address}",
                    'test_formula': test_formula,
                    'raw_result': result,
                    'original_value': original_value,
                    'original_formula': original_formula
                }
            }
            
        except Exception as e:
            print(f"[DEBUG] ❌❌❌ Excel COM 計算發生嚴重異常 ❌❌❌")
            print(f"[DEBUG] 異常: {e}")
            import traceback
            print(f"[DEBUG] 詳細錯誤:")
            for line in traceback.format_exc().split('\n'):
                print(f"[DEBUG]   {line}")
            
            return {
                'success': False,
                'error': str(e),
                'indirect_content': indirect_content
            }
            
        finally:
            # 確保工作簿關閉
            try:
                if wb:
                    print(f"[DEBUG] 🔄 強制關閉工作簿...")
                    wb.Close(SaveChanges=False)
                    print(f"[DEBUG] ✅ 強制關閉成功")
            except Exception as e:
                print(f"[DEBUG] ⚠️ 強制關閉失敗: {e}")
    
    def _is_excel_error(self, result):
        """檢查是否為 Excel 錯誤值"""
        if isinstance(result, int) and result < 0:
            return True
        if isinstance(result, str) and result.startswith('#'):
            return True
        return False
    
    def _clean_formula(self, formula):
        """清理公式（保持原有邏輯）"""
        if not formula:
            return formula
        
        # 處理雙反斜線
        cleaned = formula.replace('\\\\', '\\')
        # 解碼 URL 編碼字符
        cleaned = unquote(cleaned)
        # 處理雙引號問題
        cleaned = re.sub(r"''([^']*?)''", r"'\1'", cleaned)
        
        return cleaned
    
    def _create_base_node(self, workbook_path, sheet_name, cell_address, current_depth, 
                         root_workbook_path, cell_info, original_formula, processed_formula,
                         indirect_resolved, indirect_details):
        """創建基本節點（包含 INDIRECT 信息）"""
        # 顯示地址邏輯（保持原有邏輯）
        filename = os.path.basename(workbook_path)
        dir_path = os.path.dirname(workbook_path)
        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
        
        if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
            # 外部引用
            filename_without_ext = filename.rsplit('.', 1)[0] if filename.endswith(('.xlsx', '.xls', '.xlsm')) else filename
            short_display_address = f"[{filename_without_ext}]{sheet_name}!{cell_address}"
            full_display_address = f"'{dir_path.replace(chr(92), '/')}/[{filename}]{sheet_name}'!{cell_address}"
            display_address = short_display_address
        else:
            # 本地引用
            short_display_address = f"{sheet_name}!{cell_address}"
            full_display_address = f"'{dir_path.replace(chr(92), '/')}/[{filename}]{sheet_name}'!{cell_address}"
            display_address = short_display_address
        
        # 構建節點
        node = {
            'address': display_address,
            'short_address': short_display_address,
            'full_address': full_display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'value': cell_info.get('display_value', 'N/A'),
            'calculated_value': cell_info.get('calculated_value', 'N/A'),
            'formula': processed_formula,
            'type': cell_info.get('cell_type', 'unknown'),
            'children': [],
            'depth': current_depth,
            'error': None,
            # *** 新增：INDIRECT 相關信息 ***
            'indirect_resolved': indirect_resolved,
            'original_formula': original_formula if indirect_resolved else None,
            'indirect_details': indirect_details
        }
        
        return node
    
    def _create_limit_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path):
        """創建深度限制節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, root_workbook_path)
        return {
            'address': display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'value': 'Max depth reached',
            'formula': None,
            'type': 'limit_reached',
            'children': [],
            'depth': current_depth,
            'error': 'Maximum recursion depth reached',
            'indirect_resolved': False
        }
    
    def _create_circular_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path):
        """創建循環引用節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, root_workbook_path)
        return {
            'address': display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'value': 'Circular reference',
            'formula': None,
            'type': 'circular_ref',
            'children': [],
            'depth': current_depth,
            'error': 'Circular reference detected',
            'indirect_resolved': False
        }
    
    def _create_error_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, error_msg):
        """創建錯誤節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, root_workbook_path)
        return {
            'address': display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'value': 'Error',
            'formula': None,
            'type': 'error',
            'children': [],
            'depth': current_depth,
            'error': error_msg,
            'indirect_resolved': False
        }
    
    def _create_exception_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, error_msg):
        """創建異常節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, root_workbook_path)
        return {
            'address': display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'value': 'Error',
            'formula': None,
            'type': 'error',
            'children': [],
            'depth': current_depth,
            'error': error_msg,
            'indirect_resolved': False
        }
    
    def _create_reference_error_node(self, ref, current_depth, root_workbook_path, error_msg):
        """創建引用錯誤節點"""
        display_address = self._get_display_address(ref['workbook_path'], ref['sheet_name'], ref['cell_address'], root_workbook_path)
        return {
            'address': display_address,
            'workbook_path': ref['workbook_path'],
            'sheet_name': ref['sheet_name'],
            'cell_address': ref['cell_address'],
            'value': 'Error',
            'formula': None,
            'type': 'error',
            'children': [],
            'depth': current_depth,
            'error': error_msg,
            'indirect_resolved': False
        }
    
    def _get_display_address(self, workbook_path, sheet_name, cell_address, root_workbook_path):
        """獲取顯示地址（統一邏輯）"""
        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
        if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
            filename = os.path.basename(workbook_path)
            filename_without_ext = filename.rsplit('.', 1)[0] if filename.endswith(('.xlsx', '.xls', '.xlsm')) else filename
            return f"[{filename_without_ext}]{sheet_name}!{cell_address}"
        else:
            return f"{sheet_name}!{cell_address}"
    
    # === 保持原有的所有其他方法不變 ===
    
    def parse_formula_references(self, formula, current_workbook_path, current_sheet_name):
        """
        Enhanced formula reference parser - 修正版（保持原有邏輯）
        """
        if not formula or not formula.startswith('='):
            return []

        references = []
        processed_spans = []
        
        # Normalize backslashes to handle cases with single or double backslashes
        normalized_formula = formula.replace('\\\\', '\\')
        
        def is_span_processed(start, end):
            for p_start, p_end in processed_spans:
                if start < p_end and end > p_start:
                    return True
            return False

        def add_processed_span(start, end):
            processed_spans.append((start, end))

        # 修正後的模式匹配
        patterns = [
            # 1. 外部引用 - 修正版，更精確的捕獲
            (
                'external',
                re.compile(
                    r"'?([^']*)\[([^\]]+\.(?:xlsx|xls|xlsm|xlsb))\]([^']*?)'?\s*!\s*(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)",
                    re.IGNORECASE
                )
            ),
            # 2. 本地引用（帶引號）
            (
                'local_quoted',
                re.compile(
                    r"'([^']+)'!(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)",
                    re.IGNORECASE
                )
            ),
            # 3. 本地引用（不帶引號）
            (
                'local_unquoted',
                re.compile(
                    r"([a-zA-Z0-9_\u4e00-\u9fa5][a-zA-Z0-9_\s\.\u4e00-\u9fa5]{0,30})!(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)",
                    re.IGNORECASE
                )
            ),
            # 4. 當前工作表範圍
            (
                'current_range',
                re.compile(
                    r"(?<![!'\[\]a-zA-Z0-9_\u4e00-\u9fa5])(?!\[)(\$?[A-Z]{1,3}\$?\d{1,7}:\s*\$?[A-Z]{1,3}\$?\d{1,7})(?![a-zA-Z0-9_\]])",
                    re.IGNORECASE
                )
            ),
            # 5. 當前工作表單個儲存格
            (
                'current_single',
                re.compile(
                    r"(?<![!'\[\]a-zA-Z0-9_\u4e00-\u9fa5])(?!\[)(\$?[A-Z]{1,3}\$?\d{1,7})(?![a-zA-Z0-9_:\]])",
                    re.IGNORECASE
                )
            )
        ]

        all_matches = []
        for p_type, pattern in patterns:
            for match in pattern.finditer(normalized_formula):
                all_matches.append({'type': p_type, 'match': match, 'span': match.span()})

        # 按優先級和位置排序
        type_priority = {'external': 0, 'local_quoted': 1, 'local_unquoted': 2, 'current_range': 3, 'current_single': 4}
        all_matches.sort(key=lambda x: (type_priority.get(x['type'], 99), x['span'][0], x['span'][1] - x['span'][0]))

        for item in all_matches:
            match = item['match']
            m_type = item['type']
            start, end = item['span']

            if is_span_processed(start, end):
                continue

            try:
                if m_type == 'external':
                    # 修正版：正確處理外部引用路徑
                    path_prefix, file_name, sheet_suffix, cell_ref = match.groups()
                    
                    # 組合完整檔案路徑
                    if path_prefix:
                        # 有路徑前綴，直接組合
                        full_file_path = os.path.join(path_prefix, file_name)
                    else:
                        # 沒有路徑前綴，檢查是否為當前檔案
                        current_file_name = os.path.basename(current_workbook_path)
                        if file_name.lower() == current_file_name.lower():
                            full_file_path = current_workbook_path
                        else:
                            # 外部檔案，使用當前目錄
                            current_dir = os.path.dirname(current_workbook_path)
                            full_file_path = os.path.join(current_dir, file_name)
                    
                    # 工作表名稱處理
                    sheet_name = sheet_suffix.strip("'") if sheet_suffix else "Sheet1"
                    
                    # Handle ranges vs single cells
                    if ':' in cell_ref:
                        range_refs = self._process_range_reference(
                            cell_ref, full_file_path, sheet_name, 'external'
                        )
                        references.extend(range_refs)
                    else:
                        references.append({
                            'workbook_path': full_file_path,
                            'sheet_name': sheet_name,
                            'cell_address': cell_ref.replace('$', ''),
                            'type': 'external'
                        })

                elif m_type in ('local_quoted', 'local_unquoted'):
                    sheet_name, cell_ref = match.groups()
                    sheet_name = sheet_name.strip("'")
                    
                    # Skip if it looks like a file name
                    if sheet_name.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
                        continue
                    
                    # Handle ranges vs single cells
                    if ':' in cell_ref:
                        range_refs = self._process_range_reference(
                            cell_ref, current_workbook_path, sheet_name, 'local'
                        )
                        references.extend(range_refs)
                    else:
                        references.append({
                            'workbook_path': current_workbook_path,
                            'sheet_name': sheet_name,
                            'cell_address': cell_ref.replace('$', ''),
                            'type': 'local'
                        })

                elif m_type in ('current_range', 'current_single'):
                    cell_ref = match.group(1)
                    
                    # Handle ranges vs single cells
                    if ':' in cell_ref:
                        range_refs = self._process_range_reference(
                            cell_ref, current_workbook_path, current_sheet_name, 'current'
                        )
                        references.extend(range_refs)
                    else:
                        references.append({
                            'workbook_path': current_workbook_path,
                            'sheet_name': current_sheet_name,
                            'cell_address': cell_ref.replace('$', ''),
                            'type': 'current'
                        })

                add_processed_span(start, end)
                
            except Exception as e:
                print(f"Warning: Could not process reference from match '{match.group(0)}': {e}")
                continue

        return references
    
    def _process_range_reference(self, range_ref, workbook_path, sheet_name, ref_type):
        """
        處理range引用，根據大小決定展開或摘要
        """
        try:
            # 計算range大小
            cell_count = self._calculate_range_size(range_ref)
            
            if cell_count <= self.range_expand_threshold:
                # 小範圍：展開為個別儲存格
                return self._expand_range_to_cells(range_ref, workbook_path, sheet_name, ref_type)
            else:
                # 大範圍：創建摘要節點
                return self._create_range_summary(range_ref, workbook_path, sheet_name, ref_type, cell_count)
                
        except Exception as e:
            print(f"Warning: Could not process range {range_ref}: {e}")
            # 發生錯誤時，創建單個摘要節點
            return [{
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': range_ref,
                'type': f'{ref_type}_range_error',
                'is_range_summary': True,
                'range_info': f'Error processing range: {e}'
            }]
    
    def _calculate_range_size(self, range_ref):
        """計算range包含的儲存格數量"""
        try:
            # 移除$符號並分割range
            clean_range = range_ref.replace('$', '').strip()
            start_cell, end_cell = clean_range.split(':')
            
            # 解析起始儲存格
            start_col, start_row = self._parse_cell_address(start_cell.strip())
            # 解析結束儲存格  
            end_col, end_row = self._parse_cell_address(end_cell.strip())
            
            # 計算行列數量
            row_count = abs(end_row - start_row) + 1
            col_count = abs(end_col - start_col) + 1
            
            return row_count * col_count
            
        except Exception as e:
            print(f"Warning: Could not calculate range size for {range_ref}: {e}")
            return 999  # 返回大數值，強制使用摘要模式
    
    def _parse_cell_address(self, cell_address):
        """解析儲存格地址為列號和行號"""
        import re
        match = re.match(r'([A-Z]+)(\d+)', cell_address.upper())
        if not match:
            raise ValueError(f"Invalid cell address: {cell_address}")
        
        col_letters = match.group(1)
        row_num = int(match.group(2))
        
        # 轉換列字母為數字 (A=1, B=2, ...)
        col_num = 0
        for char in col_letters:
            col_num = col_num * 26 + (ord(char) - ord('A') + 1)
        
        return col_num, row_num
    
    def _expand_range_to_cells(self, range_ref, workbook_path, sheet_name, ref_type):
        """將range展開為個別儲存格引用"""
        try:
            clean_range = range_ref.replace('$', '').strip()
            start_cell, end_cell = clean_range.split(':')
            
            start_col, start_row = self._parse_cell_address(start_cell.strip())
            end_col, end_row = self._parse_cell_address(end_cell.strip())
            
            # 確保起始位置小於結束位置
            min_col, max_col = min(start_col, end_col), max(start_col, end_col)
            min_row, max_row = min(start_row, end_row), max(start_row, end_row)
            
            references = []
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    # 轉換列號回字母
                    col_letters = self._col_num_to_letters(col)
                    cell_address = f"{col_letters}{row}"
                    
                    references.append({
                        'workbook_path': workbook_path,
                        'sheet_name': sheet_name,
                        'cell_address': cell_address,
                        'type': f'{ref_type}_from_range',
                        'original_range': range_ref
                    })
            
            return references
            
        except Exception as e:
            print(f"Warning: Could not expand range {range_ref}: {e}")
            return []
    
    def _col_num_to_letters(self, col_num):
        """將列號轉換為字母 (1=A, 2=B, ...)"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(ord('A') + (col_num % 26)) + result
            col_num //= 26
        return result
    
    def _create_range_summary(self, range_ref, workbook_path, sheet_name, ref_type, cell_count):
        """創建range摘要節點"""
        # 生成range的hash值用於顯示
        import hashlib
        range_hash = hashlib.md5(f"{workbook_path}|{sheet_name}|{range_ref}".encode()).hexdigest()[:8]
        
        # 計算維度信息
        try:
            clean_range = range_ref.replace('$', '').strip()
            start_cell, end_cell = clean_range.split(':')
            start_col, start_row = self._parse_cell_address(start_cell.strip())
            end_col, end_row = self._parse_cell_address(end_cell.strip())
            
            row_count = abs(end_row - start_row) + 1
            col_count = abs(end_col - start_col) + 1
            dimension_info = f"{row_count}行×{col_count}列"
        except:
            dimension_info = f"{cell_count}個儲存格"
        
        return [{
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': range_ref,
            'type': f'{ref_type}_range_summary',
            'is_range_summary': True,
            'range_info': f'Range摘要 (Hash: {range_hash}, {dimension_info}, 共{cell_count}個儲存格)'
        }]
    
    def get_explosion_summary(self, root_node):
        """
        獲取爆炸分析摘要 (增強版 - 包含 INDIRECT 統計)
        """
        def count_nodes(node):
            count = 1
            for child in node.get('children', []):
                count += count_nodes(child)
            return count
        
        def get_max_depth(node):
            if not node.get('children'):
                return node.get('depth', 0)
            children_depths = [get_max_depth(child) for child in node['children']]
            if not children_depths:
                return node.get('depth', 0)
            return max(children_depths)
        
        def count_by_type(node, type_counts=None):
            if type_counts is None:
                type_counts = {}
            
            node_type = node.get('type', 'unknown')
            type_counts[node_type] = type_counts.get(node_type, 0) + 1
            
            for child in node.get('children', []):
                count_by_type(child, type_counts)
            
            return type_counts
        
        def count_indirect_resolutions(node, indirect_stats=None):
            if indirect_stats is None:
                indirect_stats = {'resolved': 0, 'failed': 0, 'total': 0}
            
            if node.get('indirect_resolved'):
                indirect_stats['resolved'] += 1
                indirect_stats['total'] += 1
            elif node.get('indirect_details') and not node.get('indirect_resolved'):
                indirect_stats['failed'] += 1
                indirect_stats['total'] += 1
            
            for child in node.get('children', []):
                count_indirect_resolutions(child, indirect_stats)
            
            return indirect_stats
        
        # 基本統計
        basic_stats = {
            'total_nodes': count_nodes(root_node),
            'max_depth': get_max_depth(root_node),
            'type_distribution': count_by_type(root_node),
            'circular_references': len(self.circular_refs),
            'circular_ref_list': self.circular_refs
        }
        
        # INDIRECT 統計
        indirect_stats = count_indirect_resolutions(root_node)
        
        # 合併統計
        return {
            **basic_stats,
            'indirect_resolution': indirect_stats,
            'indirect_resolution_log': self.indirect_resolution_log
        }


def explode_cell_dependencies(workbook_path, sheet_name, cell_address, max_depth=10, range_expand_threshold=5, enable_indirect_resolution=True):
    """
    便捷函數：爆炸分析指定儲存格的依賴關係 (支援 INDIRECT)
    
    Args:
        workbook_path: Excel 檔案路徑
        sheet_name: 工作表名稱
        cell_address: 儲存格地址
        max_depth: 最大遞歸深度
        range_expand_threshold: 範圍展開閾值
        enable_indirect_resolution: 是否啟用 INDIRECT 解析
    
    Returns:
        tuple: (dependency_tree, summary)
    """
    exploder = DependencyExploder(
        max_depth=max_depth, 
        range_expand_threshold=range_expand_threshold,
        enable_indirect_resolution=enable_indirect_resolution
    )
    
    try:
        dependency_tree = exploder.explode_dependencies(workbook_path, sheet_name, cell_address)
        summary = exploder.get_explosion_summary(dependency_tree)
        
        return dependency_tree, summary
    finally:
        # 確保清理 Excel 進程
        exploder._cleanup_excel()


# 測試函數
if __name__ == "__main__":
    # 測試用例
    test_workbook = r"C:\Users\user\Desktop\pytest\test.xlsx"
    test_sheet = "Sheet1"
    test_cell = "A1"
    
    try:
        print("=== 增強版依賴分析測試 ===")
        print(f"測試文件: {test_workbook}")
        print(f"測試位置: {test_sheet}!{test_cell}")
        print()
        
        # 測試增強版（啟用 INDIRECT 解析）
        print("1. 測試增強版（啟用 INDIRECT 解析）:")
        tree_enhanced, summary_enhanced = explode_cell_dependencies(
            test_workbook, test_sheet, test_cell, enable_indirect_resolution=True
        )
        
        print("依賴樹（增強版）:")
        print(tree_enhanced)
        print("\n摘要（增強版）:")
        print(summary_enhanced)
        
        # 顯示 INDIRECT 解析日誌
        if summary_enhanced.get('indirect_resolution_log'):
            print("\nINDIRECT 解析日誌:")
            for log_entry in summary_enhanced['indirect_resolution_log']:
                print(f"  {log_entry}")
        
        print("\n" + "="*50)
        
        # 測試向後兼容版（禁用 INDIRECT 解析）
        print("2. 測試向後兼容版（禁用 INDIRECT 解析）:")
        tree_legacy, summary_legacy = explode_cell_dependencies(
            test_workbook, test_sheet, test_cell, enable_indirect_resolution=False
        )
        
        print("依賴樹（向後兼容）:")
        print(tree_legacy)
        print("\n摘要（向後兼容）:")
        print(summary_legacy)
        
    except Exception as e:
        print(f"測試失敗: {e}")
        import traceback
        traceback.print_exc()