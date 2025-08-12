# -*- coding: utf-8 -*-
"""
Enhanced Dependency Exploder with Ultra Safe COM Management - ULTRA SAFE VERSION
超安全版本 - 完全避免檔案鎖定問題 + INDEX 函數解析支援 (完整版本)
Last Updated: 2025-08-10 07:15:00 UTC
User: ckcm0210
"""

import re
import os
import win32com.client
import pythoncom
import time
import psutil
from urllib.parse import unquote
from utils.openpyxl_resolver import read_cell_with_resolved_references
from utils.range_processor import range_processor, process_formula_ranges
import datetime
import gc
import traceback
import hashlib
import uuid

class ProgressCallback:
    """進度回調接口 - 支持實時訊息和累積日誌"""
    def __init__(self, progress_var=None, popup_window=None, log_text_widget=None):
        self.progress_var = progress_var
        self.popup_window = popup_window
        self.log_text_widget = log_text_widget
        self.current_step = 0
        self.total_steps = 0
   
    def update_progress(self, message, step=None):
        """更新進度訊息 - 同時更新實時顯示和累積日誌"""
        if step is not None:
            self.current_step = step
            
        # 格式化訊息
        if self.total_steps > 0:
            progress_text = f"[{self.current_step}/{self.total_steps}] {message}"
        else:
            progress_text = message
            
        # 更新實時進度標籤
        if self.progress_var:
            self.progress_var.set(progress_text)
            
        # 累積到日誌區域
        if self.log_text_widget:
            try:
                import datetime
                timestamp = datetime.datetime.now().strftime("%H:%M:%S")
                log_entry = f"[{timestamp}] {progress_text}\n"
                
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
        print(f"[Explode Progress] {progress_text}")
        
    def set_total_steps(self, total):
        """設置總步驟數"""
        self.total_steps = total
        self.current_step = 0

class EnhancedDependencyExploder:
    """超安全版公式依賴鏈爆炸分析器 - 完全避免檔案鎖定 + INDEX支援"""
    
    def __init__(self, max_depth=10, range_expand_threshold=5, progress_callback=None):
        self.max_depth = max_depth
        self.range_expand_threshold = range_expand_threshold
        self.visited_cells = set()
        self.circular_refs = []
        self.progress_callback = progress_callback or ProgressCallback()
        self.processed_count = 0
        # 記錄創建的 Excel 實例和 PID
        self.our_excel_instances = {}
        self.excel_process_pids = set()  # 記錄我們創建的 Excel 程序 PID
        self.indirect_resolution_log = []
        self.index_resolution_log = []  # 新增：INDEX解析日誌
        
        # 初始化 COM
        try:
            pythoncom.CoInitialize()
        except:
            pass
    
    def __del__(self):
        """析構函數：超安全清理"""
        try:
            self._ultra_safe_cleanup()
        except:
            pass
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    # === 超安全清理相關方法（保持不變）===
    def _get_excel_processes_before(self):
        """獲取執行前的 Excel 程序列表"""
        excel_pids = set()
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'].lower() == 'excel.exe':
                    excel_pids.add(proc.info['pid'])
        except:
            pass
        return excel_pids
    
    def _get_new_excel_processes(self, before_pids):
        """獲取新創建的 Excel 程序"""
        new_pids = set()
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                if (proc.info['name'].lower() == 'excel.exe' and 
                    proc.info['pid'] not in before_pids):
                    new_pids.add(proc.info['pid'])
        except:
            pass
        return new_pids
    
    def _ultra_safe_cleanup(self):
        """超安全清理 - 確保檔案完全釋放"""
        import gc
        
        self.progress_callback.update_progress("[ULTRA-SAFE] 開始超安全清理...")
        
        # 第一階段：正常清理 COM 物件
        instance_keys = list(self.our_excel_instances.keys())
        for instance_key in instance_keys:
            try:
                self._cleanup_single_instance(instance_key)
            except Exception as e:
                self.progress_callback.update_progress(f"[ULTRA-SAFE] 清理實例失敗: {e}")
        
        # 清空記錄
        self.our_excel_instances.clear()
        
        # 第二階段：強制垃圾回收
        self.progress_callback.update_progress("[ULTRA-SAFE] 執行強制垃圾回收...")
        for i in range(5):
            gc.collect()
            time.sleep(0.2)
        
        # 第三階段：檢查並終止我們創建的 Excel 程序
        self.progress_callback.update_progress("[ULTRA-SAFE] 檢查殘留的 Excel 程序...")
        remaining_pids = self._check_and_terminate_our_excel_processes()
        
        if remaining_pids:
            self.progress_callback.update_progress(f"[ULTRA-SAFE] 已終止 {len(remaining_pids)} 個殘留的 Excel 程序")
        
        # 第四階段：等待檔案系統釋放
        self.progress_callback.update_progress("[ULTRA-SAFE] 等待檔案系統完全釋放...")
        time.sleep(1.0)  # 給檔案系統更多時間釋放鎖定
        
        # 重置內部狀態
        self.visited_cells.clear()
        self.circular_refs.clear()
        self.indirect_resolution_log.clear()
        self.index_resolution_log.clear()  # 新增：清理INDEX日誌
        self.processed_count = 0
        self.excel_process_pids.clear()
        
        self.progress_callback.update_progress("[ULTRA-SAFE] ✓ 超安全清理完成，檔案已完全釋放")
    
    def _cleanup_single_instance(self, instance_key):
        """清理單個 Excel 實例"""
        if instance_key not in self.our_excel_instances:
            return
            
        instance_info = self.our_excel_instances[instance_key]
        excel_app = instance_info.get('app')
        wb = instance_info.get('workbook')
        instance_id = instance_info.get('instance_id', 'Unknown')
        
        self.progress_callback.update_progress(f"[ULTRA-SAFE] 清理實例: {instance_id}")
        
        try:
            # 1. 關閉工作簿
            if wb:
                try:
                    wb.Saved = True  # 標記為已保存，避免保存提示
                    wb.Close(SaveChanges=False)
                    self.progress_callback.update_progress(f"[ULTRA-SAFE] 工作簿已關閉: {instance_id}")
                except Exception as wb_error:
                    self.progress_callback.update_progress(f"[ULTRA-SAFE] 關閉工作簿失敗: {wb_error}")
                
                # 強制釋放工作簿引用
                try:
                    del wb
                except:
                    pass
            
            # 2. 關閉 Excel 應用程式
            if excel_app:
                try:
                    excel_app.DisplayAlerts = False
                    excel_app.ScreenUpdating = False
                    excel_app.EnableEvents = False
                    
                    # 關閉所有工作簿
                    try:
                        while excel_app.Workbooks.Count > 0:
                            excel_app.Workbooks(1).Close(SaveChanges=False)
                    except:
                        pass
                    
                    # 退出 Excel
                    excel_app.Quit()
                    self.progress_callback.update_progress(f"[ULTRA-SAFE] Excel 應用程式已退出: {instance_id}")
                    
                except Exception as quit_error:
                    self.progress_callback.update_progress(f"[ULTRA-SAFE] 退出 Excel 失敗: {quit_error}")
                
                # 強制釋放 Excel 引用
                try:
                    del excel_app
                except:
                    pass
            
            # 3. 從記錄中移除
            try:
                del self.our_excel_instances[instance_key]
            except:
                pass
                
        except Exception as e:
            self.progress_callback.update_progress(f"[ULTRA-SAFE] 清理實例異常: {e}")
    
    def _check_and_terminate_our_excel_processes(self):
        """檢查並終止我們創建的 Excel 程序"""
        terminated_pids = []
        
        for pid in list(self.excel_process_pids):
            try:
                if psutil.pid_exists(pid):
                    proc = psutil.Process(pid)
                    if proc.name().lower() == 'excel.exe':
                        self.progress_callback.update_progress(f"[ULTRA-SAFE] 終止 Excel 程序 PID: {pid}")
                        proc.terminate()
                        
                        # 等待程序終止
                        try:
                            proc.wait(timeout=3)  # 等待最多3秒
                        except psutil.TimeoutExpired:
                            # 如果程序沒有正常終止，強制殺死
                            try:
                                proc.kill()
                                proc.wait(timeout=2)
                            except:
                                pass
                        
                        terminated_pids.append(pid)
                
                # 從記錄中移除
                self.excel_process_pids.discard(pid)
                
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                # 程序已經不存在或無法存取
                self.excel_process_pids.discard(pid)
            except Exception as e:
                self.progress_callback.update_progress(f"[ULTRA-SAFE] 終止程序 {pid} 失敗: {e}")
        
        return terminated_pids
    
    def _open_workbook_for_calculation(self, workbook_path):
        """開啟工作簿 - 超安全版本"""
        import uuid
        
        # 記錄執行前的 Excel 程序
        before_pids = self._get_excel_processes_before()
        
        # 創建唯一實例標識
        instance_key = f"{workbook_path}_{int(time.time())}_{uuid.uuid4().hex[:8]}"
        
        if instance_key in self.our_excel_instances:
            return self.our_excel_instances[instance_key]
        
        excel_app = None
        wb = None
        
        try:
            self.progress_callback.update_progress(f"[ULTRA-SAFE] 創建完全隔離的 Excel 實例: {os.path.basename(workbook_path)}")
            
            # 使用 DispatchEx 創建完全獨立的實例
            excel_app = win32com.client.DispatchEx("Excel.Application")
            
            # 記錄新創建的 Excel 程序
            time.sleep(0.5)  # 給程序啟動一點時間
            new_pids = self._get_new_excel_processes(before_pids)
            self.excel_process_pids.update(new_pids)
            if new_pids:
                self.progress_callback.update_progress(f"[ULTRA-SAFE] 記錄新 Excel 程序 PID: {new_pids}")
            
            # 設置為完全隔離模式
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            excel_app.EnableEvents = False
            excel_app.ScreenUpdating = False
            excel_app.Interactive = False
            excel_app.AskToUpdateLinks = False
            
            # 額外的隔離設置
            try:
                excel_app.ShowWindowsInTaskbar = False
                excel_app.WindowState = -4140  # xlMinimized
                excel_app.UserControl = False  # 重要：設置為非用戶控制
            except:
                pass
            
            self.progress_callback.update_progress(f"[ULTRA-SAFE] 隔離設置完成，開啟工作簿...")
            
            # 以唯讀模式開啟，使用更嚴格的參數
            wb = excel_app.Workbooks.Open(
                workbook_path,
                UpdateLinks=0,          # 不更新連結
                ReadOnly=True,          # 唯讀模式
                Format=5,               # 自動格式
                Password="",            # 無密碼
                WriteResPassword="",    # 無寫入密碼
                IgnoreReadOnlyRecommended=True,  # 忽略唯讀建議
                Origin=1,               # 原始格式
                Delimiter="",           # 無分隔符
                Editable=False,         # 不可編輯
                Notify=False,           # 不通知
                Converter=0,            # 無轉換器
                AddToMru=False,         # 不加入最近使用清單
                Local=False,            # 非本地化
                CorruptLoad=0           # 不載入損壞檔案
            )
            
            # 設置計算模式為手動
            try:
                excel_app.Calculation = -4135  # xlCalculationManual
                wb.Application.Calculation = -4135
            except:
                pass
            
            # 禁用所有可能的互動
            try:
                excel_app.CutCopyMode = False
                excel_app.StatusBar = False
            except:
                pass
            
            # 記錄實例信息
            instance_info = {
                'app': excel_app,
                'workbook': wb,
                'instance_id': instance_key,
                'workbook_name': wb.Name,
                'created_time': time.time(),
                'file_path': workbook_path,
                'process_pids': new_pids.copy()
            }
            
            self.our_excel_instances[instance_key] = instance_info
            
            self.progress_callback.update_progress(f"[ULTRA-SAFE] 隔離實例創建成功: {wb.Name}")
            return instance_info
            
        except Exception as e:
            self.progress_callback.update_progress(f"[ULTRA-SAFE] 創建隔離實例失敗: {e}")
            
            # 清理失敗的實例
            try:
                if wb:
                    wb.Close(SaveChanges=False)
                    del wb
            except:
                pass
            
            try:
                if excel_app:
                    excel_app.Quit()
                    del excel_app
            except:
                pass
            
            # 清理可能創建的程序
            new_pids = self._get_new_excel_processes(before_pids)
            for pid in new_pids:
                try:
                    proc = psutil.Process(pid)
                    proc.terminate()
                except:
                    pass
            
            raise e
    
    def _close_specific_instance(self, instance_key):
        """關閉特定實例 - 超安全版本"""
        if instance_key not in self.our_excel_instances:
            return
        
        self.progress_callback.update_progress(f"[ULTRA-SAFE] 開始關閉實例: {instance_key}")
        
        try:
            self._cleanup_single_instance(instance_key)
            
            # 額外等待時間確保完全釋放
            time.sleep(0.5)
            
            self.progress_callback.update_progress(f"[ULTRA-SAFE] 實例已安全關閉: {instance_key}")
            
        except Exception as e:
            self.progress_callback.update_progress(f"[ULTRA-SAFE] 關閉實例失敗: {e}")
    
    def _calculate_indirect_safely(self, indirect_content, workbook_path, sheet_name, cell_address):
        """安全計算 INDIRECT - 使用完全隔離的實例"""
        temp_instance = None
        temp_cell = None
        original_value = None
        original_formula = None
        
        try:
            self.progress_callback.update_progress(f"[ULTRA-SAFE-CALC] 開始安全計算: {indirect_content}")
            
            # 驗證文件路徑
            if not os.path.exists(workbook_path):
                return {
                    'success': False,
                    'error': f'文件不存在: {workbook_path}',
                    'indirect_content': indirect_content
                }
            
            # 為計算創建臨時的完全隔離實例
            temp_instance = self._open_workbook_for_calculation(workbook_path)
            
            wb = temp_instance['workbook']
            excel_app = temp_instance['app']
            instance_key = temp_instance['instance_id']
            
            # 定位目標儲存格
            try:
                ws = wb.Worksheets(sheet_name)
                temp_cell = ws.Range(cell_address)
            except Exception as e:
                return {
                    'success': False,
                    'error': f'無法定位工作表或儲存格: {e}',
                    'indirect_content': indirect_content
                }
            
            # 備份原始內容
            try:
                original_value = temp_cell.Value
                original_formula = temp_cell.Formula
            except:
                pass
            
            # 執行計算
            test_formula = f"={indirect_content}"
            calculation_result = None
            old_calculation = None
            
            try:
                # 暫時啟用自動計算
                try:
                    old_calculation = excel_app.Calculation
                    excel_app.Calculation = -4105  # xlCalculationAutomatic
                except:
                    pass
                
                # 設置公式並計算
                temp_cell.Formula = test_formula
                temp_cell.Calculate()
                
                # 獲取結果
                calculation_result = temp_cell.Value
                self.progress_callback.update_progress(f"[ULTRA-SAFE-CALC] 計算完成，結果: '{calculation_result}'")
                
            except Exception as calc_error:
                self.progress_callback.update_progress(f"[ULTRA-SAFE-CALC] 計算失敗: {calc_error}")
                return {
                    'success': False,
                    'error': f'計算失敗: {calc_error}',
                    'indirect_content': indirect_content
                }
            
            finally:
                # 還原計算模式
                if old_calculation is not None:
                    try:
                        excel_app.Calculation = old_calculation
                    except:
                        pass
                
                # 還原儲存格內容
                try:
                    if original_formula:
                        temp_cell.Formula = original_formula
                    elif original_value is not None:
                        temp_cell.Value = original_value
                    else:
                        temp_cell.Clear()
                except:
                    pass
                
                # 立即關閉臨時實例
                try:
                    self._close_specific_instance(instance_key)
                    self.progress_callback.update_progress(f"[ULTRA-SAFE-CALC] 臨時實例已安全關閉")
                except:
                    pass
            
            # 處理計算結果
            if calculation_result is None or self._is_excel_error(calculation_result):
                return {
                    'success': False,
                    'error': f'Excel計算返回錯誤: {calculation_result}',
                    'indirect_content': indirect_content
                }
            
            static_reference = str(calculation_result).strip()
            if not static_reference:
                return {
                    'success': False,
                    'error': '計算結果為空字串',
                    'indirect_content': indirect_content
                }
            
            return {
                'success': True,
                'static_reference': static_reference,
                'indirect_content': indirect_content
            }
            
        except Exception as e:
            self.progress_callback.update_progress(f"[ULTRA-SAFE-CALC] 計算異常: {e}")
            
            # 異常情況下的清理
            if temp_instance:
                try:
                    self._close_specific_instance(temp_instance['instance_id'])
                except:
                    pass
            
            return {
                'success': False,
                'error': str(e),
                'indirect_content': indirect_content
            }

    # === INDEX 解析方法（來自debug版本）===
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
                temp_references = self._parse_formula_references_accurate(f"={array_param}", workbook_path, sheet_name)
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
                        row_calc = self._calculate_indirect_safely(row_param, workbook_path, sheet_name, cell_address)
                        col_calc = self._calculate_indirect_safely(col_param, workbook_path, sheet_name, cell_address)
                        
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

    # === 完整的 explode_dependencies 方法 ===
    def explode_dependencies(self, workbook_path, sheet_name, cell_address, current_depth=0, root_workbook_path=None):
        """遞歸展開公式依賴鏈 - 超安全版本 + INDEX支援 (完整版本)"""
        # 更新進度
        if current_depth == 0:
            self.progress_callback.update_progress("正在初始化依賴關係分析...")
            self.processed_count = 0
        
        self.processed_count += 1
        
        # 創建唯一標識符
        cell_id = f"{workbook_path}|{sheet_name}|{cell_address}"
        
        # 顯示當前處理的儲存格
        filename = os.path.basename(workbook_path)
        current_ref = f"{filename}!{sheet_name}!{cell_address}"
        self.progress_callback.update_progress(
            f"正在分析 {current_ref} (深度: {current_depth}/{self.max_depth}, 已處理: {self.processed_count})"
        )
        
        # 檢查遞歸深度限制
        if current_depth >= self.max_depth:
            self.progress_callback.update_progress(f"警告：達到最大遞歸深度限制 ({self.max_depth})")
            return self._create_limit_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
        
        # 檢查循環引用
        if cell_id in self.visited_cells:
            self.circular_refs.append(cell_id)
            self.progress_callback.update_progress(f"警告：檢測到循環引用 {current_ref}")
            return self._create_circular_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
        
        # 標記為已訪問
        self.visited_cells.add(cell_id)
        
        try:
            # 檢查是否為範圍引用，如果是則跳過openpyxl讀取
            if ':' in cell_address:
                self.progress_callback.update_progress(f"檢測到範圍引用，跳過讀取: {current_ref}")
                return self._create_error_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, "範圍引用不支持直接讀取")
            
            # 讀取儲存格內容
            self.progress_callback.update_progress(f"正在讀取儲存格內容: {current_ref}")
            cell_info = read_cell_with_resolved_references(workbook_path, sheet_name, cell_address)
            
            if 'error' in cell_info:
                self.progress_callback.update_progress(f"錯誤：無法讀取 {current_ref} - {cell_info['error']}")
                return self._create_error_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, cell_info['error'])
            
            # 處理公式清理和動態函數解析
            original_formula = cell_info.get('formula')
            fixed_formula = None
            resolved_formula = None
            indirect_info = None
            index_info = None
            
            if original_formula:
                self.progress_callback.update_progress(f"正在處理公式: {current_ref}")
                fixed_formula = self._clean_formula(original_formula)
                resolved_formula = fixed_formula  # 默認等於fixed_formula
                
                # INDIRECT 處理
                if 'INDIRECT' in fixed_formula.upper():
                    self.progress_callback.update_progress(f"正在解析INDIRECT函數: {current_ref}")
                    try:
                        resolved_result = self._resolve_indirect_with_excel(
                            fixed_formula, workbook_path, sheet_name, cell_address
                        )
                        if resolved_result and resolved_result['success']:
                            resolved_formula = resolved_result['resolved_formula']
                            indirect_info = {
                                'has_indirect': True,
                                'success': True,
                                'resolved_formula': resolved_formula,
                                'details': resolved_result,
                                'internal_references': resolved_result.get('internal_references', [])
                            }
                            self.progress_callback.update_progress(f"INDIRECT解析完成，resolved: {resolved_formula}")
                            
                            # 記錄解析日誌
                            self.indirect_resolution_log.append({
                                'cell': f"{sheet_name}!{cell_address}",
                                'original': original_formula,
                                'resolved': resolved_formula,
                                'details': resolved_result
                            })
                        else:
                            indirect_info = {
                                'has_indirect': True,
                                'success': False,
                                'error': resolved_result.get('error', 'Unknown error'),
                                'internal_references': []
                            }
                            self.progress_callback.update_progress(f"INDIRECT解析失敗: {indirect_info['error']}")
                    except Exception as e:
                        indirect_info = {
                            'has_indirect': True,
                            'success': False,
                            'error': str(e),
                            'internal_references': []
                        }
                        self.progress_callback.update_progress(f"INDIRECT解析異常: {str(e)}")
                
                # INDEX 處理
                if 'INDEX(' in fixed_formula.upper():
                    self.progress_callback.update_progress(f"正在解析INDEX函數: {current_ref}")
                    try:
                        index_result = self._resolve_index_with_excel_corrected_simple(
                            resolved_formula, workbook_path, sheet_name, cell_address
                        )
                        if index_result and index_result['success']:
                            resolved_formula = index_result['resolved_formula']
                            index_info = {
                                'has_index': True,
                                'success': True,
                                'resolved_formula': resolved_formula,
                                'details': index_result,
                                'internal_references': index_result.get('internal_references', [])
                            }
                            self.progress_callback.update_progress(f"INDEX解析完成，resolved: {resolved_formula}")
                            
                            # 記錄INDEX解析日誌
                            self.index_resolution_log.append({
                                'cell': f"{sheet_name}!{cell_address}",
                                'original': original_formula,
                                'resolved': resolved_formula,
                                'details': index_result
                            })
                        else:
                            index_info = {
                                'has_index': True,
                                'success': False,
                                'error': index_result.get('error', 'Unknown error'),
                                'internal_references': []
                            }
                            self.progress_callback.update_progress(f"INDEX解析失敗: {index_info['error']}")
                    except Exception as e:
                        index_info = {
                            'has_index': True,
                            'success': False,
                            'error': str(e),
                            'internal_references': []
                        }
                        self.progress_callback.update_progress(f"INDEX解析異常: {str(e)}")

            # 創建節點
            node = self._create_node_with_dynamic_functions(
                workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, 
                cell_info, fixed_formula, resolved_formula, indirect_info, index_info
            )
            
            # 如果是公式，解析依賴關係
            if cell_info.get('cell_type') == 'formula' and cell_info.get('formula'):
                self.progress_callback.update_progress(f"正在解析公式依賴關係: {current_ref}")
                
                # 處理 INDIRECT 內部的引用
                if indirect_info and indirect_info.get('internal_references'):
                    self.progress_callback.update_progress(f"找到 {len(indirect_info['internal_references'])} 個 INDIRECT 內部引用，正在分析...")
                    
                    for i, internal_ref in enumerate(indirect_info['internal_references'], 1):
                        try:
                            ref_display = f"{os.path.basename(internal_ref['workbook_path'])}!{internal_ref['sheet_name']}!{internal_ref['cell_address']}"
                            self.progress_callback.update_progress(f"正在處理 INDIRECT 內部引用 {i}/{len(indirect_info['internal_references'])}: {ref_display}")
                            
                            child_node = self.explode_dependencies(
                                internal_ref['workbook_path'],
                                internal_ref['sheet_name'],
                                internal_ref['cell_address'],
                                current_depth + 1,
                                root_workbook_path or workbook_path
                            )
                            child_node['from_indirect_internal'] = True
                            node['children'].append(child_node)
                        except Exception as e:
                            self.progress_callback.update_progress(f"錯誤：處理 INDIRECT 內部引用失敗 {ref_display} - {str(e)}")
                
                # 處理 INDEX 內部的引用
                if index_info and index_info.get('internal_references'):
                    self.progress_callback.update_progress(f"找到 {len(index_info['internal_references'])} 個 INDEX 內部引用，正在分析...")
                    
                    for i, internal_ref in enumerate(index_info['internal_references'], 1):
                        try:
                            ref_display = f"{os.path.basename(internal_ref['workbook_path'])}!{internal_ref['sheet_name']}!{internal_ref['cell_address']}"
                            self.progress_callback.update_progress(f"正在處理 INDEX 內部引用 {i}/{len(index_info['internal_references'])}: {ref_display}")
                            
                            child_node = self.explode_dependencies(
                                internal_ref['workbook_path'],
                                internal_ref['sheet_name'],
                                internal_ref['cell_address'],
                                current_depth + 1,
                                root_workbook_path or workbook_path
                            )
                            child_node['from_index_internal'] = True
                            node['children'].append(child_node)
                        except Exception as e:
                            self.progress_callback.update_progress(f"錯誤：處理 INDEX 內部引用失敗 {ref_display} - {str(e)}")
                
                # 處理範圍地址
                formula_for_ranges = resolved_formula if resolved_formula else cell_info['formula']
                ranges = process_formula_ranges(formula_for_ranges, workbook_path, sheet_name)
                if ranges:
                    self.progress_callback.update_progress(f"找到 {len(ranges)} 個範圍，正在處理...")
                    
                    for i, range_info in enumerate(ranges, 1):
                        try:
                            range_display = f"{os.path.basename(range_info['workbook_path'])}!{range_info['sheet_name']}!{range_info['address']}"
                            self.progress_callback.update_progress(f"正在處理範圍 {i}/{len(ranges)}: {range_display}")
                            
                            range_node = self._create_range_node(range_info, current_depth + 1, root_workbook_path)
                            node['children'].append(range_node)
                            
                        except Exception as e:
                            self.progress_callback.update_progress(f"錯誤：處理範圍失敗 {range_display} - {str(e)}")
                            error_node = self._create_error_node(
                                range_info['workbook_path'], range_info['sheet_name'], range_info['address'], 
                                current_depth + 1, root_workbook_path, str(e)
                            )
                            node['children'].append(error_node)
                
                # 處理單個儲存格引用
                formula_to_parse = resolved_formula if resolved_formula else cell_info['formula']
                references = self._parse_formula_references_accurate(formula_to_parse, workbook_path, sheet_name)
                
                if references:
                    self.progress_callback.update_progress(f"找到 {len(references)} 個儲存格引用，正在遞歸分析...")
                    
                    for i, ref in enumerate(references, 1):
                        try:
                            # 檢查是否為範圍引用，如果是則跳過直接讀取
                            if ':' in ref['cell_address']:
                                self.progress_callback.update_progress(f"跳過範圍引用: {ref['cell_address']}")
                                continue
                                
                            ref_display = f"{os.path.basename(ref['workbook_path'])}!{ref['sheet_name']}!{ref['cell_address']}"
                            self.progress_callback.update_progress(f"正在處理引用 {i}/{len(references)}: {ref_display}")
                            
                            child_node = self.explode_dependencies(
                                ref['workbook_path'],
                                ref['sheet_name'],
                                ref['cell_address'],
                                current_depth + 1,
                                root_workbook_path or workbook_path
                            )
                            if resolved_formula != fixed_formula:
                                if indirect_info and indirect_info.get('success'):
                                    child_node['from_indirect_resolved'] = True
                                if index_info and index_info.get('success'):
                                    child_node['from_index_resolved'] = True
                            node['children'].append(child_node)
                        except Exception as e:
                            self.progress_callback.update_progress(f"錯誤：處理引用失敗 {ref_display} - {str(e)}")
                            error_node = self._create_error_node(
                                ref['workbook_path'], ref['sheet_name'], ref['cell_address'], 
                                current_depth + 1, root_workbook_path, str(e)
                            )
                            node['children'].append(error_node)
            
            # 移除已訪問標記
            self.visited_cells.discard(cell_id)
            
            # 在根節點完成時超安全清理
            if current_depth == 0:
                total_nodes = self._count_nodes(node)
                max_depth = self._get_max_depth(node)
                indirect_count = len([log for log in self.indirect_resolution_log if log.get('resolved')])
                index_count = len([log for log in self.index_resolution_log if log.get('resolved')])
                self.progress_callback.update_progress(
                    f"分析完成！共處理 {self.processed_count} 次，生成 {total_nodes} 個節點，最大深度: {max_depth}，成功解析 {indirect_count} 個 INDIRECT，{index_count} 個 INDEX"
                )
                
                # 超安全清理（只清理我們的實例）
                self.progress_callback.update_progress("[ULTRA-SAFE] 正在超安全釋放資源...")
                try:
                    self._ultra_safe_cleanup()
                    self.progress_callback.update_progress("[ULTRA-SAFE] ✓ 資源超安全釋放完成，您的Excel檔案完全不受影響")
                except Exception as cleanup_error:
                    self.progress_callback.update_progress(f"[ULTRA-SAFE] 超安全清理過程出錯: {cleanup_error}")
            
            return node
            
        except Exception as e:
            # 異常時也要超安全清理
            if current_depth == 0:
                try:
                    self._ultra_safe_cleanup()
                except:
                    pass
            self.visited_cells.discard(cell_id)
            self.progress_callback.update_progress(f"錯誤：處理 {current_ref} 時發生異常 - {str(e)}")
            return self._create_error_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, str(e))

    def _create_node_with_dynamic_functions(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, cell_info, fixed_formula, resolved_formula=None, indirect_info=None, index_info=None):
        """創建支持動態函數的節點"""
        filename = os.path.basename(workbook_path)
        dir_path = os.path.dirname(workbook_path)
        
        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
        
        if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
            short_display_address = f"[{filename}]{sheet_name}!{cell_address}"
            full_display_address = f"'{dir_path}\\[{filename}]{sheet_name}'!{cell_address}"
            display_address = short_display_address
        else:
            short_display_address = f"[{filename}]{sheet_name}!{cell_address}"
            full_display_address = f"'{dir_path}\\[{filename}]{sheet_name}'!{cell_address}"
            display_address = short_display_address

        node = {
            'address': display_address,
            'short_address': short_display_address,
            'full_address': full_display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'value': cell_info.get('display_value', 'N/A'),
            'calculated_value': cell_info.get('calculated_value', 'N/A'),
            'formula': fixed_formula,
            'type': cell_info.get('cell_type', 'unknown'),
            'children': [],
            'depth': current_depth,
            'error': None
        }
        
        # 處理動態函數信息
        has_dynamic_resolution = False
        
        # INDIRECT 信息
        if indirect_info and indirect_info.get('has_indirect'):
            node['has_indirect'] = True
            if indirect_info.get('success'):
                has_dynamic_resolution = True
                node['indirect_details'] = indirect_info.get('details')
                node['internal_references_count'] = len(indirect_info.get('internal_references', []))
            else:
                node['indirect_error'] = indirect_info.get('error')
        else:
            node['has_indirect'] = False
        
        # INDEX 信息
        if index_info and index_info.get('has_index'):
            node['has_index'] = True
            if index_info.get('success'):
                has_dynamic_resolution = True
                node['index_details'] = index_info.get('details')
                node['index_internal_references_count'] = len(index_info.get('internal_references', []))
            else:
                node['index_error'] = index_info.get('error')
        else:
            node['has_index'] = False
        
        # 設置resolved_formula
        if has_dynamic_resolution and resolved_formula and resolved_formula != fixed_formula:
            node['resolved_formula'] = resolved_formula
            
        return node

    # === 其他輔助方法（保持原有實現）===
    def force_cleanup(self):
        """公開的超安全清理方法"""
        self.progress_callback.update_progress("[USER] 用戶觸發超安全清理...")
        self._ultra_safe_cleanup()
        self.progress_callback.update_progress("[USER] 超安全清理完成，檔案已完全釋放")

    def _parse_formula_references_accurate(self, formula, current_workbook_path, current_sheet_name):
        """最準確的公式引用解析器"""
        if not formula or not formula.startswith('='):
            return []

        references = []
        processed_spans = []
        
        # 標準化反斜線
        normalized_formula = formula.replace('\\\\', '\\')
        
        def is_span_processed(start, end):
            for p_start, p_end in processed_spans:
                if start < p_end and end > p_start:
                    return True
            return False

        def add_processed_span(start, end):
            processed_spans.append((start, end))

        # 精確的模式匹配
        patterns = [
            # 外部引用
            (
                'external',
                re.compile(
                    r"'?([^']*)\[([^\]]+\.(?:xlsx|xls|xlsm|xlsb))\]([^']*?)'?\s*!\s*(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)",
                    re.IGNORECASE
                )
            ),
            # 本地引用（帶引號）
            (
                'local_quoted',
                re.compile(
                    r"'([^']+)'!(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)",
                    re.IGNORECASE
                )
            ),
            # 本地引用（不帶引號）
            (
                'local_unquoted',
                re.compile(
                    r"([a-zA-Z0-9_\u4e00-\u9fa5][a-zA-Z0-9_\s\.\u4e00-\u9fa5]{0,30})!(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)",
                    re.IGNORECASE
                )
            ),
            # 當前工作表範圍
            (
                'current_range',
                re.compile(
                    r"(?<![!'\[\]a-zA-Z0-9_\u4e00-\u9fa5])(?!\[)(\$?[A-Z]{1,3}\$?\d{1,7}:\s*\$?[A-Z]{1,3}\$?\d{1,7})(?![a-zA-Z0-9_\]])",
                    re.IGNORECASE
                )
            ),
            # 當前工作表單個儲存格
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
                    path_prefix, file_name, sheet_suffix, cell_ref = match.groups()
                    
                    # 組合完整檔案路徑
                    if path_prefix:
                        full_file_path = os.path.join(path_prefix, file_name)
                    else:
                        current_file_name = os.path.basename(current_workbook_path)
                        if file_name.lower() == current_file_name.lower():
                            full_file_path = current_workbook_path
                        else:
                            current_dir = os.path.dirname(current_workbook_path)
                            full_file_path = os.path.join(current_dir, file_name)
                    
                    sheet_name = sheet_suffix.strip("'") if sheet_suffix else "Sheet1"
                    
                    # 處理範圍 vs 單個儲存格
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
                    
                    # 跳過看起來像檔案名的工作表
                    if sheet_name.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
                        continue
                    
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
                self.progress_callback.update_progress(f"Warning: Could not process reference from match '{match.group(0)}': {e}")
                continue

        return references
    
    def _process_range_reference(self, range_ref, workbook_path, sheet_name, ref_type):
        """處理範圍引用，根據大小決定展開或摘要"""
        try:
            cell_count = self._calculate_range_size(range_ref)
            
            if cell_count <= self.range_expand_threshold:
                return self._expand_range_to_cells(range_ref, workbook_path, sheet_name, ref_type)
            else:
                return self._create_range_summary(range_ref, workbook_path, sheet_name, ref_type, cell_count)
                
        except Exception as e:
            self.progress_callback.update_progress(f"Warning: Could not process range {range_ref}: {e}")
            return [{
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': range_ref,
                'type': f'{ref_type}_range_error',
                'is_range_summary': True,
                'range_info': f'Error processing range: {e}'
            }]
    
    def _calculate_range_size(self, range_ref):
        """計算範圍包含的儲存格數量"""
        try:
            clean_range = range_ref.replace('$', '').strip()
            start_cell, end_cell = clean_range.split(':')
            
            start_col, start_row = self._parse_cell_address(start_cell.strip())
            end_col, end_row = self._parse_cell_address(end_cell.strip())
            
            row_count = abs(end_row - start_row) + 1
            col_count = abs(end_col - start_col) + 1
            
            return row_count * col_count
            
        except Exception as e:
            self.progress_callback.update_progress(f"Warning: Could not calculate range size for {range_ref}: {e}")
            return 999
    
    def _parse_cell_address(self, cell_address):
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
    
    def _expand_range_to_cells(self, range_ref, workbook_path, sheet_name, ref_type):
        """將範圍展開為個別儲存格引用"""
        try:
            clean_range = range_ref.replace('$', '').strip()
            start_cell, end_cell = clean_range.split(':')
            
            start_col, start_row = self._parse_cell_address(start_cell.strip())
            end_col, end_row = self._parse_cell_address(end_cell.strip())
            
            min_col, max_col = min(start_col, end_col), max(start_col, end_col)
            min_row, max_row = min(start_row, end_row), max(start_row, end_row)
            
            references = []
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
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
            self.progress_callback.update_progress(f"Warning: Could not expand range {range_ref}: {e}")
            return []
    
    def _col_num_to_letters(self, col_num):
        """將列號轉換為字母"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(ord('A') + (col_num % 26)) + result
            col_num //= 26
        return result
    
    def _create_range_summary(self, range_ref, workbook_path, sheet_name, ref_type, cell_count):
        """創建範圍摘要節點"""
        import hashlib
        range_hash = hashlib.md5(f"{workbook_path}|{sheet_name}|{range_ref}".encode()).hexdigest()[:8]
        
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
                temp_references = self._parse_formula_references_accurate(
                    f"={indirect_func['content']}", workbook_path, sheet_name
                )
                internal_references.extend(temp_references)
            
            # 逐個解析 INDIRECT 函數
            for i, indirect_func in enumerate(indirect_functions):
                self.progress_callback.update_progress(f"[INDIRECT] 處理第 {i+1} 個: {indirect_func['content']}")
                
                calc_result = self._calculate_indirect_safely(
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

    def _is_excel_error(self, result):
        """檢查是否為 Excel 錯誤值"""
        if isinstance(result, int) and result < 0:
            return True
        if isinstance(result, str) and result.startswith('#'):
            return True
        return False
    
    def _clean_formula(self, formula):
        """清理公式中的路徑問題"""
        if not formula:
            return formula
            
        fixed_formula = formula.replace('\\\\', '\\')
        fixed_formula = unquote(fixed_formula)
        fixed_formula = re.sub(r"''([^']*?)''", r"'\1'", fixed_formula)
        
        return fixed_formula
    
    def _create_limit_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path):
        """創建深度限制節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
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
            'has_indirect': False,
            'has_index': False
        }
    
    def _create_circular_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path):
        """創建循環引用節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
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
            'has_indirect': False,
            'has_index': False
        }
    
    def _create_error_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, error_msg):
        """創建錯誤節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
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
            'has_indirect': False,
            'has_index': False
        }
    
    def _create_range_node(self, range_info, current_depth, root_workbook_path):
        """創建範圍節點"""
        workbook_path = range_info['workbook_path']
        sheet_name = range_info['sheet_name']
        range_address = range_info['address']
        
        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
        if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
            filename = os.path.basename(workbook_path)
            if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                filename = filename.rsplit('.', 1)[0]
            display_address = f"[{filename}]{sheet_name}!{range_address}"
            short_display_address = display_address
            full_display_address = f"'{os.path.dirname(workbook_path)}\\[{os.path.basename(workbook_path)}]{sheet_name}'!{range_address}"
        else:
            display_address = f"{sheet_name}!{range_address}"
            short_display_address = display_address
            full_display_address = display_address
        
        rows = range_info.get('rows', 0)
        columns = range_info.get('columns', 0)
        hash_short = range_info.get('hash_short', 'N/A')
        range_value = f"{rows}Rx{columns}C | Hash: {hash_short}"
        
        return {
            'address': display_address,
            'short_address': short_display_address,
            'full_address': full_display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': range_address,
            'value': range_value,
            'calculated_value': range_value,
            'formula': None,
            'type': 'range',
            'children': [],
            'depth': current_depth,
            'error': range_info.get('error'),
            'has_indirect': False,
            'has_index': False,
            'range_info': {
                'dimensions': {
                    'rows': rows,
                    'columns': columns,
                    'total_cells': range_info.get('total_cells', 0),
                    'dimension_summary': f"{rows}行 x {columns}列"
                },
                'hash': {
                    'full_hash': range_info.get('hash', 'N/A'),
                    'short_hash': hash_short,
                    'content_summary': range_info.get('content_summary', '無內容摘要')
                }
            }
        }
    
    def _get_display_address(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path):
        """獲取顯示地址"""
        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
        if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
            filename = os.path.basename(workbook_path)
            if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                filename = filename.rsplit('.', 1)[0]
            return f"[{filename}]{sheet_name}!{cell_address}"
        else:
            return f"{sheet_name}!{cell_address}"
    
    def _count_nodes(self, node):
        """計算節點總數"""
        count = 1
        for child in node.get('children', []):
            count += self._count_nodes(child)
        return count
    
    def _get_max_depth(self, node):
        """獲取最大深度"""
        if not node.get('children'):
            return node.get('depth', 0)
        return max(self._get_max_depth(child) for child in node['children'])
    
    def get_explosion_summary(self, root_node):
        """獲取爆炸分析摘要 - 支援INDEX統計"""
        def count_nodes(node):
            count = 1
            for child in node.get('children', []):
                count += count_nodes(child)
            return count
        
        def get_max_depth(node):
            if not node.get('children'):
                return node.get('depth', 0)
            return max(get_max_depth(child) for child in node['children'])
        
        def count_by_type(node, type_counts=None):
            if type_counts is None:
                type_counts = {}
            
            node_type = node.get('type', 'unknown')
            type_counts[node_type] = type_counts.get(node_type, 0) + 1
            
            for child in node.get('children', []):
                count_by_type(child, type_counts)
            
            return type_counts
        
        def count_dynamic_function_nodes(node, dynamic_stats=None):
            if dynamic_stats is None:
                dynamic_stats = {
                    'total_indirect_nodes': 0,
                    'successful_indirect_resolutions': 0,
                    'failed_indirect_resolutions': 0,
                    'internal_references': 0,
                    'indirect_resolved_references': 0,
                    'total_index_nodes': 0,
                    'successful_index_resolutions': 0,
                    'failed_index_resolutions': 0,
                    'index_resolved_references': 0,
                    'index_internal_references': 0
                }
            
            # INDIRECT 統計
            if node.get('has_indirect'):
                dynamic_stats['total_indirect_nodes'] += 1
                if node.get('indirect_details'):
                    dynamic_stats['successful_indirect_resolutions'] += 1
                    dynamic_stats['internal_references'] += node.get('internal_references_count', 0)
                else:
                    dynamic_stats['failed_indirect_resolutions'] += 1
            
            if node.get('from_indirect_resolved'):
                dynamic_stats['indirect_resolved_references'] += 1
            
            # INDEX 統計
            if node.get('has_index'):
                dynamic_stats['total_index_nodes'] += 1
                if node.get('index_details'):
                    dynamic_stats['successful_index_resolutions'] += 1
                    dynamic_stats['index_internal_references'] += node.get('index_internal_references_count', 0)
                else:
                    dynamic_stats['failed_index_resolutions'] += 1
            
            if node.get('from_index_resolved'):
                dynamic_stats['index_resolved_references'] += 1
            
            for child in node.get('children', []):
                count_dynamic_function_nodes(child, dynamic_stats)
            
            return dynamic_stats
        
        basic_stats = {
            'total_nodes': count_nodes(root_node),
            'max_depth': get_max_depth(root_node),
            'type_distribution': count_by_type(root_node),
            'circular_references': len(self.circular_refs),
            'circular_ref_list': self.circular_refs
        }
        
        dynamic_stats = count_dynamic_function_nodes(root_node)
        
        return {
            **basic_stats,
            'dynamic_function_resolution': dynamic_stats,
            'indirect_resolution_log': self.indirect_resolution_log,
            'index_resolution_log': self.index_resolution_log,
            'our_instances_count': len(self.our_excel_instances)
        }


def explode_cell_dependencies_with_progress(workbook_path, sheet_name, cell_address, max_depth=10, range_expand_threshold=5, progress_callback=None):
    """
    便捷函數：爆炸分析指定儲存格的依賴關係 - 超安全版本 + INDEX支援 (完整版本)
    """
    exploder = EnhancedDependencyExploder(max_depth=max_depth, range_expand_threshold=range_expand_threshold, progress_callback=progress_callback)
    
    try:
        # 執行分析
        dependency_tree = exploder.explode_dependencies(workbook_path, sheet_name, cell_address)
        summary = exploder.get_explosion_summary(dependency_tree)
        
        # 分析完成後超安全清理
        if progress_callback:
            progress_callback.update_progress("[FINAL] 分析完成，超安全清理中...")
        
        exploder._ultra_safe_cleanup()
        
        if progress_callback:
            progress_callback.update_progress("[FINAL] ✓ 分析完成，您的Excel檔案完全不受影響")
        
        return dependency_tree, summary
        
    except Exception as e:
        # 異常時也要超安全清理
        try:
            exploder._ultra_safe_cleanup()
        except:
            pass
        raise e
    finally:
        # 最終確保清理
        try:
            exploder._ultra_safe_cleanup()
            del exploder
            import gc
            gc.collect()
        except:
            pass


# 測試函數
if __name__ == "__main__":
    test_workbook = r"C:\Users\user\Desktop\pytest\test.xlsx"
    test_sheet = "Sheet1" 
    test_cell = "A1"
    
    try:
        print("=== 超安全版本測試（完全不會影響用戶Excel）+ INDEX支援 (完整版本) ===")
        print(f"測試文件: {test_workbook}")
        print(f"測試位置: {test_sheet}!{test_cell}")
        print()
        
        progress = ProgressCallback()
        
        tree, summary = explode_cell_dependencies_with_progress(
            test_workbook, test_sheet, test_cell, 
            progress_callback=progress
        )
        
        print("依賴樹:")
        print(tree)
        print("\n摘要:")
        print(summary)
        
        # 顯示INDEX解析統計
        if 'index_resolution_log' in summary:
            print(f"\n INDEX 解析統計:")
            print(f"  - 成功解析: {summary['dynamic_function_resolution']['successful_index_resolutions']}")
            print(f"  - 失敗解析: {summary['dynamic_function_resolution']['failed_index_resolutions']}")
        
    except Exception as e:
        print(f"測試失敗: {e}")
        import traceback
        traceback.print_exc()
                    