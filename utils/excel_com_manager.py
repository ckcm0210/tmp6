# -*- coding: utf-8 -*-
"""
Excel COM Manager - 超安全版本
專門負責Excel COM連接的安全管理
從 progress_enhanced_exploder.py 中提取的Excel管理邏輯
Last Updated: 2025-01-16
"""

import re
import os
import win32com.client
import pythoncom
import time
import psutil
import datetime
import gc
import traceback
import hashlib
import uuid
import tempfile
import shutil

class ExcelComManager:
    """超安全版Excel COM管理器 - 完全避免檔案鎖定問題"""
    
    def __init__(self, progress_callback=None):
        # 記錄創建的 Excel 實例和 PID
        self.our_excel_instances = {}
        self.excel_process_pids = set()  # 記錄我們創建的 Excel 程序 PID
        self.progress_callback = progress_callback
        
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
        
        if self.progress_callback:
            self.progress_callback.update_progress("[ULTRA-SAFE] 開始超安全清理...")
        
        # 第一階段：正常清理 COM 物件
        instance_keys = list(self.our_excel_instances.keys())
        for instance_key in instance_keys:
            try:
                self._cleanup_single_instance(instance_key)
            except Exception as e:
                if self.progress_callback:
                    self.progress_callback.update_progress(f"[ULTRA-SAFE] 清理實例失敗: {e}")
        
        # 清空記錄
        self.our_excel_instances.clear()
        
        # 第二階段：強制垃圾回收
        if self.progress_callback:
            self.progress_callback.update_progress("[ULTRA-SAFE] 執行強制垃圾回收...")
        for i in range(5):
            gc.collect()
            time.sleep(0.2)
        
        # 第三階段：檢查並終止我們創建的 Excel 程序
        if self.progress_callback:
            self.progress_callback.update_progress("[ULTRA-SAFE] 檢查殘留的 Excel 程序...")
        remaining_pids = self._check_and_terminate_our_excel_processes()
        
        if remaining_pids and self.progress_callback:
            self.progress_callback.update_progress(f"[ULTRA-SAFE] 已終止 {len(remaining_pids)} 個殘留的 Excel 程序")
        
        # 第四階段：等待檔案系統釋放
        if self.progress_callback:
            self.progress_callback.update_progress("[ULTRA-SAFE] 等待檔案系統完全釋放...")
        time.sleep(1.0)  # 給檔案系統更多時間釋放鎖定
        
        # 重置內部狀態
        self.excel_process_pids.clear()
        
        if self.progress_callback:
            self.progress_callback.update_progress("[ULTRA-SAFE] ✓ 超安全清理完成，檔案已完全釋放")
    
    def _cleanup_single_instance(self, instance_key):
        """清理單個 Excel 實例"""
        if instance_key not in self.our_excel_instances:
            return
        
        instance_info = self.our_excel_instances[instance_key]
        excel_app = instance_info.get('app')
        wb = instance_info.get('workbook')
        instance_id = instance_info.get('instance_id', 'Unknown')
        
        try:
            if self.progress_callback:
                self.progress_callback.update_progress(f"[ULTRA-SAFE] 開始清理實例: {instance_id}")
            
            # 1. 關閉工作簿
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                    if self.progress_callback:
                        self.progress_callback.update_progress(f"[ULTRA-SAFE] 工作簿已關閉: {instance_id}")
                except Exception as close_error:
                    if self.progress_callback:
                        self.progress_callback.update_progress(f"[ULTRA-SAFE] 關閉工作簿失敗: {close_error}")
                finally:
                    # 刪除臨時複本
                    try:
                        temp_dir = instance_info.get('temp_dir')
                        if temp_dir and os.path.exists(temp_dir):
                            shutil.rmtree(temp_dir, ignore_errors=True)
                            if self.progress_callback:
                                self.progress_callback.update_progress(f"[ULTRA-SAFE] 已刪除臨時資料夾: {temp_dir}")
                    except Exception as _:
                        pass
                
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
                    if self.progress_callback:
                        self.progress_callback.update_progress(f"[ULTRA-SAFE] Excel 應用程式已退出: {instance_id}")
                    
                except Exception as quit_error:
                    if self.progress_callback:
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
            if self.progress_callback:
                self.progress_callback.update_progress(f"[ULTRA-SAFE] 清理實例異常: {e}")
    
    def _check_and_terminate_our_excel_processes(self):
        """檢查並終止我們創建的 Excel 程序"""
        terminated_pids = []
        
        for pid in list(self.excel_process_pids):
            try:
                if psutil.pid_exists(pid):
                    proc = psutil.Process(pid)
                    if proc.name().lower() == 'excel.exe':
                        if self.progress_callback:
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
                if self.progress_callback:
                    self.progress_callback.update_progress(f"[ULTRA-SAFE] 終止程序 {pid} 失敗: {e}")
        
        return terminated_pids
    
    def open_workbook_for_calculation(self, workbook_path):
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
            if self.progress_callback:
                self.progress_callback.update_progress(f"[ULTRA-SAFE] 創建完全隔離的 Excel 實例: {os.path.basename(workbook_path)}")
            
            # 使用原始檔案路徑（暫時還原，不使用臨時複本）
            temp_dir = None
            temp_path = workbook_path
            
            # 使用 DispatchEx 創建完全獨立的實例
            excel_app = win32com.client.DispatchEx("Excel.Application")
            
            # 記錄新創建的 Excel 程序
            time.sleep(0.5)  # 給程序啟動一點時間
            new_pids = self._get_new_excel_processes(before_pids)
            self.excel_process_pids.update(new_pids)
            if new_pids and self.progress_callback:
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
            
            if self.progress_callback:
                self.progress_callback.update_progress(f"[ULTRA-SAFE] 隔離設置完成，開啟工作簿...")
            
            # 以唯讀模式開啟臨時複本，使用更嚴格的參數
            wb = excel_app.Workbooks.Open(
                temp_path,
                UpdateLinks=0,
                ReadOnly=True,
                Format=5,
                Password="",
                WriteResPassword="",
                IgnoreReadOnlyRecommended=True,
                Origin=1,
                Delimiter="",
                Editable=False,
                Notify=False,
                Converter=0,
                AddToMru=False,
                Local=False,
                CorruptLoad=0
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
                'process_pids': new_pids.copy(),
                'temp_dir': temp_dir,
                'temp_path': temp_path
            }
            
            self.our_excel_instances[instance_key] = instance_info
            
            if self.progress_callback:
                self.progress_callback.update_progress(f"[ULTRA-SAFE] 隔離實例創建成功: {wb.Name}")
            return instance_info
            
        except Exception as e:
            if self.progress_callback:
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
    
    def close_specific_instance(self, instance_key):
        """關閉特定實例 - 超安全版本"""
        if instance_key not in self.our_excel_instances:
            return
        
        if self.progress_callback:
            self.progress_callback.update_progress(f"[ULTRA-SAFE] 開始關閉實例: {instance_key}")
        
        try:
            self._cleanup_single_instance(instance_key)
            
            # 額外等待時間確保完全釋放
            time.sleep(0.5)
            
            if self.progress_callback:
                self.progress_callback.update_progress(f"[ULTRA-SAFE] 實例已安全關閉: {instance_key}")
            
        except Exception as e:
            if self.progress_callback:
                self.progress_callback.update_progress(f"[ULTRA-SAFE] 關閉實例失敗: {e}")
    
    def calculate_safely(self, indirect_content, workbook_path, sheet_name, cell_address):
        """安全計算 INDIRECT - 使用完全隔離的實例"""
        temp_instance = None
        temp_cell = None
        original_value = None
        original_formula = None
        
        try:
            if self.progress_callback:
                self.progress_callback.update_progress(f"[ULTRA-SAFE-CALC] 開始安全計算: {indirect_content}")
            
            # 驗證文件路徑
            if not os.path.exists(workbook_path):
                return {
                    'success': False,
                    'error': f'文件不存在: {workbook_path}',
                    'indirect_content': indirect_content
                }
            
            # 為計算創建臨時的完全隔離實例
            temp_instance = self.open_workbook_for_calculation(workbook_path)
            
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
                if self.progress_callback:
                    self.progress_callback.update_progress(f"[ULTRA-SAFE-CALC] 計算完成，結果: '{calculation_result}'")
                
            except Exception as calc_error:
                if self.progress_callback:
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
                
                # 還原原始內容
                try:
                    if original_formula:
                        temp_cell.Formula = original_formula
                    else:
                        temp_cell.Value = original_value
                except:
                    pass
            
            # 驗證結果
            if calculation_result is None:
                return {
                    'success': False,
                    'error': '計算結果為空',
                    'indirect_content': indirect_content
                }
            
            # 檢查是否為Excel錯誤
            if self._is_excel_error(calculation_result):
                return {
                    'success': False,
                    'error': f'Excel計算錯誤: {calculation_result}',
                    'indirect_content': indirect_content
                }
            
            # 轉換結果為字符串
            static_reference = str(calculation_result).strip()
            
            if self.progress_callback:
                self.progress_callback.update_progress(f"[ULTRA-SAFE-CALC] ✓ 計算成功: '{static_reference}'")
            
            return {
                'success': True,
                'static_reference': static_reference,
                'calculation_result': calculation_result,
                'indirect_content': indirect_content
            }
            
        except Exception as e:
            if self.progress_callback:
                self.progress_callback.update_progress(f"[ULTRA-SAFE-CALC] 計算異常: {e}")
            return {
                'success': False,
                'error': str(e),
                'indirect_content': indirect_content
            }
        
        finally:
            # 確保清理臨時實例
            if temp_instance:
                try:
                    self.close_specific_instance(temp_instance['instance_id'])
                except Exception as cleanup_error:
                    if self.progress_callback:
                        self.progress_callback.update_progress(f"[ULTRA-SAFE-CALC] 清理臨時實例失敗: {cleanup_error}")
    
    def _is_excel_error(self, result):
        """檢查是否為Excel錯誤值"""
        if result is None:
            return True
        
        error_values = [
            '#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#GETTING_DATA'
        ]
        
        result_str = str(result).upper()
        return any(error in result_str for error in error_values)