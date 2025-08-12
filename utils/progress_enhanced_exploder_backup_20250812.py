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
from utils.excel_com_manager import ExcelComManager
from utils.indirect_solver import IndirectSolver
from utils.index_solver import IndexSolver
from utils.vlookup_solver import VLookupSolver
from utils.hlookup_solver import HLookupSolver
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
        self.indirect_resolution_log = []
        self.index_resolution_log = []  # 新增：INDEX解析日誌
        
        # 創建專門的管理器實例 - 修改：使用拆分的模組
        self.excel_manager = ExcelComManager(self.progress_callback)
        self.index_solver = IndexSolver(self.excel_manager, self.progress_callback, self)
        self.vlookup_solver = VLookupSolver(self.excel_manager, self.progress_callback, self)
        self.hlookup_solver = HLookupSolver(self.excel_manager, self.progress_callback, self)
        self.indirect_solver = IndirectSolver(self.excel_manager, self.progress_callback, self)
        
        # 初始化 COM
        try:
            pythoncom.CoInitialize()
        except:
            pass
    
    def __del__(self):
        """析構函數：超安全清理"""
        try:
            self.excel_manager._ultra_safe_cleanup()  # 修改：使用excel_manager
        except:
            pass
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    # === 跳過已拆分的重複方法，直接從 explode_dependencies 開始 ===
    
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
            fixed_formula = self._clean_formula(original_formula) if original_formula else None
            resolved_formula = fixed_formula
            indirect_info = None
            index_info = None
            vlookup_info = None
            hlookup_info = None
            
            if fixed_formula and fixed_formula.startswith('='):
                # INDIRECT 處理 - 修改：使用拆分的模組
                if 'INDIRECT' in fixed_formula.upper():
                    self.progress_callback.update_progress(f"正在解析INDIRECT函數: {current_ref}")
                    try:
                        resolved_result = self.indirect_solver._resolve_indirect_with_excel(
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
                
                # INDEX/VLOOKUP 處理 - 修改：使用拆分的模組
                if 'INDEX(' in (resolved_formula or fixed_formula).upper():
                    self.progress_callback.update_progress(f"正在解析INDEX函數: {current_ref}")
                    try:
                        index_result = self.index_solver._resolve_index_with_excel_corrected_simple(
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
                            
                            # 記錄解析日誌
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
            
            # VLOOKUP/HLOOKUP 處理 - 新增：使用拆分的模組（安全判斷）
            _formula_text = resolved_formula if resolved_formula else fixed_formula
            if _formula_text and _formula_text.startswith('=') and ('VLOOKUP(' in _formula_text.upper() or 'HLOOKUP(' in _formula_text.upper()):
                self.progress_callback.update_progress(f"正在解析查找函數 (VLOOKUP/HLOOKUP): {current_ref}")
                try:
                    # 先嘗試 VLOOKUP
                    vres = None
                    if 'VLOOKUP(' in _formula_text.upper():
                        vres = self.vlookup_solver.resolve_vlookup(resolved_formula or fixed_formula, workbook_path, sheet_name, cell_address)
                        if vres and vres.get('success'):
                            resolved_formula = vres.get('resolved_formula', resolved_formula)
                            vlookup_info = {
                                'has_vlookup': True,
                                'success': True,
                                'resolved_formula': resolved_formula,
                                'details': vres,
                                'internal_references': vres.get('internal_references', [])
                            }
                            self.progress_callback.update_progress(f"VLOOKUP解析完成，resolved: {resolved_formula}")
                            _formula_text = resolved_formula if resolved_formula else fixed_formula
                        elif vres:
                            vlookup_info = {
                                'has_vlookup': True,
                                'success': False,
                                'error': vres.get('error', 'Unknown error'),
                                'internal_references': vres.get('internal_references', [])
                            }
                            self.progress_callback.update_progress(f"VLOOKUP解析失敗: {vlookup_info['error']}")
                    
                    # 再嘗試 HLOOKUP
                    hres = None
                    if 'HLOOKUP(' in _formula_text.upper():
                        hres = self.hlookup_solver.resolve_hlookup(resolved_formula or fixed_formula, workbook_path, sheet_name, cell_address)
                        if hres and hres.get('success'):
                            resolved_formula = hres.get('resolved_formula', resolved_formula)
                            hlookup_info = {
                                'has_hlookup': True,
                                'success': True,
                                'resolved_formula': resolved_formula,
                                'details': hres,
                                'internal_references': hres.get('internal_references', [])
                            }
                            self.progress_callback.update_progress(f"HLOOKUP解析完成，resolved: {resolved_formula}")
                        elif hres:
                            hlookup_info = {
                                'has_hlookup': True,
                                'success': False,
                                'error': hres.get('error', 'Unknown error'),
                                'internal_references': hres.get('internal_references', [])
                            }
                            self.progress_callback.update_progress(f"HLOOKUP解析失敗: {hlookup_info['error']}")
                except Exception as e:
                    vlookup_info = {
                        'has_vlookup': True,
                        'success': False,
                        'error': str(e),
                        'internal_references': []
                    }
                    self.progress_callback.update_progress(f"VLOOKUP解析異常: {str(e)}")
            
            # 創建節點
            return self._create_node_with_dynamic_functions(
                workbook_path, sheet_name, cell_address, current_depth, root_workbook_path,
                cell_info, fixed_formula, resolved_formula, indirect_info, index_info, vlookup_info, hlookup_info
            )
            
        except Exception as e:
            error_msg = f"處理過程中發生錯誤: {str(e)}"
            self.progress_callback.update_progress(f"錯誤：{error_msg}")
            return self._create_error_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, error_msg)
        
        finally:
            # 從已訪問集合中移除，允許其他路徑重新訪問
            self.visited_cells.discard(cell_id)

    def _create_node_with_dynamic_functions(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, cell_info, fixed_formula, resolved_formula=None, indirect_info=None, index_info=None, vlookup_info=None, hlookup_info=None):
        """創建支持動態函數的節點"""
        filename = os.path.basename(workbook_path)
        dir_path = os.path.dirname(workbook_path)
        
        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
        
        if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
            short_display_address = f"[{filename}]{sheet_name}!{cell_address}"
            full_display_address = f"'{dir_path.replace(chr(92), '/')}/[{filename}]{sheet_name}'!{cell_address}"
            display_address = short_display_address
        else:
            short_display_address = f"[{filename}]{sheet_name}!{cell_address}"
            full_display_address = f"'{dir_path.replace(chr(92), '/')}/[{filename}]{sheet_name}'!{cell_address}"
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
        
        # INDIRECT/VLOOKUP/INDEX 信息
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
        
        # VLOOKUP/HLOOKUP 信息
        if vlookup_info and vlookup_info.get('has_vlookup'):
            node['has_vlookup'] = True
            if vlookup_info.get('success'):
                has_dynamic_resolution = True
                node['vlookup_details'] = vlookup_info.get('details')
                node['vlookup_internal_references_count'] = len(vlookup_info.get('internal_references', []))
            else:
                node['vlookup_error'] = vlookup_info.get('error')
        else:
            node['has_vlookup'] = False
        
        if hlookup_info and hlookup_info.get('has_hlookup'):
            node['has_hlookup'] = True
            if hlookup_info.get('success'):
                has_dynamic_resolution = True
                node['hlookup_details'] = hlookup_info.get('details')
                node['hlookup_internal_references_count'] = len(hlookup_info.get('internal_references', []))
            else:
                node['hlookup_error'] = hlookup_info.get('error')
        else:
            node['has_hlookup'] = False
        if vlookup_info and vlookup_info.get('has_vlookup'):
            node['has_vlookup'] = True
            if vlookup_info.get('success'):
                has_dynamic_resolution = True
                node['vlookup_details'] = vlookup_info.get('details')
                node['vlookup_internal_references_count'] = len(vlookup_info.get('internal_references', []))
            else:
                node['vlookup_error'] = vlookup_info.get('error')
        else:
            node['has_vlookup'] = False
        
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
        
        # 解析公式引用並創建子節點
        # Gemini修復 (2025-08-12): 確保使用解析後的公式(resolved_formula)進行下一步分析，而不是原始的fixed_formula
        formula_to_parse = resolved_formula if resolved_formula and has_dynamic_resolution else fixed_formula
        if formula_to_parse and formula_to_parse.startswith('='):
            try:
                references = self._parse_formula_references_accurate(formula_to_parse, workbook_path, sheet_name)
                for ref in references:
                    # 特別處理：若是範圍摘要（大於閾值的range），不要遞迴到單一儲存格，直接建立範圍節點
                    if ref.get('is_range_summary'):
                        try:
                            rp_info = range_processor.process_range(
                                ref.get('workbook_path', workbook_path),
                                ref.get('sheet_name', sheet_name),
                                ref.get('cell_address', '')
                            )
                            # range_processor 返回的資料中已包含 rows/columns/hash 等欄位
                            range_node = self._create_range_node(rp_info, current_depth + 1, root_workbook_path)
                            node['children'].append(range_node)
                            continue
                        except Exception as re:
                            self.progress_callback.update_progress(f"範圍摘要節點建立失敗: {str(re)}")
                            error_node = self._create_error_node(
                                ref.get('workbook_path', workbook_path),
                                ref.get('sheet_name', sheet_name),
                                ref.get('cell_address', ''),
                                current_depth + 1,
                                root_workbook_path,
                                str(re)
                            )
                            node['children'].append(error_node)
                            continue

                    child_node = self.explode_dependencies(
                        ref['workbook_path'], ref['sheet_name'], ref['cell_address'],
                        current_depth + 1, root_workbook_path or workbook_path
                    )
                    if child_node:
                        # 標記此子節點是來自於一個動態解析的結果
                        if has_dynamic_resolution:
                            if node.get('has_indirect'):
                                child_node['from_indirect_resolved'] = True
                            if node.get('has_index'):
                                child_node['from_index_resolved'] = True
                        node['children'].append(child_node)
            except Exception as e:
                self.progress_callback.update_progress(f"解析引用時發生錯誤: {str(e)}")
        
        # 處理範圍地址（恢復 1800 行版本的行為：基於公式/解析後公式偵測範圍，生成範圍節點，含維度與雜湊）
        try:
            formula_for_ranges = resolved_formula if resolved_formula else cell_info.get('formula')
            ranges = process_formula_ranges(formula_for_ranges, workbook_path, sheet_name) if formula_for_ranges else []
            if ranges:
                self.progress_callback.update_progress(f"找到 {len(ranges)} 個範圍，正在處理...")
                for i, range_info in enumerate(ranges, 1):
                    try:
                        range_display = f"{os.path.basename(range_info['workbook_path'])}!{range_info['sheet_name']}!{range_info['address']}"
                        self.progress_callback.update_progress(f"處理範圍 {i}/{len(ranges)}: {range_display}")
                        range_node = self._create_range_node(range_info, current_depth + 1, root_workbook_path)
                        node['children'].append(range_node)
                    except Exception as e:
                        self.progress_callback.update_progress(f"錯誤：處理範圍失敗 {range_display} - {str(e)}")
                        error_node = self._create_error_node(
                            range_info.get('workbook_path', workbook_path),
                            range_info.get('sheet_name', sheet_name),
                            range_info.get('address', ''),
                            current_depth + 1,
                            root_workbook_path,
                            str(e)
                        )
                        node['children'].append(error_node)
        except Exception as e:
            self.progress_callback.update_progress(f"警告：處理範圍時發生異常 - {str(e)}")
        
        return node

    def force_cleanup(self):
        """公開的超安全清理方法"""
        self.progress_callback.update_progress("[USER] 用戶觸發超安全清理...")
        self.excel_manager._ultra_safe_cleanup()  # 修改：使用excel_manager
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
                        range_info = self._process_range_reference(cell_ref, full_file_path, sheet_name, 'external')
                        if range_info:
                            references.extend(range_info)
                    else:
                        references.append({
                            'workbook_path': full_file_path,
                            'sheet_name': sheet_name,
                            'cell_address': cell_ref,
                            'ref_type': 'external'
                        })
                    
                    add_processed_span(start, end)

                elif m_type in ['local_quoted', 'local_unquoted']:
                    sheet_name, cell_ref = match.groups()
                    
                    if ':' in cell_ref:
                        range_info = self._process_range_reference(cell_ref, current_workbook_path, sheet_name, 'local')
                        if range_info:
                            references.extend(range_info)
                    else:
                        references.append({
                            'workbook_path': current_workbook_path,
                            'sheet_name': sheet_name,
                            'cell_address': cell_ref,
                            'ref_type': 'local'
                        })
                    
                    add_processed_span(start, end)

                elif m_type in ['current_range', 'current_single']:
                    cell_ref = match.group(1)
                    
                    if ':' in cell_ref:
                        range_info = self._process_range_reference(cell_ref, current_workbook_path, current_sheet_name, 'current')
                        if range_info:
                            references.extend(range_info)
                    else:
                        references.append({
                            'workbook_path': current_workbook_path,
                            'sheet_name': current_sheet_name,
                            'cell_address': cell_ref,
                            'ref_type': 'current'
                        })
                    
                    add_processed_span(start, end)

            except Exception as e:
                continue

        return references

    def _process_range_reference(self, range_ref, workbook_path, sheet_name, ref_type):
        """處理範圍引用"""
        try:
            range_size = self._calculate_range_size(range_ref)
            
            if range_size <= self.range_expand_threshold:
                # 展開為個別儲存格
                return self._expand_range_to_cells(range_ref, workbook_path, sheet_name, ref_type)
            else:
                # 創建範圍摘要
                return [self._create_range_summary(range_ref, workbook_path, sheet_name, ref_type, range_size)]
        except Exception as e:
            return []

    def _calculate_range_size(self, range_ref):
        """計算範圍包含的儲存格數量"""
        try:
            clean_range = range_ref.replace('$', '').strip()
            if ':' not in clean_range:
                return 1
            
            start_cell, end_cell = clean_range.split(':')
            start_col, start_row = self._parse_cell_address(start_cell.strip())
            end_col, end_row = self._parse_cell_address(end_cell.strip())
            
            col_count = abs(end_col - start_col) + 1
            row_count = abs(end_row - start_row) + 1
            
            return col_count * row_count
        except Exception:
            return 1

    def _parse_cell_address(self, cell_address):
        """解析儲存格地址為列號和行號"""
        try:
            clean_address = cell_address.replace('$', '').strip()
            
            # 分離字母和數字
            col_letters = ''
            row_number = ''
            
            for char in clean_address:
                if char.isalpha():
                    col_letters += char
                elif char.isdigit():
                    row_number += char
            
            # 轉換列字母為數字
            col_num = 0
            for char in col_letters.upper():
                col_num = col_num * 26 + (ord(char) - ord('A') + 1)
            
            return col_num, int(row_number)
        except Exception:
            return 1, 1

    def _expand_range_to_cells(self, range_ref, workbook_path, sheet_name, ref_type):
        """將範圍展開為個別儲存格引用"""
        try:
            clean_range = range_ref.replace('$', '').strip()
            start_cell, end_cell = clean_range.split(':')
            
            start_col, start_row = self._parse_cell_address(start_cell.strip())
            end_col, end_row = self._parse_cell_address(end_cell.strip())
            
            cells = []
            for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
                for col in range(min(start_col, end_col), max(start_col, end_col) + 1):
                    col_letter = self._col_num_to_letters(col)
                    cell_address = f"{col_letter}{row}"
                    cells.append({
                        'workbook_path': workbook_path,
                        'sheet_name': sheet_name,
                        'cell_address': cell_address,
                        'ref_type': ref_type
                    })
            
            return cells
        except Exception as e:
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
        filename = os.path.basename(workbook_path)
        
        if ref_type == 'external':
            display_address = f"[{filename}]{sheet_name}!{range_ref}"
        elif ref_type == 'local':
            display_address = f"[{filename}]{sheet_name}!{range_ref}"
        else:  # current
            display_address = f"{sheet_name}!{range_ref}"
        
        return {
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': range_ref,
            'ref_type': f'{ref_type}_range',
            'is_range_summary': True,
            'range_info': f'範圍包含 {cell_count} 個儲存格',
            'display_address': display_address,
            'cell_count': cell_count
        }

    def _clean_formula(self, formula):
        """清理公式，移除不必要的字符和格式"""
        if not formula:
            return formula
        
        # 移除前後空白
        cleaned = formula.strip()
        
        # 標準化引號
        cleaned = cleaned.replace('"', '"').replace('"', '"')
        
        return cleaned

    def _create_limit_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path):
        """創建深度限制節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
        return {
            'address': display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'value': f'[達到最大深度限制: {self.max_depth}]',
            'formula': f'[達到最大深度限制: {self.max_depth}]',
            'type': 'limit',
            'children': [],
            'depth': current_depth,
            'error': None,
            'is_limit': True
        }

    def _create_circular_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path):
        """創建循環引用節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
        return {
            'address': display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'value': '[循環引用]',
            'formula': '[循環引用]',
            'type': 'circular',
            'children': [],
            'depth': current_depth,
            'error': None,
            'is_circular': True
        }

    def _create_error_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, error_msg):
        """創建錯誤節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
        return {
            'address': display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'value': f'[錯誤: {error_msg}]',
            'formula': f'[錯誤: {error_msg}]',
            'type': 'error',
            'children': [],
            'depth': current_depth,
            'error': error_msg,
            'is_error': True
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
            'formula': None,
            'type': 'range',
            'children': [],
            'depth': current_depth,
            'error': None,
            'is_range': True,
            'range_info': range_info
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
        
        def count_dynamic_function_nodes(node, dynamic_stats=None):
            if dynamic_stats is None:
                dynamic_stats = {
                    'total_indirect_nodes': 0,
                    'successful_indirect_resolutions': 0,
                    'failed_indirect_resolutions': 0,
                    'internal_references': 0,
                    'indirect_resolved_references': 0,
                    'total_index_nodes': 0,
                    'total_vlookup_nodes': 0,
                    'total_hlookup_nodes': 0,
                    'successful_hlookup_resolutions': 0,
                    'failed_hlookup_resolutions': 0,
                    'hlookup_internal_references': 0,
                    'successful_vlookup_resolutions': 0,
                    'failed_vlookup_resolutions': 0,
                    'vlookup_internal_references': 0,
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
            
            # VLOOKUP/HLOOKUP 統計
            if node.get('has_vlookup'):
                dynamic_stats['total_vlookup_nodes'] += 1
                if node.get('vlookup_details'):
                    dynamic_stats['successful_vlookup_resolutions'] += 1
                    dynamic_stats['vlookup_internal_references'] += node.get('vlookup_internal_references_count', 0)
                else:
                    dynamic_stats['failed_vlookup_resolutions'] += 1

            if node.get('has_hlookup'):
                dynamic_stats['total_hlookup_nodes'] += 1
                if node.get('hlookup_details'):
                    dynamic_stats['successful_hlookup_resolutions'] += 1
                    dynamic_stats['hlookup_internal_references'] += node.get('hlookup_internal_references_count', 0)
                else:
                    dynamic_stats['failed_hlookup_resolutions'] += 1

            # INDEX 統計
            if node.get('has_vlookup'):
                dynamic_stats['total_vlookup_nodes'] += 1
                if node.get('vlookup_details'):
                    dynamic_stats['successful_vlookup_resolutions'] += 1
                    dynamic_stats['vlookup_internal_references'] += node.get('vlookup_internal_references_count', 0)
                else:
                    dynamic_stats['failed_vlookup_resolutions'] += 1

            # INDEX 統計
            if node.get('has_index'):
                dynamic_stats['total_index_nodes'] += 1
                if node.get('index_details'):
                    dynamic_stats['successful_index_resolutions'] += 1
                    dynamic_stats['index_internal_references'] += node.get('index_internal_references_count', 0)
                else:
                    dynamic_stats['failed_index_resolutions'] += 1
            
            for child in node.get('children', []):
                count_dynamic_function_nodes(child, dynamic_stats)
            
            return dynamic_stats
        
        total_nodes = count_nodes(root_node)
        max_depth = get_max_depth(root_node)
        type_counts = count_by_type(root_node)
        dynamic_stats = count_dynamic_function_nodes(root_node)
        
        return {
            'total_nodes': total_nodes,
            'max_depth': max_depth,
            'max_depth_reached': max_depth,
            'circular_references': len(self.circular_refs),
            'circular_ref_list': self.circular_refs,
            'our_instances_count': len(self.excel_manager.our_excel_instances),  # 修改：使用excel_manager
            'type_distribution': type_counts,
            'dynamic_function_stats': dynamic_stats
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
        
        exploder.excel_manager._ultra_safe_cleanup()  # 修改：使用excel_manager
        
        if progress_callback:
            progress_callback.update_progress("[FINAL] 分析完成，您的Excel檔案完全不受影響")
        
        return dependency_tree, summary
        
    except Exception as e:
        # 異常時也要超安全清理
        try:
            exploder.excel_manager._ultra_safe_cleanup()  # 修改：使用excel_manager
        except:
            pass
        raise e