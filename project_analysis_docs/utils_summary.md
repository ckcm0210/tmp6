# `utils` 資料夾模組深度分析報告 (詳細版 - 更新於 2025年8月12日)

## 引言

`utils` (Utilities) 資料夾是本專案的「引擎室」，它包含了專案中最核心、最複雜的底層演算法和輔助工具。經過近期重構，`utils` 內部已形成一個更清晰、更穩健的「**引擎-解析器-服務**」架構模式。

---

## 核心分析引擎 (Engine-Solver Pattern)

這是專案的心臟，採用了「策略模式 (Strategy Pattern)」的設計思想。由一個總引擎根據遇到的函式類型，選擇不同的解析器（策略）來處理。

### `progress_enhanced_exploder.py` (總引擎)
*   **職責**: 作為**分析流程的總協調器 (Orchestrator)**。它負責遞迴地遍歷依賴關係樹，當遇到需要動態解析的函式時，將任務**委派**給對應的解析器模組。
*   **主要函式**: 
    *   `explode_dependencies(...)`: 核心的遞迴函式，負責遍歷依賴鏈。
    *   `_create_node_with_dynamic_functions(...)`: 節點工廠，負責組裝分析結果的資料節點。

### `index_solver.py`, `vlookup_solver.py`, `hlookup_solver.py`, `indirect_solver.py` (解析器)
*   **職責**: **[新增模組]** 這一系列 `..._solver.py` 模組是專門的**函式解析器**。每一個都封裝了針對特定 Excel 動態函式的解析邏輯。
*   **主要函式**: 每個模組都包含一個 `resolve_...` 或類似的函式，接收公式字串和上下文，返回解析後的靜態引用。
*   **模組互動**: 它們都被 `progress_enhanced_exploder` 呼叫。在需要時，它們會透過統一的 `excel_com_manager` 來安全地執行 Excel 計算。

---

## Excel 連接與 I/O 服務

這一層負責所有與 Excel 檔案的直接互動，並對上層模組隱藏了複雜的細節。

### `excel_com_manager.py`
*   **職責**: **[核心安全服務]** 專案中**唯一**直接與 `win32com` 互動的模組，是一個**安全的 COM 連接管理器**。
*   **核心功能**:
    *   `_open_workbook_for_calculation(...)`: 使用 `DispatchEx` 建立一個完全隔離的、隱藏的 Excel 程序來執行計算。
    *   `_ultra_safe_cleanup()`: 在操作完成後，執行多階段清理，包括正常關閉、強制垃圾回收，以及最後根據 PID **終止**可能殘留的「殭屍程序」，從根本上解決檔案鎖定問題。

### `openpyxl_resolver.py`
*   **職責**: 一個「**增強版**」的 `.xlsx` 檔案讀取器，採用了「**裝飾器模式 (Decorator Pattern)**」。它包裝了 `openpyxl` 的物件，在不侵入原始碼的情況下，透明地將 `openpyxl` 返回的 `=[1]Sheet1!A1` 這類索引式公式，自動解析為帶有完整檔案路徑的公式。
*   **主要函式**: `load_resolved_workbook(...)` 是其對外的統一入口。

### `excel_io.py`
*   **職責**: 一個向下相容的檔案讀取器，主要使用 `xlrd` 函式庫來處理舊版的 `.xls` 檔案格式。

---

## 輔助工具與優化

### `safe_cache.py`
*   **職責**: 一個高效能、執行緒安全的**記憶體快取系統**，採用「**單例模式 (Singleton Pattern)**」確保全域唯一。
*   **核心功能**: 主要用於快取 `openpyxl` 載入的工作簿物件，避免重複的磁碟 I/O。它實現了 **LRU (最久未使用) 淘汰策略**，並能在檔案被外部修改時**自動讓快取失效**，設計非常完善。

### `excel_helpers.py`
*   **職責**: 提供高層次的、與 UI 互動緊密的輔助函式。
*   **主要函式**: `replace_links_in_excel`，封裝了在 Excel 中進行連結批量替換的完整流程。
*   **設計建議**: `replace_links_in_excel` 函式接收了 18 個參數，是典型的「程式碼壞味道」，應作為**高優先級的技術債**進行重構。

### `range_processor.py` & `range_optimizer.py`
*   **職責**: 一組處理 Excel **範圍 (Range)** 的工具。
    *   `range_processor`: 負責從公式中提取範圍，並計算其維度與內容雜湊值。
    *   `range_optimizer`: 提供了將離散的儲存格地址列表，優化為簡潔表達式（如 `A1:C5`）的演算法。

### `dependency_converter.py`
*   **職責**: 一個**資料格式轉換器**。它接收分析引擎產生的樹狀資料，並將其轉換為「節點列表」和「邊列表」，以供 `core/graph_generator.py` 用來生成視覺化圖表。

---

## 已廢棄模組

以下模組的功能已被新的「引擎-解析器」架構完全取代，應在下一步的清理計畫中**予以刪除**：

*   `dependency_exploder.py` (舊版引擎)
*   `workbook_cache.py` (已被 `safe_cache.py` 取代)
*   `core_indirect_resolver.py`, `indirect_processor.py`, `pure_indirect_logic.py`, `simple_indirect_resolver.py` (所有舊的 `INDIRECT` 解析嘗試)
*   `helpers.py`, `excel_utils.py` (內容為空或未被使用)