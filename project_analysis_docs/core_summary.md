# `core` 資料夾模組深度分析報告 (詳細版 - 更新於 2025年8月12日)

## 引言

`core` 套件是本應用程式的業務邏輯中樞。經過重構，其內部職責劃分更加清晰，主要圍繞著一個「協調者-管理員」的模式運作。本報告旨在詳細闡述每個模組的內部結構、主要功能、互動方式及改進建議。

---

## 核心架構：協調者與管理員

### `worksheet_tree.py` (核心協調者)
*   **總體功用**: 作為 UI 事件的**中心分派器**。它不再直接處理複雜的業務邏輯，而是接收來自 `ui.worksheet_ui` 的事件（如點擊、排序），並將任務**委派**給專門的管理器。
*   **主要函式**:
    *   `apply_filter(...)`: 根據 UI 篩選條件，過濾 `self.all_formulas` 中的資料並刷新 Treeview。
    *   `sort_column(...)`: 處理 Treeview 的欄位排序邏輯。
    *   `on_select(...)` / `on_double_click(...)`: 將對應事件直接轉發給 `details_manager` 處理。
*   **模組互動**: 它是 `ui` 層和 `core` 層之間的關鍵橋樑。它呼叫 `utils.progress_enhanced_exploder` 來啟動分析，並將結果儲存在 `self.nodes` 中，供其他模組使用。
*   **設計建議**: 目前的重構非常成功，職責清晰。可考慮將 `self.nodes` 等核心資料的管理也移交給一個專門的「資料管理器」，讓 `worksheet_tree` 成為一個更純粹的事件協調者。

### `details_manager.py` (詳細資訊管理器)
*   **總體功用**: **[新增模組]** 處理所有與「顯示詳細資料」和「深入分析」相關的邏輯。
*   **主要函式**:
    *   `on_select(...)`: 接收 `worksheet_tree` 傳來的事件，從 Treeview 中提取節點資料，並在 UI 下方的 `TabManager` 中顯示該節點的詳細公式和值。
    *   `on_double_click(...)`: 接收雙擊事件，實例化 `ui.dependency_exploder_view`，從而彈出「深入分析」視窗。
*   **模組互動**: 從 `worksheet_tree` 接收指令，呼叫 `ui.dependency_exploder_view` 建立新視窗。
*   **設計建議**: 職責單一，設計良好。

### `navigation_manager.py` (導航管理器)
*   **總體功用**: **[新增模組]** 專門負責「跳轉到 Excel 中對應儲存格」的功能。
*   **主要函式**:
    *   `go_to_reference(...)`: 接收目標工作簿、工作表和儲存格地址，呼叫 `excel_connector` 中的函式來啟用 Excel 視窗並選定儲存格。
*   **模組互動**: 被 `details_manager` 和 `dependency_exploder_view` 等多個模組呼叫，是實現「跳轉」功能的統一出口。它依賴於 `excel_connector` 來執行底層操作。
*   **設計建議**: 實現了關注點分離，是優秀的重構範例。

---

## 底層服務與核心功能

### `excel_connector.py` (Excel 連接器)
*   **總體功用**: 作為與底層 `utils.excel_com_manager` 溝通的**唯一正式橋樑**，採用了「外觀模式 (Facade Pattern)」。它為上層模組提供了一個更乾淨、更安全的介面來操作 Excel。
*   **主要函式**:
    *   `reconnect_to_excel(...)`: 重新連接到 Excel 實例。
    *   `activate_excel_window(...)`: 將指定的 Excel 視窗帶到前景。
*   **設計建議**: 模組中的錯誤處理（`except Exception as e:`）過於寬泛，應改為捕捉更具體的 `pywintypes.com_error`，以便進行更精細的錯誤處理。

### `excel_scanner.py` (Excel 掃描器)
*   **總體功用**: 執行核心的**公式掃描**功能。
*   **主要函式**:
    *   `refresh_data(...)`: 協調整個掃描流程，包括連接、確定範圍、呼叫 `_get_formulas_from_excel`，並處理進度回饋。
    *   `_get_formulas_from_excel(...)`: 使用 `worksheet.UsedRange.SpecialCells(constants.xlCellTypeFormulas)` 這個高效的 COM 方法來一次性獲取所有包含公式的儲存格，避免了逐行遍歷，效能很高。
*   **設計建議**: `refresh_data` 函式非常長，且有多層巢狀的 `try...except`，是重構的主要候選者。應將其分解為 `_prepare_scan`, `_execute_scan`, `_update_ui_with_results` 等多個更小的私有函式以提高可讀性。

### `graph_generator.py` (圖表產生器)
*   **總體功用**: 將分析後的樹狀資料，轉換為一個獨立的、可互動的 **HTML 視覺化圖表**。
*   **主要函式**:
    *   `generate_graph()`: 協調整個流程，呼叫 `_generate_standalone_html` 來產生 HTML 內容，並寫入檔案。
    *   `_generate_standalone_html()`: 將一個巨大的、包含完整 JS 函式庫的 HTML 範本字串與節點資料結合。
*   **設計建議**: 將巨大的 JavaScript 和 CSS 字串直接嵌入 Python 程式碼中，使得前端邏輯極難維護。強烈建議將 JS 和 CSS 分離到獨立的 `.js` 和 `.css` 檔案中，Python 只負責讀取範本並注入數據。此外，如 `TODO.md` 中所述，此模組需要更新以支援 `vlookup` 等新的解析類型。

---

## 輔助模組

### `worksheet_export.py` & `worksheet_summary.py`
*   **總體功用**: 分別提供「匯出到 Excel」和「摘要外部連結」的功能。它們是相對獨立的功能模組，作為 `ui.worksheet_ui` 中對應按鈕的後端邏輯。

### `formula_classifier.py` & `link_analyzer.py`
*   **總體功用**: 兩個模組共同組成了**公式解析引擎**。`link_analyzer` 使用正規表示式從公式字串中提取引用，而 `formula_classifier` 則使用其結果來判斷公式的總體類型（外部連結、內部連結等）。
*   **設計建議**: `link_analyzer` 中的 `get_referenced_cell_values` 函式在每次呼叫時都會重新編譯多個正規表示式。應將這些模式移到模組級別作為常數，以提高效能。

### `mode_manager.py`
*   **總體功用**: 一個簡單的**狀態管理器**，採用了「觀察者模式 (Observer Pattern)」，允許其他模組註冊回呼函式，以便在模式變更時接收通知。

### `data_processor.py` & `formula_comparator.py`
*   **總體功用**: 這兩個模組與 `ui` 層的 `dual_pane_controller.py` 共同構建了「**公式比較器**」分頁的完整功能。

---

## 已廢棄模組

### `models.py` & `worksheet_refresh.py`
*   **狀態**: **[應刪除]** 根據 `Comprehensive_Analysis_Report.md` 的分析，這兩個檔案已不再被任何活躍的程式碼使用，應在下一步的清理計畫中**予以刪除**。
