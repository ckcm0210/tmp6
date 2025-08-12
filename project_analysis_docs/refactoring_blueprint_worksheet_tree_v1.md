### **重構藍圖與分析報告**

#### **第一部分：`core/worksheet_tree.py` 職責分析**

首先，我們來深入探討 `core/worksheet_tree.py` 這個檔案。經過分析，我發現它目前混合了至少 **五種** 完全不同的功能類別，這也是它變得如此龐大的根本原因：

1.  **數據管理與顯示 (Data Management & Display)**
    *   **功能**: 負責主視窗中公式列表 (Treeview) 的數據處理。
    *   **對應函數**: `apply_filter()`, `sort_column()`。
    *   **問題**: 這部分與數據的呈現方式高度相關，但與具體的業務邏輯（如導航、分析）關係不大。

2.  **UI 事件處理 (UI Event Handling)**
    *   **功能**: 處理使用者在 Treeview 上的各種互動。
    *   **對應函數**: `on_select()`, `on_double_click()`。
    *   **問題**: 這些是典型的 UI 控制器邏輯，負責響應使用者操作。

3.  **Excel 應用程式導航 (Excel Application Navigation)**
    *   **功能**: 處理所有需要直接操作 Excel 應用程式（透過 `win32com`）來跳轉、選取儲存格的任務。
    *   **對應函數**: `go_to_reference()`, `on_double_click()` (部分邏輯), `go_to_reference_new_tab()`。
    *   **問題**: 這是與外部應用程式通訊的底層邏輯，應該被封裝起來，而不是散落在 UI 處理代碼中。

4.  **詳細資訊面板管理 (Details Panel Management)**
    *   **功能**: 當使用者點擊某個公式時，在下方的「詳細資訊」頁籤中生成並顯示該公式的完整資訊、其引用的所有儲存格，並為每個引用綁定新的「Go to Reference」按鈕。
    *   **對應函數**: `on_select()` (大部分邏輯), `go_to_reference_new_tab()`。
    *   **問題**: 這部分包含了複雜的 UI 生成和業務邏輯，與 `on_select` 這個單純的事件處理函數耦合過深。

5.  **「依賴關係爆炸」彈出視窗 (Dependency Exploder Popup)**
    *   **功能**: 創建一個全新的、功能完整的彈出視窗，用於顯示和互動一個儲存格的完整依賴關係樹。
    *   **對應函數**: `explode_dependencies_popup()`。
    *   **問題**: 這本身就是一個獨立的、複雜的 UI 元件，包含了大量的 UI 佈局代碼、事件處理和狀態管理。將其放在 `worksheet_tree.py` 中，是造成該檔案臃腫的最主要原因。

#### **第二部分：建議的新架構與職責劃分**

基於以上分析，我建議將 `core/worksheet_tree.py` 的功能拆分到以下幾個新的、職責單一的模組中：

| 新/修改的檔案 | 預計職責 |
| :--- | :--- |
| **`core/worksheet_tree.py`** (重構後) | **(精簡)** 只專注於主視窗公式列表 (Treeview) 的**數據管理**和**基礎事件綁定**。它將作為一個輕量級的協調者。 |
| **`ui/dependency_exploder_view.py`** (新) | **(UI 元件)** 完整封裝「依賴關係爆炸」彈出視窗的所有 UI 和互動邏輯。 |
| **`core/navigation_manager.py`** (新) | **(核心服務)** 封裝所有與 Excel 應用程式互動的底層**導航操作**，提供簡單的介面給上層呼叫。 |
| **`core/details_manager.py`** (新) | **(核心服務)** 專門負責生成「詳細資訊」面板中的內容，包括解析引用、創建按鈕等。 |

#### **第三部分：詳細的程式碼遷移計畫**

以下是 `core/worksheet_tree.py` 中每個主要函數的詳細遷移路徑：

| 函數/邏輯 | 當前職責 | 建議遷移目標 | 理由 |
| :--- | :--- | :--- | :--- |
| `apply_filter()` | 根據篩選條件更新 Treeview | **保留**在 `core/worksheet_tree.py` | 這是該模組的核心職責：管理 Treeview 的數據呈現。 |
| `sort_column()` | 對 Treeview 的列進行排序 | **保留**在 `core/worksheet_tree.py` | 同上，屬於 Treeview 的數據管理。 |
| `on_double_click()` | 處理雙擊事件，跳轉到 Excel | **保留**在 `core/worksheet_tree.py` | 事件綁定本身保留，但其內部的導航邏輯將呼叫 `navigation_manager`。 |
| `on_select()` | 處理單擊事件，生成詳細資訊 | **保留**在 `core/worksheet_tree.py` | 事件綁定保留，但其內部的 UI 生成邏輯將呼叫 `details_manager`。 |
| `go_to_reference()` | 操作 `win32com` 跳轉到指定儲存格 | **遷移**到 `core/navigation_manager.py` | 這是純粹的 Excel 導航功能，應集中管理。 |
| `go_to_reference_new_tab()` | 導航並創建新頁籤 | **遷移**到 `core/navigation_manager.py` | 同上，屬於導航功能。新頁籤的創建可以透過回呼(callback)或與 `TabManager` 互動來完成。 |
| `read_reference_openpyxl()` | 使用 `openpyxl` 讀取儲存格 | **遷移**到 `core/navigation_manager.py` | 雖然不是 `win32com`，但它本質上是「獲取遠端儲存格資訊」的導航/讀取類操作，適合放在一起。 |
| `explode_dependencies_popup()` | 創建一個完整的彈出視窗 | **遷移**到 `ui/dependency_exploder_view.py` | 這是最重要的一步。將這個巨大的 UI 元件完全獨立出去，封裝成一個類別。 |

#### **第四部分：重構後各檔案的角色**

*   **`core/worksheet_tree.py` (新角色: 輕量級協調者)**
    *   **唯一職責**: 管理主視窗公式列表的數據（過濾/排序）和基本事件（點擊/雙擊）。
    *   當事件發生時，它不再自己處理複雜邏輯，而是**委派**任務給其他專業模組。例如：`on_double_click` -> 呼叫 `navigation_manager.go_to_cell()`；`on_select` -> 呼叫 `details_manager.populate_details()`。

*   **`ui/dependency_exploder_view.py` (新角色: 獨立的 UI 元件)**
    *   **唯一職責**: 負責「依賴關係爆炸」視窗的創建、佈局、事件處理和銷毀。它是一個可以被任何地方實例化並使用的獨立元件。

*   **`core/navigation_manager.py` (新角色: Excel 導航專家)**
    *   **唯一職責**: 處理所有與「跳轉到」或「讀取」遠端 Excel 儲存格相關的底層操作。它隱藏了 `win32com` 和 `openpyxl` 的複雜細節，向上層提供乾淨、簡單的 API，如 `navigate_to(workbook, sheet, cell)`。

*   **`core/details_manager.py` (新角色: 詳細資訊產生器)**
    *   **唯一職責**: 接收一個公式的數據，解析其所有引用，並產生一個可用於顯示在 Text 元件中的、包含各種按鈕和格式的內容。它不關心事件是從哪裡來的，只負責內容的生成。
