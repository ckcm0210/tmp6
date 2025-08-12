# `ui` 資料夾模組深度分析報告 (詳細版 - 更新於 2025年8月12日)

## 引言

`ui` (User Interface) 資料夾是本專案的門面，它包含了所有使用者可見、可互動的視窗、按鈕和版面配置。本報告旨在詳細闡述 UI 層的架構，幫助開發者理解各個視覺元件的職責與它們之間的協作方式。

---

## UI 架構概覽

本專案的 UI 架構可以分為三大類：

1.  **主應用程式視圖 (Main Views)**: 構成應用程式主視窗的頂層分頁，如「工作區」和「檢查模式」。
2.  **核心可複用元件 (Core Component)**: 即「公式分析面板」，它被設計為一個可重用的模組，在「正常模式」和「檢查模式」中都被使用。
3.  **獨立功能視窗 (Standalone Windows)**: 為特定功能（如深入分析、摘要、視覺化）彈出的獨立視窗。

---

## 模組詳解

### 主應用程式視圖

#### `workspace_view.py`
*   **職責**: 實現「**工作區 (Workspace)**」分頁，用於顯示和管理當前開啟的 Excel 檔案。
*   **主要類別**: 
    *   `Workspace`: 整個頁面的主類別，處理所有 UI 和邏輯。
    *   `AccumulateListbox`: 自訂的 `tk.Listbox`，支援更流暢的拖曳選取。
*   **核心功能**: 
    *   `show_names()`: 在獨立執行緒中呼叫 `get_open_excel_files()` 以非同步方式刷新開啟的檔案列表，避免 UI 凍結。
    *   `save_workspace()` / `load_workspace()`: 將當前開啟的檔案路徑儲存到 `.xlsx` 檔案中，或從中載入。
*   **設計建議**: `Workspace` 類別的職責較多，可考慮將 `win32com` 相關的批次操作（儲存、關閉等）邏輯提取到一個新的輔助模組中，以降低此 UI 類別的複雜度。

#### `modes/inspect_mode.py`
*   **職責**: 實現「**檢查模式 (Inspect Mode)**」，提供一個簡化的單窗格或雙窗格介面。
*   **主要類別**: 
    *   `InspectModeView`: 搭建雙窗格佈局的頂層 UI。
    *   `SimplifiedWorksheetController`: 繼承自標準的 `WorksheetController`，是此模式的核心。
*   **設計模式**: 採用「**繼承 (Inheritance)**」模式來複用程式碼。`SimplifiedWorksheetController` 在繼承父類別所有功能的基礎上，以程式化的方式**隱藏**了在檢查模式下不需要的 UI 元件。
*   **設計建議**: 「繼承後再隱藏」的作法雖然能快速開發，但耦合度較高。若未來基礎 `WorksheetView` 的版面變動較大，此處的隱藏邏輯可能會失效。更穩健的作法是「**組合優於繼承**」，可考慮建立一個不含 UI 的 `BaseWorksheetController`，讓兩個模式的控制器都去繼承它，並各自獨立地建立自己所需的 UI。

---

### 核心可複用元件：公式分析面板 (`worksheet/`)

這是專案中最核心的 UI 元件，採用了 MVC (模型-視圖-控制器) 的變體設計模式，職責劃分清晰。

#### `worksheet/controller.py` (控制器)
*   **職責**: `WorksheetController` 是單個分析面板的「**大腦 (Controller)**」。它不包含任何 UI 元件，而是負責儲存該面板的所有**狀態和資料**（如 `self.all_formulas` 列表）以及 UI 狀態變數（如 `tk.BooleanVar`）。

#### `worksheet/view.py` & `worksheet_ui.py` (視圖)
*   **職責**: 這兩個檔案共同構建了分析面板的「**視覺介面 (View)**」，採用了「**UI 工廠 (UI Factory)**」的設計模式。
    *   `view.py` (`WorksheetView`): 作為一個 `ttk.Frame` 容器，定義了面板的整體框架，但自身不建立具體元件。
    *   `worksheet_ui.py`: 作為「工廠」，提供 `create_ui_widgets` 和 `bind_ui_commands` 兩個函式，負責實際建立所有按鈕、標籤、Treeview 等，並將其事件直接綁定到 `core` 層的邏輯函式。
*   **設計建議**: 這種將 UI 的「建立」和「版面」分離的作法非常清晰，使得 `WorksheetView` 保持乾淨，而將複雜的 UI 建立細節封裝在 `worksheet_ui.py` 中。

#### `worksheet/tab_manager.py` (子元件管理器)
*   **職責**: `TabManager` 是一個專門的管理器，負責控制面板下方「詳細資訊」區域的**多分頁介面** (`ttk.Notebook`)。它提供了優秀的使用者體驗，支援右鍵、雙擊、中鍵等多種方式關閉分頁。

---

### 獨立功能視窗

#### `dependency_exploder_view.py`
*   **職責**: 實現「**深入分析 (Explode)**」功能的彈出視窗 (`tk.Toplevel`)。它是一個功能完備的獨立分析單元，擁有自己的控制項、進度顯示和結果樹，並直接呼叫 `utils.progress_enhanced_exploder` 來獲取分析資料。

#### `summary_window.py`
*   **職責**: 實現「**摘要外部連結**」功能的彈出視窗，並內嵌了「**取代工具 (Replace Tool)**」。
*   **設計建議**: `__init__` 建構函式承擔了過多職責（UI 建立、資料處理、快取建立）。應將其拆分為 `_setup_ui`, `_process_data`, `_bind_events` 等多個更小的私有方法。此外，「執行取代」按鈕的 `lambda` 函式傳遞了過多參數，顯示出高耦合，建議將這些參數封裝到一個物件中傳遞。

#### `visualizer.py`
*   **職責**: 實現「**視覺化圖表**」的彈出視窗。它由 `summary_window.py` 呼叫，使用 `matplotlib` 函式庫繪製一個模擬的 Excel 網格，並在圖上用「熱圖」的形式標示出受選定連結影響的所有儲存格。