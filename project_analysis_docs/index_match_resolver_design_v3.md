# `INDEX(MATCH)` 解析功能 - 技術設計方案 (v3.0)

## 1. 目標 (Objective)

本方案旨在擴充現有的 `progress_enhanced_exploder.py` 模組，賦予其解析 `INDEX` 及 `INDEX(MATCH)` 這類「計算型引用」公式的能力。最終目標是將這類公式還原為一個靜態的、單一的儲存格地址（例如 `'[data.xlsx]Sheet1'!I5`），以便我們的依賴關係分析引擎可以繼續追蹤其後續的依賴鏈路。

---

## 2. 核心原理與執行流程 (Core Principle & Workflow)

我們將採用一個穩健且高效的解析策略，以應對公式的複雜性。

1.  **關鍵字掃描 (Keyword Scanning)**: 首先，將整個公式作為一個字串，快速掃描其中是否存在 `INDEX(` 這個關鍵字。這一步是為了避免對無關的公式進行昂貴的解析操作。
2.  **偵測與提取 (Detection & Extraction)**: **只有在**偵測到關鍵字後，才啟動一個更複雜的、基於括號配對的演算法，來精準地找到並提取出最內層的 `INDEX(...)` 函式的完整內容。
3.  **遞迴解析參數 (Recursive Parameter Resolution)**: 對提取出的 `INDEX` 函式的三個核心參數 (`array`, `row_num`, `column_num`) 進行解析。
    *   如果參數是 `MATCH`, `OFFSET` 等需要計算的函式，我們將重用專案中已有的「安全 COM 計算引擎」(`_calculate_indirect_safely`) 來獲得它們的數值結果。
    *   這個過程本身也需要是遞迴的，以處理例如 `INDEX(..., INDEX(...))` 的情況。
4.  **計算最終地址 (Final Address Calculation)**: 根據 `array` 參數的起始儲存格，以及計算出的行列**偏移量**，精確定位到最終指向的單一儲存格地址。
5.  **靜態替換 (Static Replacement)**: 將原始公式中的整個 `INDEX(...)` 部分，替換為我們計算出的靜態地址，並將這個靜態地址作為新的依賴節點，交由後續的分析引擎處理。

---

## 3. 應用情境與示範 (Scenarios & Demonstrations)

### 情境一：簡單 `INDEX` (純數字偏移)
*   **公式範例**: `=INDEX(C3:Z100, 3, 7)`
*   **解析結果**: `I5`

### 情境二：`INDEX` 配合簡單 `MATCH`
*   **公式範例**: `=INDEX(A1:A100, MATCH("Apple", B1:B100, 0))`
*   **解析思路**: 首先安全計算 `MATCH(...)` 的值（假設為10），然後計算出 `INDEX` 的結果。
*   **解析結果**: `A10`

### 情境三：複雜 `INDEX(MATCH, MATCH)` 內嵌於 VLOOKUP (真實世界情境)
*   **公式範例**: `=VLOOKUP(A1, 'C:\Users\user\Documents\Financial Reports\2025\[Q3_Forecast_Internal.xlsx]Assumptions'!$A:$D, INDEX({1,2,3,4}, 0, MATCH(B1, 'C:\Users\user\Documents\Financial Reports\2025\[Q3_Forecast_Internal.xlsx]Assumptions'!$A$1:$D$1, 0)), FALSE)`
*   **解析思路**:
    1.  我們的解析器會首先偵測到 `INDEX` 函式。
    2.  **解析 `array`**: 識別出 `array` 是一個常數陣列 `{1,2,3,4}`。
    3.  **解析 `row_num`**: 識別出 `row_num` 是 `0`。
    4.  **解析 `column_num`**: 識別出 `column_num` 是一個 `MATCH` 函式。
    5.  **安全計算 `MATCH`**: 讀取當前檔案 `B1` 儲存格的值（假設為 "Q2"），然後呼叫安全 COM 計算引擎，計算 `MATCH("Q2", 'C:\...\Assumptions'!$A$1:$D$1, 0)`，假設得到結果為 `3`。
    6.  **計算 `INDEX` 結果**: 計算 `INDEX({1,2,3,4}, 0, 3)`，得到結果 `3`。
    7.  **靜態替換**: 將原始公式中的 `INDEX(...)` 部分替換為計算結果 `3`。
    8.  **最終公式 (用於後續分析)**: `=VLOOKUP(A1, 'C:\...\Assumptions'!$A:$D, 3, FALSE)`。這樣，我們就成功消除了一個動態引用。

---

## 4. 潛在風險、盲點與技術邊界 (Potential Risks, Blind Spots, and Technical Boundaries)

在實作前，我們必須意識到此功能的潛在挑戰：

1.  **陣列結果 (Array Results)**: `INDEX` 函式在 `row_num` 或 `column_num` 為 0 或省略時，可以返回一整個行或列。我們 v1.0 的實作將**只專注於解析返回單一儲存格地址的情況**。對於返回陣列的情況，我們將其標記為「無法解析的 `INDEX` 範圍」，並在日誌中記錄，留待未來作為增強功能進行開發。

2.  **揮發性函式 (Volatile Functions)**: 如果 `MATCH` 的參數中包含 `NOW()`, `TODAY()`, `RAND()` 等函式，我們的解析結果雖然在分析當下是準確的，但它不是一個穩定的依賴。我們需要接受這個限制，並在報告中註明我們解析的是「分析時間點」的快照。

3.  **效能權衡 (Performance Trade-off)**: 每一次對 `MATCH` 或 `OFFSET` 的解析都意味著一次 COM 呼叫和一個隱藏的 Excel 計算過程。對於包含大量此類巢狀函式的公式，解析時間會相應增加。這是在用**執行時間**換取**解析的準確性**，是一個必須接受的權衡。

---

## 5. 程式碼實現與模組化 (Code Implementation & Modularization)

根據您的建議，我們將此複雜邏輯完全封裝在一個獨立的新檔案中，以保證主引擎的整潔和可維護性。

*   **新檔案**: `utils/index_resolver.py`
*   **核心類別**: `IndexResolver`

```python
# 檔案: utils/index_resolver.py

class IndexResolver:
    def __init__(self, exploder_context):
        self.exploder = exploder_context # 接收主分析器的實例，以呼叫其安全計算等輔助功能

    def resolve_formula(self, formula):
        # ... 實現上述的核心原理與執行流程 ...
        pass

    # ... 其他私有輔助函式 ...
```

---

## 6. 交接備忘與明日工作計畫 (Handover Memo & Next Day's Work Plan)

*   **今日已完成 (Completed Today):**
    1.  修正了 `core/worksheet_tree.py` 的循環匯入致命錯誤。
    2.  對 `core`, `ui`, `utils` 三個資料夾下的所有 действующий的檔案，都進行了 `import` 修正和逐一的深度分析。
    3.  產出了三份極度詳盡的、無佔位符的子資料夾分析報告。
    4.  共同制定並敲定了這份關於 `INDEX(MATCH)` 新功能的 v3.0 版技術設計方案。

*   **待確認 (Pending Confirmation):**
    *   請您最後審閱本 v3.0 設計方案，確認其中的解析流程、風險分析和下一步計畫是否完全符合您的預期。

*   **明日工作起點 (Starting Point for Tomorrow):**
    *   一旦您批准此方案，我們明日的第一個任務將是：在 `utils/` 目錄下建立 `index_resolver.py` 檔案，並開始編寫 `IndexResolver` 這個類別的 `__init__` 和 `resolve_formula` 主方法，正式將此設計付諸實現。
