# 功能規格書：互動式顏色編碼依賴鏈追蹤

## 1. 總體目標

本功能旨在為 Excel 依賴關係圖引入一個全新的視覺化維度：**互動式顏色編碼追蹤**。其核心目標是，當使用者點擊一個節點時，系統會動態地為該節點公式（或其解析結果）中的每一個獨立「靜態引用」，分配一個獨一無二的背景高亮顏色，並將這個顏色應用到該引用所指向的子節點的「地址」上。

這將徹底改變使用者與圖表的互動方式，將其從一個靜態的關係展示圖，提升為一個具有深度分析能力的「殺手級」互動式儀表板，讓使用者可以極其直觀地追蹤任何一條複雜的依賴鏈。

## 2. 功能詳解與視覺化範例

### 2.1. 互動模式：點擊觸發 (On-Click Trigger)

為保持介面預設的簡潔性，此功能並非預設開啟。

*   **預設狀態：** 圖表載入時，所有節點的文字均不帶有任何特殊的背景高亮。
*   **觸發狀態：** 當使用者點擊任一節點（父節點）時，系統會即時計算並顯示與該節點相關的顏色依賴鏈。
*   **恢復狀態：** 當使用者點擊畫布空白處，或選擇另一個節點時，先前顯示的顏色高亮將會消失。

### 2.2. 場景一：標準靜態引用 (來自 Formula)

此場景展示了當一個節點的 `Formula` 直接包含多個靜態引用時的理想效果。

*   **父節點 (P)**
    *   **Address:** `Sheet1!D10`
    *   **Formula:** `=IF([data.xlsx]Input!$B$2>0, Sheet1!C5, 0)`

*   **子節點 (C1)**
    *   **Address:** `[data.xlsx]Input!$B$2`

*   **子節點 (C2)**
    *   **Address:** `Sheet1!C5`

**期望的視覺化效果 (當使用者點擊節點 P 時)：**

---
**節點 P (父)**
```
Address : Sheet1!D10

Formula : =IF([background: #FFB3BA][data.xlsx]Input!$B$2[/background]>0, [background: #BAE1FF]Sheet1!C5[/background], 0)

Value   : 123.45
```
---
**節點 C1 (子)**
```
Address : [background: #FFB3BA][data.xlsx]Input!$B$2[/background]

Formula : 50

Value   : 50
```
---
**節點 C2 (子)**
```
Address : [background: #BAE1FF]Sheet1!C5[/background]

Formula : =A5+B5

Value   : 123.45
```
---

### 2.3. 場景二：進階動態與遞歸追蹤 (來自 Resolved 且鏈接延續)

此場景展示了當依賴關係由動態函數產生，並且**該依賴鏈需要繼續向下延伸**時的理想效果。

*   **父節點 (P_Dynamic)**
    *   **Address:** `Sheet1!E11`
    *   **Formula:** `=INDIRECT("Sheet3!" & "A" & B1)`
    *   **Resolved:** `Sheet3!A5`

*   **子節點 (C_Target)**
    *   **Address:** `Sheet3!A5`
    *   **Formula:** `=SUM(X5, Y5)`

*   **孫節點 (C_Grandchild_1)**
    *   **Address:** `X5`

*   **孫節點 (C_Grandchild_2)**
    *   **Address:** `Y5`

**期望的視覺化效果 (當使用者點擊節點 P_Dynamic 時)：**

---
**節點 P_Dynamic (父)**
```
Address : Sheet1!E11

Formula : =INDIRECT("Sheet3!" & "A" & B1)

Resolved: [background: #BAFFC9]Sheet3!A5[/background]

Value   : 999
```
---
**節點 C_Target (子)**
```
Address : [background: #BAFFC9]Sheet3!A5[/background]

Formula : =SUM(X5, Y5)

Value   : 999
```
---
**解讀：** 點擊 `P_Dynamic` 後，其 `Resolved` 結果 `Sheet3!A5` 被高亮為**淡綠色**，並成功傳遞到子節點 `C_Target` 的 `Address` 上。`C_Target` 的 `Formula` 此時不顯示高亮，因為追蹤的起點是 `P_Dynamic`。

**期望的視覺化效果 (當使用者點擊節點 C_Target 時)：**

---
**節點 P_Dynamic (父)**
```
Address : Sheet1!E11

Formula : =INDIRECT("Sheet3!" & "A" & B1)

Resolved: Sheet3!A5

Value   : 999
```
---
**節點 C_Target (子)**
```
Address : [background: #BAFFC9]Sheet3!A5[/background]

Formula : =SUM([background: #FFFFBA]X5[/background], [background: #E0BBE4]Y5[/background])

Value   : 999
```
---
**節點 C_Grandchild_1 (孫)**
```
Address : [background: #FFFFBA]X5[/background]

Formula : 400

Value   : 400
```
---
**節點 C_Grandchild_2 (孫)**
```
Address : [background: #E0BBE4]Y5[/background]

Formula : 599

Value   : 599
```
---
**解讀：** 當使用者轉而點擊 `C_Target` 時，`C_Target` 成為新的「焦點」。它 `Address` 上的淡綠色高亮依然保留（因為這是它的身份），同時其 `Formula` 中的 `X5` 和 `Y5` 分別被賦予了新的**淡黃色**和**淡紫色**高亮，並成功傳遞到對應的孫節點上，完美體現了功能的**遞歸性**。

## 3. 設計與技術規範

### 3.1. 顏色使用規範

為確保功能的可用性和視覺清晰度，顏色的選用需遵循以下原則：

*   **高對比度：** 應選用一組對比度高、易於人眼區分的基礎色調。
*   **推薦色板：**
    *   淡紅色 (`#FFB3BA`)
    *   橙色 (`#FFDFBA`)
    *   淡黃色 (`#FFFFBA`)
    *   淡綠色 (`#BAFFC9`)
    *   天藍色 (`#BAE1FF`)
    *   淡紫色 (`#E0BBE4`)
*   **避免事項：** 應極力避免使用難以區分的、飽和度或亮度相近的顏色（例如：同時使用淺藍、天藍、寶藍），以免對使用者造成視覺混淆。

### 3.2. 技術實現簡述

*   **後端 (Python):**
    1.  需建立一個在圖表生成期間持續存在的「顏色註冊中心」（如字典）。
    2.  此中心負責為全域唯一的靜態引用字串（無論來自 `Formula` 還是 `Resolved`）分配並儲存一個顏色。
    3.  節點資料結構需升級，將 `Address`、`Formula`、`Resolved` 從單純的字串，改為能攜帶 `text` 和 `highlight_color` 屬性的 `parts` 陣列。

*   **前端 (JavaScript):**
    1.  `drawNode` 函式需重構，使其能根據「焦點節點」的狀態，來決定是否渲染 `highlight_color` 指定的背景色。
    2.  節點的點擊事件處理器 (`handleNodeHighlight`) 需增強，以管理「焦點節點」狀態、尋找關聯節點，並觸發全局重繪。
