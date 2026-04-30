# MyPyrevitTools (MyTab)

This repository contains multiple custom pyRevit tools, primarily divided into four main panels: **GreenBuilding**, **Data**, **Sheet**, and **Areas**. These panels provide functionalities for calculating window areas for green building daily energy saving indicators, bidirectional export and import of Revit parameters with external CSV spreadsheets, batch management of sheet title blocks, and automatic creation of spatial boundaries.

此資料夾包含多個自訂的 pyRevit 工具，主要分為四大面板：**GreenBuilding**、**Data**、**Sheet** 與 **Areas**，分別提供綠建築日常節能指標檢核的開窗面積計算、Revit 參數與外部 CSV 表單的雙向匯出與匯入回寫功能、圖紙圖框批次管理，以及空間邊界自動建立功能。

## 🛠️ 工具清單與功能說明 / Tool List and Features

### 1. WindowOrientation (外殼等價開窗面積計算 / Equivalent Window Area Calculation)
- **位置 / Location**: `MyTab` tab -> `GreenBuilding` panel
- **功能描述 / Description**:
  Automatically collects windows and curtain walls in the model, and calculates the corresponding azimuth angle and solar radiation correction factor (fk) based on regional climate, True North, or Project North. It calculates the "Envelope Equivalent Window Area (ΣAgi*fk*Ki)" for each window and exports it as a CSV report for Taiwan's Green Building Daily Energy Saving Indicator (Req) review.
  自動收集模型中的窗戶（Windows）與幕牆（Curtain Wall, DW），並依地區氣候、真北或專案北的方向，計算出對應的方位角及日射修正係數（fk）。最終計算出各窗戶的「外殼等價開窗面積（ΣAgi*fk*Ki）」，並匯出為 CSV 報表，方便進行台灣綠建築日常節能指標 (Req) 之檢核計算。
- **特色亮點 / Key Features**:
  - **Built-in Climate Data (內建氣候區數據)**: Supports region selection to automatically apply the corresponding solar radiation correction factor.
  - **Orientation Switch (方位切換)**: Choose between "True North" or "Project North" for accurate 16-direction window judgment.
  - **Curtain Wall Support (支援幕牆)**: Target windows, curtain walls, or both for calculation.
  - **Automated Reporting (自動輸出報表)**: Generates summary (`Req_D1_rows.csv`) and detailed (`Req_instances.csv`) reports.

---

### Data 面板工具 (CSV 資料匯出/匯入) / Data Panel Tools (CSV Export/Import)

#### 2. Import Sheet CSV (自 CSV 批次回寫圖框參數 / Batch Update Titleblock Parameters from CSV)
- **位置 / Location**: `MyTab` tab -> `Data` panel 
- **功能描述 / Description**:
  Updates parameters for Revit Sheets and Titleblocks from an external CSV file using a Unique Identifier for matching. It syncs the data back to the Revit project without relying on native Revit schedules.
  將外部編輯好的 CSV 表單資料，透過指定的**獨特碼 (Unique Identifier)** 進行比對，將整批資料回寫同步至 Revit 專案的「圖紙 (Views Sheet)」以及「圖紙上的圖框實體 (Title Block)」內部參數。
- **特色亮點 / Key Features**:
  - **Custom Unique Key (自訂辨識鍵值)**: Flexibly assign any CSV column as the matching key.
  - **High Encoding Compatibility (高相容性編碼讀取)**: Supports UTF-8, UTF-8-sig, and CP950/Big5 natively.
  - **Flexible Mapping (彈性對應建置)**: Users can map CSV columns to Revit parameters directly from the UI.
  - **Smart Parameter Mapping & Caching (智慧參數對應與快取)**: Caches one-time user inputs for unmatched CSV headers for future use.
  - **Safe Transaction (安全防護機制)**: Built-in rollback functionality to prevent data corruption on errors.

#### 3. Export Room CSV (匯出房間參數至 CSV / Export Room Parameters to CSV)
- **位置 / Location**: `MyTab` tab -> `Data` panel -> `RoomTools` group
- **功能描述 / Description**:
  Collects placed Rooms in the model, allows users to select parameters to export (number, name, area, finishes, etc.), and generates a UTF-8 CSV file for easy editing in Excel.
  自動收集模型中已放置的房間 (Rooms)，讓使用者勾選欲匯出的參數，並將資料整理匯出為 CSV。
- **特色亮點 / Key Features**:
  - **Auto Unit Conversion (自動轉換 Revit 單位)**: Retrieves formatted values shown in Revit UI instead of raw internal units.
  - **Smart Pre-selection (智慧預選)**: Memorizes and pre-selects commonly used room parameter fields.

#### 4. Import Room CSV (自 CSV 批次回寫房間參數 / Batch Update Room Parameters from CSV)
- **位置 / Location**: `MyTab` tab -> `Data` panel -> `RoomTools` group
- **功能描述 / Description**:
  Reads externally modified room CSV files and updates room parameters (department, name, finishes, etc.) in Revit using a specified matching key (default is room number).
  讀取外部修改好的房間 CSV 檔案，使用指定的「對應鍵」將修改後的資料回寫至 Revit 的房間參數中。
- **特色亮點 / Key Features**:
  - **Unit System Parsing (支援單位系統解析)**: Parses strings with units via Revit's UnitUtils to ensure correct internal values.
  - **Read-Only Protection (避免唯讀屬性錯誤)**: Safely ignores read-only properties (like Area) to ensure smooth execution.

---

### Sheet 面板工具 (圖紙/圖框管理) / Sheet Panel Tools (Sheet/Titleblock Management)

#### 5. Change TitleBlock (批次更換圖框 / Batch Change TitleBlock)
- **位置 / Location**: `MyTab` tab -> `Sheet` panel
- **功能描述 / Description**:
  Batch change the TitleBlock Type for selected sheets. Uses `ChangeTypeId()` to preserve TitleBlock position and shared parameter values.
  批次更換所選圖紙的圖框類型 (TitleBlock Type)。使用 `ChangeTypeId()` 保留圖框位置與共用參數值。
- **特色亮點 / Key Features**:
  - **Multi-select Sheets (多選圖紙)**: Built-in sheet selector supporting multi-selection and filtering.
  - **Pre-execution Preview (執行前預覽)**: Lists the current TitleBlock status of each sheet before execution.
  - **Safe Replacement (安全替換)**: Preserves element positions and parameters using `ChangeTypeId()`.
  - **Smart Skip (智慧跳過)**: Skips sheets that already have the target TitleBlock applied.

---

### Areas 面板工具 (空間邊界/尺度設定) / Areas Panel Tools (Space Boundaries/Dimensions)

#### 6. Boundary create (自動建立房間邊界與標註 / Auto-create Room Boundaries and Dimensions)
- **位置 / Location**: `MyTab` tab -> `Areas` panel 
- **功能描述 / Description**:
  Automatically selects Rooms in Revit and generates Detail Curves, Area Boundary Lines, and Length Dimensions in target plan views (like Area Plans) based on user requirements.
  自動選取 Revit 內的房間（Rooms），並依據需求在指定的目標平面視圖中建立對應的細部線、面積邊界線以及長度標註。
- **特色亮點 / Key Features**:
  - **Multi-functional Creation (多功能建立)**: Converts room geometry boundaries to detail or area boundary lines.
  - **Cross-view Processing (跨視圖處理)**: Supports different Levels and Target Views (prioritizing Area Plans).
  - **Error Handling & Prevention (預設防呆與錯誤處理)**: Accurately grabs boundaries via Revit API with duplicate dimension prevention.

## 📥 安裝與路徑設定 / Installation and Path Configuration
The tools are packaged as a `MyTools.extension` extension. Add the path of this folder via the pyRevit Settings Manager (Settings -> Custom Extension Directories) and Reload to show the `MyTab` tab in the top menu.
工具已經封裝為 `MyTools.extension` 擴充套件格式。可以透過 pyRevit 設定管理員將此資料夾的路徑加入，並重新載入 (Reload) 以在上方選單列顯示 `MyTab` 標籤。
