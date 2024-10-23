Excel 轉 JSON 工具
介紹
這是一個使用 Python 開發的圖形化應用程式，用來將 Excel 檔案中的數據轉換為 JSON 格式。使用者可以選擇 Excel 檔案中的多個 Sheet，並將其各自轉換為獨立的 JSON 檔案。

功能
選擇 Excel 文件（支援 .xlsx 和 .xls 格式）
支援選擇多個 Sheet 進行轉換
轉換後的數據會以 JSON 格式保存
簡單的圖形使用者介面，便於操作
提供進度顯示及錯誤提示功能
依賴項目
在使用此工具前，請確保已安裝以下 Python 依賴包：

bash
複製程式碼
pip install pandas tqdm openpyxl
此外，tkinter 是 Python 標準庫的一部分，通常無需額外安裝。

使用方法
執行此 Python 腳本。
在彈出的視窗中點擊“選擇 Excel 文件”按鈕，選擇要轉換的 Excel 文件。
選擇要轉換的 Sheet。
選擇保存 JSON 文件的路徑，並點擊“確認”完成轉換。
轉換完成後，將彈出成功提示。
範例
假設有一個 Excel 文件 data.xlsx，包含兩個 Sheet：Sheet1 和 Sheet2。當你選擇該文件並選定 Sheet1 和 Sheet2 進行轉換，程序會在指定的路徑生成兩個 JSON 文件：

data_Sheet1.json
data_Sheet2.json
日誌
如在轉換過程中發生錯誤，程序會記錄錯誤並在日誌中顯示，提示使用者檢查轉換狀態。

版本
1.0.0

作者
此工具由 BuffaloMachinery 開發，歡迎使用並提出改進建議。

此 README 提供了一個簡單的使用說明，涵蓋了依賴安裝、操作步驟和功能介紹。
