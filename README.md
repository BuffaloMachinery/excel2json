# Excel 轉 JSON 工具

## 介紹
這是一個使用 Python 開發的圖形化應用程式，能將 Excel 檔案轉換為 JSON 格式。使用者可以選擇多個 Sheet 並將其各自轉換為單獨的 JSON 文件。

## 功能
- 支援選擇 Excel 文件（`.xlsx` 和 `.xls` 格式）
- 支援轉換多個 Sheet
- 將數據轉換為 JSON 格式並保存
- 簡單易用的圖形使用者介面（GUI）
- 顯示轉換進度及錯誤提示

## 安裝需求
在使用此工具前，請確保已安裝以下 Python 依賴包：

```bash
pip install pandas tqdm openpyxl

tkinter 是 Python 標準庫的一部分，通常無需額外安裝。

如何使用
運行腳本：執行該 Python 腳本。

bash
複製程式碼
python script.py
選擇 Excel 文件：點擊「選擇 Excel 文件」按鈕，選擇你想要轉換的 Excel 文件。

選擇 Sheet：在彈出的窗口中，選擇你想要轉換的 Sheet，允許多選。

選擇保存路徑：選擇 JSON 文件的保存路徑。每個 Sheet 會生成一個單獨的 JSON 文件，並以 Sheet 名稱作為文件名稱的一部分。

確認轉換：點擊「確認」後，程序會將選擇的 Sheet 轉換為 JSON 文件，並在完成後彈出提示。

轉換範例
假設你有一個名為 data.xlsx 的 Excel 文件，包含 Sheet1 和 Sheet2。如果你選擇這兩個 Sheet 進行轉換，程式會在指定路徑生成以下 JSON 文件：

data_Sheet1.json
data_Sheet2.json
JSON 文件結構範例
轉換後的 JSON 文件結構將根據 pandas 的 orient 參數生成。例如，使用默認的 records 格式時，JSON 文件將如下所示：

json
複製程式碼
[
    {
        "Column1": "value1",
        "Column2": "value2"
    },
    {
        "Column1": "value3",
        "Column2": "value4"
    }
]
日誌與錯誤處理
錯誤處理：如果在轉換過程中出現任何錯誤，會自動彈出錯誤提示，並將詳細的錯誤資訊記錄到日誌中。
日誌檔案：錯誤日誌將存儲在當前工作目錄下，以便用戶檢查轉換過程中可能發生的問題。
版本
1.0.0

未來功能計劃
增加更多的 JSON 格式選項，例如 table 或 index 格式。
支援將所有選定 Sheet 數據合併為一個 JSON 文件。
增加批次處理功能，允許轉換多個 Excel 文件。
貢獻
如果您對此專案有任何建議或改進，歡迎通過 GitHub 提交 Issue 或 Pull Request。

作者
此工具由 [Your Name] 開發，您可以通過 [your-email@example.com] 聯繫我。如有任何問題或改進建議，請隨時聯繫。