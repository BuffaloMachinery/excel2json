import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, Listbox
from tqdm.auto import tqdm
import logging
import threading

def excel_to_json(file_path, sheet_name=None, indent=4, orient='records'):
    try:
        # 讀取 Excel 文件，假設以 1000 行為單位讀取
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        # 將 DataFrame 轉換為 JSON
        json_data = df.to_json(indent=indent, orient=orient, force_ascii=False)
        return json_data
    except Exception as e:
        logging.error(f"發生未知錯誤: {str(e)}")
        messagebox.showerror("錯誤", "發生未知錯誤，請檢查日誌")
        return None

def save_json(save_path, json_data):
    try:
        with open(save_path, 'w', encoding='utf-8') as f:
            f.write(json_data)
        messagebox.showinfo("成功", "JSON 文件已成功保存！")
    except Exception as e:
        logging.error(f"保存文件失敗: {str(e)}")
        messagebox.showerror("錯誤", f"保存文件失敗: {str(e)}")

def select_excel_file():
    file_path = filedialog.askopenfilename(
        title="選擇 Excel 文件",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        save_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")],
            title="保存 JSON 文件"
        )
        if save_path:
            # 獲取所有 Sheet 名稱
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names

            # 創建 Sheet 選擇窗口
            sheet_window = tk.Toplevel()
            sheet_window.title("選擇 Sheet")

            # 創建一個變量來存儲選定的 Sheet 名稱
            selected_sheets = []

            def on_select(event):
                widget = event.widget
                selection = widget.curselection()
                selected_sheets.clear()  # 清除之前選擇的 Sheet
                for index in selection:
                    selected_sheets.append(widget.get(index))

            sheet_listbox = Listbox(sheet_window, selectmode=tk.MULTIPLE)
            sheet_listbox.bind('<<ListboxSelect>>', on_select)
            for sheet_name in sheet_names:
                sheet_listbox.insert(tk.END, sheet_name)
            sheet_listbox.pack()

            confirm_button = tk.Button(sheet_window, text="確認", command=lambda: confirm_selection(selected_sheets, file_path, save_path, sheet_window))
            confirm_button.pack()

def confirm_selection(selected_sheets, file_path, save_path, sheet_window):
    all_sheets_data = {}
    for sheet_name in selected_sheets:
        json_data = excel_to_json(file_path, sheet_name)
        if json_data:
            all_sheets_data[sheet_name] = json_data
            # 保存每個 sheet 的 JSON 文件，可以根據需要組合成一個或多個文件
            sheet_save_path = save_path.replace('.json', f'_{sheet_name}.json')
            save_json(sheet_save_path, json_data)
    messagebox.showinfo("成功", "已成功轉換所選 Sheet")
    sheet_window.destroy()  # 關閉選擇窗口

# 創建主窗口
root = tk.Tk()
root.title("Excel 轉 JSON")

# 創建選擇文件按鈕
select_button = tk.Button(root, text="選擇 Excel 文件", command=select_excel_file)
select_button.pack(pady=20)

# 運行 GUI 主循環
root.mainloop()
