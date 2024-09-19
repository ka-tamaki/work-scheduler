# gui.py

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from excel_generator import generate_schedule

def create_gui():
    def on_generate():
        title = title_entry.get().strip()
        start_year = int(start_year_var.get())
        start_month = int(start_month_var.get())
        end_year = int(end_year_var.get())
        end_month = int(end_month_var.get())
    
        if not title:
            messagebox.showwarning("入力不足", "タイトルを入力してください。")
            return
    
        # 日付の整合性チェック
        try:
            start = datetime(start_year, start_month, 1)
            end = datetime(end_year, end_month, 1)
            if start > end:
                messagebox.showwarning("入力エラー", "開始年月が終了年月より後になっています。")
                return
        except ValueError as ve:
            messagebox.showwarning("入力エラー", f"日付の入力に誤りがあります。\n{ve}")
            return
    
        generate_schedule(title, start_year, start_month, end_year, end_month)
    
    # GUIの構築
    root = tk.Tk()
    root.title("製造工程表自動生成ツール")
    
    # タイトル入力
    tk.Label(root, text="タイトル:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
    title_entry = tk.Entry(root, width=30)
    title_entry.grid(row=0, column=1, padx=10, pady=5)
    
    # 開始年月の選択
    tk.Label(root, text="開始年月:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
    
    start_year_var = tk.StringVar()
    start_month_var = tk.StringVar()
    
    current_year = datetime.now().year
    years = [str(y) for y in range(current_year, current_year + 10)]
    months = [f"{m:02}" for m in range(1, 13)]
    
    start_year_cb = ttk.Combobox(root, textvariable=start_year_var, values=years, width=5, state='readonly')
    start_year_cb.set(str(current_year))
    start_year_cb.grid(row=1, column=1, padx=(10,0), pady=5, sticky='w')
    
    start_month_cb = ttk.Combobox(root, textvariable=start_month_var, values=months, width=3, state='readonly')
    start_month_cb.set(f"{datetime.now().month:02}")
    start_month_cb.grid(row=1, column=1, padx=(60,0), pady=5, sticky='w')
    
    # 終了年月の選択
    tk.Label(root, text="終了年月:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
    
    end_year_var = tk.StringVar()
    end_month_var = tk.StringVar()
    
    end_year_cb = ttk.Combobox(root, textvariable=end_year_var, values=years, width=5, state='readonly')
    end_year_cb.set(str(current_year + 1))
    end_year_cb.grid(row=2, column=1, padx=(10,0), pady=5, sticky='w')
    
    end_month_cb = ttk.Combobox(root, textvariable=end_month_var, values=months, width=3, state='readonly')
    end_month_cb.set("07")
    end_month_cb.grid(row=2, column=1, padx=(60,0), pady=5, sticky='w')
    
    # 生成ボタン
    generate_btn = tk.Button(root, text="生成", command=on_generate, width=20)
    generate_btn.grid(row=3, column=0, columnspan=2, padx=10, pady=20)
    
    return root
