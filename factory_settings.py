# factory_settings.py

import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from datetime import datetime, timedelta
import config

class FactorySettingsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("工場設定")
        self.geometry("400x400")
        self.resizable(False, False)
        
        # 年の選択
        tk.Label(self, text="設定する年:").pack(pady=10)
        self.year_var = tk.StringVar()
        current_year = datetime.now().year
        years = [str(y) for y in range(current_year, current_year + 10)]
        self.year_cb = ttk.Combobox(self, textvariable=self.year_var, values=years, width=10, state='readonly')
        self.year_cb.set(str(current_year))
        self.year_cb.pack(pady=5)
        
        # カレンダーの表示
        self.calendar = Calendar(self, selectmode='day', year=current_year, month=1, day=1, locale='ja_JP')
        self.calendar.pack(pady=20)
        
        # 既存の休日を設定
        self.load_holidays()
        
        # 休日選択ボタン
        self.select_btn = tk.Button(self, text="休日を選択/解除", command=self.toggle_holiday)
        self.select_btn.pack(pady=10)
        
        # 保存ボタン
        self.save_btn = tk.Button(self, text="保存", command=self.save_holidays)
        self.save_btn.pack(pady=10)
    
    def load_holidays(self):
        """既存の休日をカレンダーに反映"""
        selected_year = int(self.year_var.get())
        holidays = config.HOLIDAYS.get(selected_year, [])
        for day in holidays:
            try:
                self.calendar.calevent_create(datetime(selected_year, self.calendar.month, day), '休日', 'holiday')
                self.calendar.tag_config('holiday', background='red', foreground='white')
            except:
                pass
    
    def toggle_holiday(self):
        """選択された日を休日としてトグル"""
        selected_date = self.calendar.selection_get()
        selected_year = selected_date.year
        selected_day = selected_date.day
        
        # 既に休日に設定されているか確認
        holidays = config.HOLIDAYS.get(selected_year, [])
        if selected_day in holidays:
            # 休日から削除
            holidays.remove(selected_day)
            config.HOLIDAYS[selected_year] = holidays
            self.calendar.calevent_remove('holiday', datetime(selected_year, selected_date.month, selected_day))
        else:
            # 休日に追加
            holidays.append(selected_day)
            config.HOLIDAYS[selected_year] = holidays
            self.calendar.calevent_create(selected_date, '休日', 'holiday')
            self.calendar.tag_config('holiday', background='red', foreground='white')
    
    def save_holidays(self):
        """休日情報を保存"""
        selected_year = int(self.year_var.get())
        # Sort holidays
        holidays = sorted(config.HOLIDAYS.get(selected_year, []))
        config.HOLIDAYS[selected_year] = holidays
        messagebox.showinfo("保存完了", f"{selected_year}年の休日が保存されました。")
        self.destroy()
