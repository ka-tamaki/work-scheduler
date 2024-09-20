# factory_settings.py

import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from datetime import datetime
import config
import json
import os

class FactorySettingsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("工場設定")
        self.geometry("1300x900")
        self.resizable(True, True)
        
        # 休日情報のロード
        self.load_holidays_from_file()
        
        # 年の選択
        tk.Label(self, text="設定する年:", font=("Arial", 14)).pack(pady=10)
        self.year_var = tk.StringVar()
        current_year = datetime.now().year
        years = [str(y) for y in range(current_year - 1, current_year + 10)]
        self.year_cb = ttk.Combobox(self, textvariable=self.year_var, values=years, width=10, state='readonly', font=("Arial", 12))
        self.year_cb.set(str(current_year))
        self.year_cb.pack(pady=5)
        self.year_cb.bind("<<ComboboxSelected>>", self.load_holidays)
        
        # カレンダーのフレーム
        self.calendars_frame = tk.Frame(self)
        self.calendars_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        # 各月のカレンダーウィジェットを作成
        self.calendars = {}
        for month in range(1, 13):
            frame = tk.Frame(self.calendars_frame, borderwidth=1, relief="solid")
            frame.grid(row=(month-1)//4, column=(month-1)%4, padx=5, pady=5)
            cal = Calendar(frame, selectmode='day', year=current_year, month=month, day=1, locale='ja_JP', font=("Arial", 10))
            cal.pack(padx=10, pady=10)
            cal.bind("<<CalendarSelected>>", self.toggle_holiday)
            self.calendars[month] = cal
        
        # 保存ボタン
        self.save_btn = tk.Button(self, text="保存", command=self.save_holidays, font=("Arial", 14))
        self.save_btn.pack(pady=20)
    
    def load_holidays_from_file(self):
        """休日情報をJSONファイルからロード"""
        if os.path.exists(config.HOLIDAYS_FILE):
            with open(config.HOLIDAYS_FILE, 'r', encoding='utf-8') as f:
                config.HOLIDAYS = json.load(f)
        else:
            config.HOLIDAYS = {}
    
    def save_holidays_to_file(self):
        """休日情報をJSONファイルに保存"""
        with open(config.HOLIDAYS_FILE, 'w', encoding='utf-8') as f:
            json.dump(config.HOLIDAYS, f, ensure_ascii=False, indent=4)
    
    def load_holidays(self, event=None):
        """指定された年の休日をカレンダーに反映"""
        selected_year = int(self.year_var.get())
        for month, cal in self.calendars.items():
            # 既存の休日をクリア
            cal.calevent_remove('all')
            
            holidays = config.HOLIDAYS.get(str(selected_year), {}).get(str(month), [])
            for day in holidays:
                try:
                    date = datetime(selected_year, month, day)
                    cal.calevent_create(date, '休日', 'holiday')
                except:
                    pass
            # カレンダーの色設定
            cal.tag_config('holiday', background='red', foreground='white')
    
    def toggle_holiday(self, event):
        """カレンダーの日付をクリックすると休日を追加・解除し、色を変更"""
        cal = event.widget
        selected_date = cal.selection_get()
        year = selected_date.year
        month = selected_date.month
        day = selected_date.day
        
        # 文字列でキーを扱うために変換
        year_str = str(year)
        month_str = str(month)
        
        # HOLIDAYSに追加・解除
        if year_str not in config.HOLIDAYS:
            config.HOLIDAYS[year_str] = {}
        if month_str not in config.HOLIDAYS[year_str]:
            config.HOLIDAYS[year_str][month_str] = []
        
        if day in config.HOLIDAYS[year_str][month_str]:
            # 休日から削除
            config.HOLIDAYS[year_str][month_str].remove(day)
            cal.calevent_remove('holiday', selected_date)
        else:
            # 休日に追加
            config.HOLIDAYS[year_str][month_str].append(day)
            cal.calevent_create(selected_date, '休日', 'holiday')
        
        # カレンダーの色設定
        cal.tag_config('holiday', background='red', foreground='white')
    
    def save_holidays(self):
        """休日情報を保存"""
        self.save_holidays_to_file()
        messagebox.showinfo("保存完了", f"{self.year_var.get()}年の休日が保存されました。")
        self.destroy()
