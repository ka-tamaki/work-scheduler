# factory_settings.py

import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from datetime import datetime
import config

class FactorySettingsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("工場設定")
        self.geometry("1200x800")
        self.resizable(True, True)
        
        # 年の選択
        tk.Label(self, text="設定する年:", font=("Arial", 12)).pack(pady=10)
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
        self.save_btn = tk.Button(self, text="保存", command=self.save_holidays, font=("Arial", 12))
        self.save_btn.pack(pady=20)
        
        # 初期ロード
        self.load_holidays()
    
    def load_holidays(self, event=None):
        """指定された年の休日をカレンダーに反映"""
        selected_year = int(self.year_var.get())
        for month, cal in self.calendars.items():
            # 既存の休日をクリア
            cal.calevent_remove('all')
            
            holidays = config.HOLIDAYS.get(selected_year, {}).get(month, [])
            for day in holidays:
                try:
                    date = datetime(selected_year, month, day)
                    cal.calevent_create(date, '休日', 'holiday')
                except:
                    pass
            # カレンダーの再描画
            cal.tag_config('holiday', background='red', foreground='white')
    
    def toggle_holiday(self, event):
        """カレンダーの日付をクリックすると休日を追加・解除し、色を変更"""
        cal = event.widget
        selected_date = cal.selection_get()
        year = selected_date.year
        month = selected_date.month
        day = selected_date.day
        
        # HOLIDAYSに追加・解除
        if year in config.HOLIDAYS and month in config.HOLIDAYS[year]:
            if day in config.HOLIDAYS[year][month]:
                config.HOLIDAYS[year][month].remove(day)
                cal.calevent_remove('holiday', selected_date)
            else:
                config.HOLIDAYS[year][month].append(day)
                cal.calevent_create(selected_date, '休日', 'holiday')
        else:
            if year not in config.HOLIDAYS:
                config.HOLIDAYS[year] = {}
            if month not in config.HOLIDAYS[year]:
                config.HOLIDAYS[year][month] = []
            config.HOLIDAYS[year][month].append(day)
            cal.calevent_create(selected_date, '休日', 'holiday')
        
        # カレンダーの色設定
        cal.tag_config('holiday', background='red', foreground='white')
    
    def save_holidays(self):
        """休日情報を保存"""
        # ここでは単にメッセージを表示しますが、必要に応じてファイルに保存することも可能です。
        messagebox.showinfo("保存完了", f"{self.year_var.get()}年の休日が保存されました。")
        self.destroy()
