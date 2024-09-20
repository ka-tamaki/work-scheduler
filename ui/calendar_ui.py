# ui/calendar_ui.py

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import calendar
from data.holidays import HolidaysManager
from ui.tooltips import CreateToolTip
import jpholiday

class CalendarUI(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("工場設定 - 休日設定")
        self.geometry("1200x800")  # サイズを拡大して12か月分表示
        self.parent = parent

        # 休日マネージャーの初期化
        self.manager = HolidaysManager()

        # スクロールバーの追加
        canvas = tk.Canvas(self, borderwidth=0)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 年の選択コンボボックス
        control_frame = tk.Frame(self.scrollable_frame)
        control_frame.pack(pady=10)

        tk.Label(control_frame, text="年:").pack(side=tk.LEFT, padx=5)
        current_year = datetime.now().year
        self.year_var = tk.IntVar(value=current_year)
        years = list(range(current_year - 5, 2031))  # 5年前から2030年まで
        self.year_cb = ttk.Combobox(control_frame, textvariable=self.year_var, values=years, width=5, state='readonly')
        self.year_cb.pack(side=tk.LEFT)
        self.year_cb.bind("<<ComboboxSelected>>", self.update_calendar)
        
        # カレンダーフレーム
        self.calendar_frame = tk.Frame(self.scrollable_frame)
        self.calendar_frame.pack()

        # カレンダーの初期表示
        self.generate_calendar()

    def generate_calendar(self):
        """指定された年の12か月分のカレンダーを生成"""
        # 既存のカレンダーをクリア
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()

        year = self.year_var.get()

        for month in range(1, 13):
            month_frame = tk.Frame(self.calendar_frame, borderwidth=1, relief='solid', padx=5, pady=5)
            month_frame.pack(pady=10, padx=10, fill='x')

            # 月のタイトルを表示
            month_title = f"{year}年{month}月"
            tk.Label(month_frame, text=month_title, font=("Arial", 14, "bold")).pack(pady=5)

            # 曜日ヘッダー
            days_frame = tk.Frame(month_frame)
            days_frame.pack()
            for day in ['月', '火', '水', '木', '金', '土', '日']:
                tk.Label(days_frame, text=day, width=4, borderwidth=1, relief='solid', bg='white').pack(side=tk.LEFT)

            # カレンダーの日付部分
            dates_frame = tk.Frame(month_frame)
            dates_frame.pack()

            cal = calendar.Calendar(firstweekday=0)  # 月曜日=0
            month_days = cal.monthdayscalendar(year, month)

            for week in month_days:
                week_frame = tk.Frame(dates_frame)
                week_frame.pack()

                for day in week:
                    if day == 0:
                        # 月に属さない日
                        tk.Label(week_frame, text='', width=4, height=2, borderwidth=1, relief='solid').pack(side=tk.LEFT)
                    else:
                        is_holiday = self.manager.is_holiday(year, month, day)
                        is_japanese_holiday = self.is_japanese_holiday(year, month, day)
                        if is_holiday:
                            bg_color = 'red'  # 休日・祝日は赤色
                            holiday_type = "休日"
                        else:
                            bg_color = 'white'
                            holiday_type = "平日"

                        btn = tk.Button(week_frame, text=str(day), width=4, height=2,
                                        bg=bg_color,
                                        relief='solid',
                                        bd=1,
                                        command=lambda y=year, m=month, d=day: self.toggle_holiday(y, m, d))
                        btn.pack(side=tk.LEFT, padx=1, pady=1)

                        # ツールチップの追加
                        tooltip_text = f"{year}/{month}/{day}\n{holiday_type}"
                        CreateToolTip(btn, tooltip_text)

    def is_japanese_holiday(self, year, month, day):
        """指定された日が日本の祝日かどうかを判定"""
        date_obj = datetime(year, month, day)
        return jpholiday.is_holiday(date_obj)

    def toggle_holiday(self, year, month, day):
        """休日をトグル（追加・削除）"""
        year_str = str(year)
        month_str = str(month)

        if year_str not in self.manager.holidays:
            self.manager.holidays[year_str] = {}
        if month_str not in self.manager.holidays[year_str]:
            self.manager.holidays[year_str][month_str] = []

        if day in self.manager.holidays[year_str][month_str]:
            # 休日から削除
            self.manager.remove_holiday(year, month, day)
            messagebox.showinfo("休日解除", f"{year}/{month}/{day} の休日を解除しました。")
        else:
            # 休日に追加
            self.manager.add_holiday(year, month, day)
            messagebox.showinfo("休日追加", f"{year}/{month}/{day} を休日に追加しました。")

        # カレンダーを再生成
        self.generate_calendar()

    def update_calendar(self, event=None):
        """年が変更されたときにカレンダーを更新する"""
        self.generate_calendar()
