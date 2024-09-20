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
        self.title("休日設定")
        self.geometry("1050x700")
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

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        self.scrollable_frame.bind_all("<MouseWheel>", _on_mousewheel)

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
        self.calendar_frame.pack(pady=10, padx=10)

        # ボタンを保持する辞書 {(year, month, day): button}
        self.button_refs = {}

        # カレンダーの初期表示
        self.generate_calendar()

    def generate_calendar(self):
        """指定された年の12か月分のカレンダーを4x3のグリッドで生成"""
        # 既存のカレンダーをクリア
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        self.button_refs.clear()

        year = self.year_var.get()

        # グリッドの設定
        columns = 4
        rows = 3
        for idx, month in enumerate(range(1, 13)):
            row = idx // columns
            col = idx % columns

            month_frame = tk.Frame(self.calendar_frame, borderwidth=1, relief='solid', width=250, height=300)
            month_frame.grid(row=row, column=col, sticky='nsew')
            month_frame.pack_propagate(False)  # サイズを固定

            # 月のタイトルを表示
            month_title = f"{month}月"
            title_label = tk.Label(month_frame, text=month_title, font=("Arial", 14, "bold"), bg='lightblue')
            title_label.pack(fill='x')

            # 曜日ヘッダー
            days_frame = tk.Frame(month_frame)
            days_frame.pack(fill='x')
            for day in ['月', '火', '水', '木', '金', '土', '日']:
                day_label = tk.Label(days_frame, text=day, width=4, borderwidth=1, relief='solid', bg='lightgray', font=("Arial", 10, "bold"))
                day_label.pack(side=tk.LEFT, fill='both', expand=True)

            # カレンダーの日付部分
            dates_frame = tk.Frame(month_frame)
            dates_frame.pack(fill='both', expand=True)

            cal = calendar.Calendar(firstweekday=0)  # 月曜日=0
            month_days = cal.monthdayscalendar(year, month)

            for week in month_days:
                week_frame = tk.Frame(dates_frame)
                week_frame.pack(fill='both', expand=True)
                week_frame.grid_columnconfigure(tuple(range(7)), weight=1)  # 各列に重みを付ける

                for day_idx, day in enumerate(week):
                    if day == 0:
                        # 月に属さない日も同じUIにする
                        label = tk.Label(week_frame, text=' ', width=4, height=2, borderwidth=1, relief='solid', bg='white', font=("Arial", 10))
                        label.grid(row=0, column=day_idx, sticky='nsew')
                    else:
                        is_holiday = self.manager.is_holiday(year, month, day)
                        if is_holiday:
                            bg_color = 'lightcoral'  # 休日・祝日は赤色系
                        else:
                            bg_color = 'white'

                        btn = tk.Button(week_frame, text=str(day), width=4, height=2,
                                        bg=bg_color,
                                        relief='solid',
                                        bd=1,
                                        font=("Arial", 10),
                                        command=lambda y=year, m=month, d=day: self.toggle_holiday(y, m, d))
                        btn.grid(row=0, column=day_idx, sticky='nsew')

                        # ボタン参照を保持
                        self.button_refs[(year, month, day)] = btn

        # カレンダーフレームの列と行に重みを付ける
        for col in range(columns):
            self.calendar_frame.grid_columnconfigure(col, weight=1)
        for row in range(rows):
            self.calendar_frame.grid_rowconfigure(row, weight=1)

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
            new_bg = 'white'
            holiday_type = "平日"
        else:
            # 休日に追加
            self.manager.add_holiday(year, month, day)
            new_bg = 'lightcoral'
            holiday_type = "休日"

        # ボタンの背景色を更新
        btn = self.button_refs.get((year, month, day))
        if btn:
            btn.configure(bg=new_bg)
            # ツールチップを更新
            CreateToolTip(btn, f"{year}/{month}/{day}\n{holiday_type}")

    def update_calendar(self, event=None):
        """年が変更されたときにカレンダーを更新する"""
        self.generate_calendar()
