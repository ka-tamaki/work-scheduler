# ui/main_ui.py

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
import os
from data.excel_generator import ExcelGenerator  # Excel生成機能
from ui.calendar_ui import CalendarUI
from config import EXCEL_OUTPUT_DIR  # 設定ファイルから出力ディレクトリをインポート

class MainUI(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.parent.title("工程表生成")
        self.pack(fill=tk.BOTH, expand=True)

        # メニューバーの作成
        menubar = tk.Menu(self.parent)
        self.parent.config(menu=menubar)

        # メニューの追加
        factory_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="メニュー", menu=factory_menu)
        factory_menu.add_command(label="休日設定", command=self.open_factory_settings)

        # タイトル入力欄
        title_frame = tk.Frame(self)
        title_frame.pack(pady=10)

        tk.Label(title_frame, text="工事名:").pack(side=tk.LEFT, padx=5)
        self.title_var = tk.StringVar()
        self.title_entry = tk.Entry(title_frame, textvariable=self.title_var, width=30)
        self.title_entry.pack(side=tk.LEFT)

        # 開始年月と終了年月の選択
        selection_frame = tk.Frame(self)
        selection_frame.pack(pady=20)

        # 開始年月の選択
        start_frame = tk.Frame(selection_frame)
        start_frame.pack(pady=5)
        tk.Label(start_frame, text="開始年月:").pack(side=tk.LEFT, padx=5)
        self.start_date_var = tk.StringVar()
        self.start_date_cb = ttk.Combobox(start_frame, textvariable=self.start_date_var, values=self.generate_year_month_options(), width=10, state='readonly')
        self.start_date_cb.pack(side=tk.LEFT)
        self.start_date_cb.bind("<<ComboboxSelected>>", self.update_end_date_options)

        # 終了年月の選択
        end_frame = tk.Frame(selection_frame)
        end_frame.pack(pady=5)
        tk.Label(end_frame, text="終了年月:").pack(side=tk.LEFT, padx=5)
        self.end_date_var = tk.StringVar()
        self.end_date_cb = ttk.Combobox(end_frame, textvariable=self.end_date_var, values=[], width=10, state='readonly')
        self.end_date_cb.pack(side=tk.LEFT)

        # 工程表生成ボタン
        generate_button = tk.Button(self, text="工程表生成", command=self.generate_schedule)
        generate_button.pack(pady=20)

        # Excel生成用の保存先ディレクトリ（設定ファイルから取得）
        self.save_dir = EXCEL_OUTPUT_DIR

        # 初期値の設定
        self.set_initial_dates()

    def generate_year_month_options(self):
        """年/月のオプションを生成（例: '2024/01'）"""
        current_year = datetime.now().year
        years = range(current_year - 5, 2031)
        months = range(1, 13)
        options = []
        for year in years:
            for month in months:
                options.append(f"{year}/{month:02d}")
        return options

    def set_initial_dates(self):
        """開始年月と終了年月の初期値を設定"""
        now = datetime.now()
        start_date = now.strftime("%Y/%m")
        self.start_date_var.set(start_date)
        # 計算後の終了年月を設定
        end_date = self.calculate_end_date(now.year, now.month)
        self.end_date_var.set(end_date)
        # 更新された終了年月のオプションに反映
        self.end_date_cb['values'] = self.generate_end_date_options(now.year, now.month)
        self.end_date_cb.current(0)

    def calculate_end_date(self, year, month, delta_months=6):
        """指定された年と月からdelta_months後の年月を計算"""
        month += delta_months
        year += (month - 1) // 12
        month = (month - 1) % 12 + 1
        return f"{year}/{month:02d}"

    def generate_end_date_options(self, start_year, start_month):
        """開始年月から半年後以降の終了年月のオプションを生成"""
        options = []
        for i in range(0, 12):  # 半年後から1年後までのオプション
            end_date = self.calculate_end_date(start_year, start_month, delta_months=6 + i)
            options.append(end_date)
        return options

    def update_end_date_options(self, event=None):
        """開始年月が変更されたときに終了年月のオプションを更新し、初期値を設定"""
        selected = self.start_date_var.get()
        if selected:
            year, month = map(int, selected.split('/'))
            # 新しい終了年月のオプションを生成
            new_end_options = self.generate_end_date_options(year, month)
            self.end_date_cb['values'] = new_end_options
            # 初期値として半年後を設定
            new_end_date = self.calculate_end_date(year, month)
            self.end_date_var.set(new_end_date)
            self.end_date_cb.current(0)

    def open_factory_settings(self):
        """工場設定（休日設定）画面を開く"""
        CalendarUI(self)

    def generate_schedule(self):
        """工程表を生成してExcelファイルを保存"""
        title = self.title_var.get().strip()
        if not title:
            messagebox.showerror("エラー", "タイトルを入力してください。")
            return

        start_date = self.start_date_var.get()
        end_date = self.end_date_var.get()

        if not start_date or not end_date:
            messagebox.showerror("エラー", "開始年月と終了年月を選択してください。")
            return

        # 年と月を分割
        start_year, start_month = map(int, start_date.split('/'))
        end_year, end_month = map(int, end_date.split('/'))

        # 開始年月と終了年月の整合性をチェック
        start = datetime(start_year, start_month, 1)
        end = datetime(end_year, end_month, 1)
        if end < start:
            messagebox.showerror("エラー", "終了年月が開始年月より前になっています。")
            return

        # Excel生成クラスを初期化
        generator = ExcelGenerator()

        # スケジュールを生成
        generator.generate_schedule(
            title,
            start_year,
            start_month,
            end_year,
            end_month,
            template_path=None,  # 設定ファイルからテンプレートパスを取得
            save_dir=self.save_dir
        )
