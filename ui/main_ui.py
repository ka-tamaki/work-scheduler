# ui/main_ui.py

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os
from data.excel_generator import ExcelGenerator  # Excel生成機能
from ui.calendar_ui import CalendarUI
from config import EXCEL_OUTPUT_DIR  # 設定ファイルから出力ディレクトリをインポート

class MainUI(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.parent.title("メイン画面 - スケジュール生成")
        self.pack(fill=tk.BOTH, expand=True)

        # メニューバーの作成
        menubar = tk.Menu(self.parent)
        self.parent.config(menu=menubar)

        # メニューの追加
        factory_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="メニュー", menu=factory_menu)
        factory_menu.add_command(label="工場設定", command=self.open_factory_settings)

        # 開始年月と終了年月の選択
        selection_frame = tk.Frame(self)
        selection_frame.pack(pady=20)

        # 開始年月
        start_frame = tk.Frame(selection_frame)
        start_frame.pack(pady=5)
        tk.Label(start_frame, text="開始年:").pack(side=tk.LEFT, padx=5)
        self.start_year_var = tk.IntVar(value=datetime.now().year)
        self.start_year_cb = ttk.Combobox(start_frame, textvariable=self.start_year_var, values=list(range(datetime.now().year - 5, 2031)), width=5, state='readonly')
        self.start_year_cb.pack(side=tk.LEFT)

        tk.Label(start_frame, text="開始月:").pack(side=tk.LEFT, padx=5)
        self.start_month_var = tk.IntVar(value=datetime.now().month)
        self.start_month_cb = ttk.Combobox(start_frame, textvariable=self.start_month_var, values=list(range(1, 13)), width=3, state='readonly')
        self.start_month_cb.pack(side=tk.LEFT)

        # 終了年月
        end_frame = tk.Frame(selection_frame)
        end_frame.pack(pady=5)
        tk.Label(end_frame, text="終了年:").pack(side=tk.LEFT, padx=5)
        self.end_year_var = tk.IntVar(value=datetime.now().year)
        self.end_year_cb = ttk.Combobox(end_frame, textvariable=self.end_year_var, values=list(range(datetime.now().year - 5, 2031)), width=5, state='readonly')
        self.end_year_cb.pack(side=tk.LEFT)

        tk.Label(end_frame, text="終了月:").pack(side=tk.LEFT, padx=5)
        self.end_month_var = tk.IntVar(value=datetime.now().month)
        self.end_month_cb = ttk.Combobox(end_frame, textvariable=self.end_month_var, values=list(range(1, 13)), width=3, state='readonly')
        self.end_month_cb.pack(side=tk.LEFT)

        # スケジュール生成ボタン
        generate_button = tk.Button(self, text="スケジュール生成", command=self.generate_schedule)
        generate_button.pack(pady=20)

        # Excel生成用の保存先ディレクトリ（設定ファイルから取得）
        self.save_dir = EXCEL_OUTPUT_DIR

    def open_factory_settings(self):
        """工場設定（休日設定）画面を開く"""
        CalendarUI(self)

    def generate_schedule(self):
        """スケジュールを生成してExcelファイルを保存"""
        title = "製造工程"  # タイトルは必要に応じて変更
        start_year = self.start_year_var.get()
        start_month = self.start_month_var.get()
        end_year = self.end_year_var.get()
        end_month = self.end_month_var.get()

        # 開始年月と終了年月の整合性をチェック
        if (end_year < start_year) or (end_year == start_year and end_month < start_month):
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
