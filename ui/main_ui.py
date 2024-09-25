# ui/main_ui.py
import tkinter as tk
from tkinter import ttk, messagebox
from data.excel_generator import ExcelGenerator
from utils.path_helper import resource_path
from datetime import datetime, timedelta
import os
import logging

def get_template_path():
    path = resource_path(os.path.join('templates', 'template.xlsx'))
    return path

def get_output_path(title, start_date, end_date):
    output_dir = resource_path('output')
    os.makedirs(output_dir, exist_ok=True)
    filename = f"{title}_{start_date.strftime('%Y%m')}_{end_date.strftime('%Y%m')}.xlsx"
    full_path = os.path.join(output_dir, filename)
    return full_path

class MainUI:
    def __init__(self, root):
        self.root = root
        self.root.title("製造工程表自動生成ツール")

        # タイトル入力欄
        title_frame = tk.Frame(root)
        title_frame.pack(pady=10)

        tk.Label(title_frame, text="工事名:").pack(side=tk.LEFT, padx=5)
        self.title_var = tk.StringVar()
        self.title_entry = tk.Entry(title_frame, textvariable=self.title_var, width=30)
        self.title_entry.pack(side=tk.LEFT)

        # 期間選択
        selection_frame = tk.Frame(root)
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

        # 工場選択
        factory_frame = tk.Frame(root)
        factory_frame.pack(pady=10)
        tk.Label(factory_frame, text="工場選択:").pack(side=tk.LEFT, padx=5)
        self.factory_var = tk.StringVar()
        self.factories_mapping = {
            '結城': 'yuki',
            '熊谷': 'kumagaya',
            '静岡': 'shizuoka',
            '京都': 'kyoto',
            '千葉': 'chiba',
            '富山': 'toyama'
        }
        factories_display = list(self.factories_mapping.keys())
        self.factory_combo = ttk.Combobox(factory_frame, textvariable=self.factory_var, values=factories_display, width=15, state="readonly")
        self.factory_combo.pack(side=tk.LEFT)
        self.factory_combo.set(factories_display[0])

        # 生成ボタン
        generate_button = ttk.Button(root, text="工程表生成", command=self.generate_schedule)
        generate_button.pack(pady=20)

        # 初期値の設定
        self.set_initial_dates()

    def generate_year_month_options(self):
        """年/月のオプションを生成（例: '2024/01'）"""
        current_year = datetime.now().year
        start_year = current_year - 5
        end_year = current_year + 20
        months = range(1, 13)
        options = []
        for year in range(start_year, end_year + 1):
            for month in months:
                options.append(f"{year}/{month:02d}")
        return options

    def set_initial_dates(self):
        """開始年月と終了年月の初期値を設定"""
        now = datetime.now()
        start_date = now.strftime("%Y/%m")
        self.start_date_var.set(start_date)
        # 終了年月のオプションを生成
        self.end_date_cb['values'] = self.generate_end_date_options(now.year, now.month)
        # 初期値として半年後を設定
        end_date = self.calculate_end_date(now.year, now.month)
        self.end_date_var.set(end_date)
        self.end_date_cb.current(0)

    def calculate_end_date(self, year, month, delta_months=6):
        """指定された年と月からdelta_months後の年月を計算"""
        total_month = year * 12 + month + delta_months
        new_year = total_month // 12
        new_month = total_month % 12
        if new_month == 0:
            new_month = 12
            new_year -= 1
        return f"{new_year}/{new_month:02d}"

    def generate_end_date_options(self, start_year, start_month):
        """開始年月から半年後から20年後までの終了年月のオプションを生成"""
        options = []
        # 開始年月から20年後までの各月を追加
        for delta in range(6, 12 * 20 + 1):
            end_date = self.calculate_end_date(start_year, start_month, delta)
            if end_date not in options:
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
            # 範囲内か確認
            if new_end_date not in new_end_options:
                new_end_date = new_end_options[0] if new_end_options else ""
            self.end_date_var.set(new_end_date)
            self.end_date_cb.current(0)

    def generate_schedule(self):
        title = self.title_var.get().strip()
        if not title:
            messagebox.showerror("入力エラー", "工事名を入力してください。")
            return

        try:
            start_date_str = self.start_date_var.get()
            end_date_str = self.end_date_var.get()

            # 日付のパース
            start_year, start_month = map(int, start_date_str.split('/'))
            end_year, end_month = map(int, end_date_str.split('/'))
            start_date = datetime(start_year, start_month, 1)
            end_date = datetime(end_year, end_month, 1)

            if start_date > end_date:
                messagebox.showerror("入力エラー", "開始年月が終了年月より後になっています。")
                return
        except ValueError:
            messagebox.showerror("入力エラー", "有効な年月を選択してください。")
            return
        
        factory_display = self.factory_var.get()
        factory_internal = self.factories_mapping.get(factory_display)

        if not factory_internal:
            messagebox.showerror("選択エラー", "選択された工場の識別子が見つかりません。")
            return

        template_path = get_template_path()
        if not os.path.exists(template_path):
            logging.error(f"Template file not found: {template_path}")
            raise FileNotFoundError(f"Template file not found: {template_path}")

        output_path = get_output_path(title, start_date, end_date)

        try:
            generator = ExcelGenerator(template_path)
            generator.generate_excel(title, start_date, end_date, factory_internal, output_path)
        except Exception as e:
            messagebox.showerror("エラー", f"工程表の生成中にエラーが発生しました。\n{e}")
